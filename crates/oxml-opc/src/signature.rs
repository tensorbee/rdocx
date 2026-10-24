//! OPC digital-signature discovery, verification, and coverage reporting.

use std::collections::{BTreeMap, BTreeSet, HashSet};
use std::time::{SystemTime, UNIX_EPOCH};

use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64;
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use ring::rand::SystemRandom;
use ring::signature::{
    KeyPair, RSA_PKCS1_2048_8192_SHA256, RSA_PKCS1_SHA256, RsaKeyPair, UnparsedPublicKey,
};
use sha2::{Digest, Sha256};
use x509_cert::Certificate;
use x509_cert::der::Decode;

use crate::error::{OpcError, Result};
use crate::package::OpcPackage;
use crate::relationship::{Relationship, Relationships, rel_types};

const DSIG_NS: &str = "http://www.w3.org/2000/09/xmldsig#";
const C14N_EXCLUSIVE: &str = "http://www.w3.org/2001/10/xml-exc-c14n#";
const DIGEST_SHA256: &str = "http://www.w3.org/2001/04/xmlenc#sha256";
const SIGNATURE_RSA_SHA256: &str = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
const RELATIONSHIP_TRANSFORM: &str =
    "http://schemas.openxmlformats.org/package/2006/RelationshipTransform";
const OPC_SIGNATURE_NS: &str = "http://schemas.openxmlformats.org/package/2006/digital-signature";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
const SIGNATURE_ORIGIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-origin";
const SIGNATURE_XML_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";

/// The signer identity embedded in an X.509 certificate.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SignerCertificateIdentity {
    pub subject: String,
    pub issuer: String,
    pub serial_number: String,
}

/// One relationship authenticated by an OPC relationship transform.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord)]
pub struct CoveredRelationship {
    /// `/` for package relationships, otherwise the owning part name.
    pub source_part: String,
    pub relationship_id: String,
}

/// A precise reason why a signature is invalid or incomplete.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum SignatureIssue {
    MissingPart {
        uri: String,
    },
    ChangedReference {
        uri: String,
        expected_digest: String,
        actual_digest: String,
    },
    DuplicateReference {
        uri: String,
    },
    MissingRelationship {
        source_part: String,
        relationship_id: String,
    },
    DuplicateRelationship {
        source_part: String,
        relationship_id: String,
    },
    ExternalRelationship {
        source_part: String,
        relationship_id: String,
        target: String,
    },
    MissingRelationshipTarget {
        source_part: String,
        relationship_id: String,
        target: String,
    },
    UncoveredPart {
        part_name: String,
    },
    UncoveredRelationship {
        source_part: String,
        relationship_id: String,
    },
    SignatureValueMismatch,
    OrphanSignaturePart {
        part_name: String,
    },
}

/// Verification and declared-coverage result for one signature part.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SignatureReport {
    pub signature_part: String,
    pub signer: Option<SignerCertificateIdentity>,
    /// True only when `SignedInfo`, its references, and authenticated manifest references verify.
    pub cryptographically_valid: bool,
    /// True only when every non-signature package part and relationship is covered.
    pub coverage_complete: bool,
    pub covered_parts: Vec<String>,
    pub covered_relationships: Vec<CoveredRelationship>,
    pub issues: Vec<SignatureIssue>,
}

#[derive(Debug, Clone)]
struct XmlName {
    qname: String,
    prefix: String,
    local: String,
    namespace: String,
}

#[derive(Debug, Clone)]
struct XmlAttribute {
    name: XmlName,
    value: String,
}

#[derive(Debug, Clone)]
enum XmlNode {
    Element(XmlElement),
    Text(String),
    ProcessingInstruction { target: String, content: String },
}

#[derive(Debug, Clone)]
struct XmlElement {
    name: XmlName,
    attributes: Vec<XmlAttribute>,
    children: Vec<XmlNode>,
    namespaces: BTreeMap<String, String>,
}

#[derive(Debug)]
enum ReferenceTransform {
    ExclusiveCanonicalization,
    Relationships(Vec<String>),
}

#[derive(Debug)]
struct ReferenceSpec {
    uri: String,
    digest: Vec<u8>,
    transforms: Vec<ReferenceTransform>,
}

// Bound both the original envelope and decoded allocation before parsing ASN.1.
const SIGNING_MATERIAL_LIMIT: usize = 1024 * 1024;

fn signing_material_der(
    input: &[u8],
    expected_label: &str,
) -> std::result::Result<Vec<u8>, String> {
    if input.is_empty() || input.len() > SIGNING_MATERIAL_LIMIT {
        return Err("signing material must contain between 1 byte and 1 MiB".into());
    }
    // DER is binary. Its bytes, including trailing bytes, reach the strict DER parser unchanged.
    let trimmed = input.trim_ascii();
    if !trimmed.starts_with(b"-----") {
        return Ok(input.to_vec());
    }
    if !input.is_ascii() {
        return Err("PEM framing must be ASCII".into());
    }
    let begin = format!("-----BEGIN {expected_label}-----");
    let end = format!("-----END {expected_label}-----");
    let body = trimmed
        .strip_prefix(begin.as_bytes())
        .and_then(|rest| rest.strip_suffix(end.as_bytes()))
        .ok_or_else(|| "expected exactly one matching unencrypted PEM block".to_owned())?;
    let encoded: Vec<u8> = body
        .iter()
        .copied()
        .filter(|byte| !byte.is_ascii_whitespace())
        .collect();
    if encoded.is_empty() {
        return Err("empty PEM body".into());
    }
    BASE64
        .decode(encoded)
        .map_err(|_| "invalid padded PEM base64".into())
}

pub(crate) fn create_signature(
    package: &mut OpcPackage,
    private_key_pkcs8_der: &[u8],
    certificate_der: &[u8],
) -> Result<SignatureReport> {
    package.validate_graph()?;
    let certificate_der = signing_material_der(certificate_der, "CERTIFICATE")
        .map_err(OpcError::InvalidSigningCertificate)?;
    let private_key_pkcs8_der = signing_material_der(private_key_pkcs8_der, "PRIVATE KEY")
        .map_err(OpcError::InvalidSigningKey)?;
    let certificate = Certificate::from_der(&certificate_der)
        .map_err(|error| OpcError::InvalidSigningCertificate(error.to_string()))?;
    let subject_public_key_info = certificate.tbs_certificate().subject_public_key_info();
    let key_algorithm = subject_public_key_info.algorithm.oid.to_string();
    if key_algorithm != "1.2.840.113549.1.1.1" {
        return Err(unsupported("key algorithm", &key_algorithm));
    }
    let certificate_public_key = subject_public_key_info
        .subject_public_key
        .as_bytes()
        .ok_or_else(|| {
            OpcError::InvalidSigningCertificate("non-octet-aligned public key".into())
        })?;
    let private_key = RsaKeyPair::from_pkcs8(&private_key_pkcs8_der)
        .map_err(|error| OpcError::InvalidSigningKey(error.to_string()))?;
    if private_key.public_key().as_ref() != certificate_public_key {
        return Err(OpcError::InvalidSigningKey(
            "private key does not match certificate".into(),
        ));
    }
    validate_unsigned_package_relationships(package)?;

    let origin_part = allocate_origin_part(package)?;
    let signature_part = allocate_numbered_part(package, "/_xmlsignatures", "sig", "xml")?;
    package.set_part(&origin_part, Vec::new());
    package
        .content_types
        .add_override(&origin_part, SIGNATURE_ORIGIN_CONTENT_TYPE);
    package
        .content_types
        .add_override(&signature_part, SIGNATURE_XML_CONTENT_TYPE);
    package.package_rels.add(
        rel_types::DIGITAL_SIGNATURE_ORIGIN,
        origin_part.trim_start_matches('/'),
    );
    let signature_target = signature_part
        .strip_prefix("/_xmlsignatures/")
        .ok_or_else(|| {
            OpcError::SignatureCreationFailed("signature part is outside its origin".into())
        })?;
    package
        .get_or_create_part_rels(&origin_part)
        .add(rel_types::DIGITAL_SIGNATURE, signature_target);

    let manifest = create_manifest(package)?;
    let signature_time = current_signature_time()?;
    let object = format!(
        "<ds:Object xmlns:ds=\"{DSIG_NS}\" xmlns:mdssi=\"{OPC_SIGNATURE_NS}\" Id=\"idPackageObject\"><ds:Manifest>{manifest}</ds:Manifest><ds:SignatureProperties><ds:SignatureProperty Id=\"idSignatureTime\" Target=\"#idPackageSignature\"><mdssi:SignatureTime><mdssi:Format>YYYY-MM-DDThh:mm:ssTZD</mdssi:Format><mdssi:Value>{signature_time}</mdssi:Value></mdssi:SignatureTime></ds:SignatureProperty></ds:SignatureProperties></ds:Object>"
    );
    let object_element = parse_xml(object.as_bytes())?;
    let object_digest = BASE64.encode(Sha256::digest(canonicalize(&object_element)));
    let signed_info = format!(
        "<ds:SignedInfo xmlns:ds=\"{DSIG_NS}\"><ds:CanonicalizationMethod Algorithm=\"{C14N_EXCLUSIVE}\"></ds:CanonicalizationMethod><ds:SignatureMethod Algorithm=\"{SIGNATURE_RSA_SHA256}\"></ds:SignatureMethod><ds:Reference URI=\"#idPackageObject\"><ds:Transforms><ds:Transform Algorithm=\"{C14N_EXCLUSIVE}\"></ds:Transform></ds:Transforms><ds:DigestMethod Algorithm=\"{DIGEST_SHA256}\"></ds:DigestMethod><ds:DigestValue>{object_digest}</ds:DigestValue></ds:Reference></ds:SignedInfo>"
    );
    let signed_info_element = parse_xml(signed_info.as_bytes())?;
    let signed_info_canonical = canonicalize(&signed_info_element);
    let mut signature = vec![0_u8; private_key.public().modulus_len()];
    private_key
        .sign(
            &RSA_PKCS1_SHA256,
            &SystemRandom::new(),
            &signed_info_canonical,
            &mut signature,
        )
        .map_err(|_| OpcError::SignatureCreationFailed("RSA-SHA256 signing failed".into()))?;
    let signature_value = BASE64.encode(signature);
    let certificate_value = BASE64.encode(certificate_der);
    let signature_xml = format!(
        "<?xml version=\"1.0\"?><ds:Signature xmlns:ds=\"{DSIG_NS}\" xmlns:mdssi=\"{OPC_SIGNATURE_NS}\" Id=\"idPackageSignature\">{signed_info}<ds:SignatureValue>{signature_value}</ds:SignatureValue><ds:KeyInfo><ds:X509Data><ds:X509Certificate>{certificate_value}</ds:X509Certificate></ds:X509Data></ds:KeyInfo>{object}</ds:Signature>"
    );
    package.set_part(&signature_part, signature_xml.into_bytes());

    let reports = verify_signatures(package)?;
    if let Some(invalid) = reports
        .iter()
        .find(|report| !report.cryptographically_valid || !report.coverage_complete)
    {
        return Err(OpcError::SignatureCreationFailed(format!(
            "staged verification rejected {} with {:?}",
            invalid.signature_part, invalid.issues
        )));
    }
    let report = reports
        .into_iter()
        .find(|report| report.signature_part == signature_part)
        .ok_or_else(|| {
            OpcError::SignatureCreationFailed("staged signature was not discoverable".into())
        })?;
    Ok(report)
}

fn create_manifest(package: &OpcPackage) -> Result<String> {
    let infrastructure = signature_infrastructure_parts(package);
    let mut references = String::new();
    append_reference(
        &mut references,
        "/%5BContent_Types%5D.xml",
        &content_types_bytes(package)?,
        "",
    );

    let mut part_names: Vec<&String> = package
        .parts
        .keys()
        .filter(|part_name| !infrastructure.contains(&crate::package::part_identity(part_name)))
        .collect();
    part_names.sort();
    for part_name in part_names {
        let content_type = package
            .content_types
            .content_type_for(part_name)
            .ok_or_else(|| {
                OpcError::SignatureCreationFailed(format!(
                    "signed part has no content type: {part_name}"
                ))
            })?;
        append_reference(
            &mut references,
            &reference_uri_with_content_type(part_name, content_type),
            &package.parts[part_name],
            "",
        );
    }

    append_relationship_reference(&mut references, package, "/", &package.package_rels)?;
    let mut sources: Vec<&String> = package.part_rels.keys().collect();
    sources.sort();
    for source in sources {
        validate_relationship_source(package, source)?;
        append_relationship_reference(
            &mut references,
            package,
            source,
            &package.part_rels[source],
        )?;
    }
    Ok(references)
}

fn validate_relationship_source(package: &OpcPackage, source: &str) -> Result<()> {
    let normalized = source.starts_with('/')
        && source.len() > 1
        && !source.ends_with('/')
        && !source.contains("//")
        && !source
            .split('/')
            .any(|segment| matches!(segment, "." | ".."));
    if normalized && package.contains_part(source) {
        return Ok(());
    }
    Err(OpcError::SignatureCreationFailed(format!(
        "relationship source is not an existing normalized package part: {source:?}"
    )))
}

fn validate_unsigned_package_relationships(package: &OpcPackage) -> Result<()> {
    validate_unsigned_relationship_set(package, "/", &package.package_rels)?;
    let mut sources: Vec<&String> = package.part_rels.keys().collect();
    sources.sort();
    for source in sources {
        validate_relationship_source(package, source)?;
        validate_unsigned_relationship_set(package, source, &package.part_rels[source])?;
    }
    Ok(())
}

fn validate_unsigned_relationship_set(
    package: &OpcPackage,
    source: &str,
    relationships: &Relationships,
) -> Result<()> {
    if relationships.items.iter().any(|relationship| {
        relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN
            || relationship.rel_type == rel_types::DIGITAL_SIGNATURE
    }) {
        return Err(OpcError::SignatureCreationFailed(format!(
            "unsigned package contains a digital-signature relationship from {source}"
        )));
    }
    let ids: Vec<String> = relationships
        .items
        .iter()
        .map(|relationship| relationship.id.clone())
        .collect();
    let mut covered = BTreeSet::new();
    let mut issues = Vec::new();
    relationship_transform(
        package,
        source,
        relationships,
        &ids,
        &mut covered,
        &mut issues,
    )
    .ok_or_else(|| {
        OpcError::SignatureCreationFailed(format!(
            "invalid relationship graph before signing: {issues:?}"
        ))
    })?;
    Ok(())
}

fn current_signature_time() -> Result<String> {
    let seconds = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map_err(|error| OpcError::SignatureCreationFailed(error.to_string()))?
        .as_secs();
    let days = i64::try_from(seconds / 86_400)
        .map_err(|error| OpcError::SignatureCreationFailed(error.to_string()))?;
    let seconds_of_day = seconds % 86_400;
    let (year, month, day) = civil_date_from_unix_days(days);
    let hour = seconds_of_day / 3_600;
    let minute = seconds_of_day % 3_600 / 60;
    let second = seconds_of_day % 60;
    Ok(format!(
        "{year:04}-{month:02}-{day:02}T{hour:02}:{minute:02}:{second:02}Z"
    ))
}

fn civil_date_from_unix_days(days: i64) -> (i64, i64, i64) {
    let shifted = days + 719_468;
    let era = shifted.div_euclid(146_097);
    let day_of_era = shifted.rem_euclid(146_097);
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let mut year = year_of_era + era * 400;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_prime = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_prime + 2) / 5 + 1;
    let month = month_prime + if month_prime < 10 { 3 } else { -9 };
    year += i64::from(month <= 2);
    (year, month, day)
}

fn append_relationship_reference(
    output: &mut String,
    package: &OpcPackage,
    source: &str,
    relationships: &Relationships,
) -> Result<()> {
    let mut ids: Vec<String> = relationships
        .items
        .iter()
        .filter(|relationship| {
            relationship.rel_type != rel_types::DIGITAL_SIGNATURE_ORIGIN
                && relationship.rel_type != rel_types::DIGITAL_SIGNATURE
        })
        .map(|relationship| relationship.id.clone())
        .collect();
    ids.sort();
    if ids.is_empty() {
        return Ok(());
    }

    let mut covered = BTreeSet::new();
    let mut issues = Vec::new();
    let transformed = relationship_transform(
        package,
        source,
        relationships,
        &ids,
        &mut covered,
        &mut issues,
    )
    .ok_or_else(|| {
        OpcError::SignatureCreationFailed(format!(
            "cannot authenticate relationships for {source}: {issues:?}"
        ))
    })?;
    let uri = if source == "/" {
        "/_rels/.rels".to_string()
    } else {
        relationship_part_name(source)
    };
    let uri = reference_uri_with_content_type(&uri, crate::content_types::RELATIONSHIPS);
    let mut transforms =
        format!("<ds:Transforms><ds:Transform Algorithm=\"{RELATIONSHIP_TRANSFORM}\">");
    for id in ids {
        transforms.push_str("<mdssi:RelationshipReference SourceId=\"");
        escape_attribute(&id, &mut transforms);
        transforms.push_str("\"></mdssi:RelationshipReference>");
    }
    transforms.push_str(&format!(
        "</ds:Transform><ds:Transform Algorithm=\"{C14N_EXCLUSIVE}\"></ds:Transform></ds:Transforms>"
    ));
    append_reference(output, &uri, &transformed, &transforms);
    Ok(())
}

fn append_reference(output: &mut String, uri: &str, bytes: &[u8], transforms: &str) {
    output.push_str("<ds:Reference URI=\"");
    escape_attribute(uri, output);
    output.push_str("\">");
    output.push_str(transforms);
    output.push_str("<ds:DigestMethod Algorithm=\"");
    output.push_str(DIGEST_SHA256);
    output.push_str("\"></ds:DigestMethod><ds:DigestValue>");
    output.push_str(&BASE64.encode(Sha256::digest(bytes)));
    output.push_str("</ds:DigestValue></ds:Reference>");
}

fn relationship_part_name(source: &str) -> String {
    let split = source.rfind('/').unwrap_or(0);
    let directory = &source[..=split];
    let filename = &source[split + 1..];
    format!("{directory}_rels/{filename}.rels")
}

fn percent_encode_uri(uri: &str) -> String {
    let mut encoded = String::with_capacity(uri.len());
    for byte in uri.bytes() {
        if byte.is_ascii_alphanumeric() || matches!(byte, b'/' | b'-' | b'.' | b'_' | b'~') {
            encoded.push(char::from(byte));
        } else {
            encoded.push('%');
            encoded.push(char::from(b"0123456789ABCDEF"[(byte >> 4) as usize]));
            encoded.push(char::from(b"0123456789ABCDEF"[(byte & 0x0f) as usize]));
        }
    }
    encoded
}

fn reference_uri_with_content_type(part_name: &str, content_type: &str) -> String {
    format!(
        "{}?ContentType={content_type}",
        percent_encode_uri(part_name)
    )
}

fn allocate_origin_part(package: &OpcPackage) -> Result<String> {
    let conventional = "/_xmlsignatures/origin.sigs";
    if !part_name_occupied(package, conventional) {
        return Ok(conventional.to_string());
    }
    allocate_numbered_part(package, "/_xmlsignatures", "origin", "sigs")
}

fn allocate_numbered_part(
    package: &OpcPackage,
    directory: &str,
    stem: &str,
    extension: &str,
) -> Result<String> {
    let prefix = format!("{directory}/{stem}");
    let suffix = format!(".{extension}");
    let mut maximum = 0_usize;
    for part_name in package
        .parts
        .keys()
        .chain(package.part_rels.keys())
        .chain(package.content_types.overrides.keys())
    {
        let lowercase = part_name.to_ascii_lowercase();
        let Some(number) = lowercase
            .strip_prefix(&prefix)
            .and_then(|value| value.strip_suffix(&suffix))
            .and_then(|value| value.parse::<usize>().ok())
            .filter(|number| *number > 0)
        else {
            continue;
        };
        maximum = maximum.max(number);
    }
    let mut number = maximum.checked_add(1).unwrap_or(1);
    let start = number;
    loop {
        let candidate = format!("{prefix}{number}{suffix}");
        if !part_name_occupied(package, &candidate) {
            return Ok(candidate);
        }
        number = number.checked_add(1).unwrap_or(1);
        if number == start {
            return Err(OpcError::SignatureCreationFailed(format!(
                "no free {stem} part name"
            )));
        }
    }
}

fn part_name_occupied(package: &OpcPackage, name: &str) -> bool {
    package.contains_part(name)
        || package.get_part_rels(name).is_some()
        || package.content_types.contains_override(name)
}

pub(crate) fn verify_signatures(package: &OpcPackage) -> Result<Vec<SignatureReport>> {
    let mut signature_parts = discover_signature_parts(package)?;
    let discovered: HashSet<String> = signature_parts
        .iter()
        .map(|name| crate::package::part_identity(name))
        .collect();
    let mut orphaned: Vec<String> = package
        .parts
        .keys()
        .filter(|name| {
            is_signature_xml_part(name)
                && !discovered.contains(&crate::package::part_identity(name))
        })
        .cloned()
        .collect();
    orphaned.sort();
    signature_parts.sort();

    let mut reports = Vec::with_capacity(signature_parts.len() + orphaned.len());
    for part_name in signature_parts {
        let xml = package.get_part(&part_name).ok_or_else(|| {
            OpcError::InvalidSignatureXml(format!("missing signature part {part_name}"))
        })?;
        reports.push(verify_signature_part(package, &part_name, xml)?);
    }
    reports.extend(orphaned.into_iter().map(|part_name| SignatureReport {
        signature_part: part_name.clone(),
        signer: None,
        cryptographically_valid: false,
        coverage_complete: false,
        covered_parts: Vec::new(),
        covered_relationships: Vec::new(),
        issues: vec![SignatureIssue::OrphanSignaturePart { part_name }],
    }));
    Ok(reports)
}

fn discover_signature_parts(package: &OpcPackage) -> Result<Vec<String>> {
    let origins = package
        .package_rels
        .get_all_by_type(rel_types::DIGITAL_SIGNATURE_ORIGIN);
    let mut parts: Vec<String> = Vec::new();
    for origin in origins {
        if is_external(origin) {
            return Err(OpcError::InvalidSignatureXml(format!(
                "external signature origin relationship {}",
                origin.id
            )));
        }
        let origin_part = OpcPackage::resolve_rel_target("/", &origin.target);
        let Some(stored_origin) = package.stored_part_name(&origin_part) else {
            return Err(OpcError::InvalidSignatureXml(format!(
                "signature origin target is missing: {origin_part}"
            )));
        };
        let origin_part = stored_origin.to_owned();
        let origin_rels = package.get_part_rels(&origin_part).ok_or_else(|| {
            OpcError::InvalidSignatureXml(format!(
                "signature origin has no relationships: {origin_part}"
            ))
        })?;
        for signature in origin_rels.get_all_by_type(rel_types::DIGITAL_SIGNATURE) {
            if is_external(signature) {
                return Err(OpcError::InvalidSignatureXml(format!(
                    "external signature relationship {}",
                    signature.id
                )));
            }
            let target = OpcPackage::resolve_rel_target(&origin_part, &signature.target);
            let part_name = package
                .stored_part_name(&target)
                .unwrap_or(&target)
                .to_owned();
            if parts.iter().any(|existing| {
                crate::package::part_identity(existing) == crate::package::part_identity(&part_name)
            }) {
                return Err(OpcError::InvalidSignatureXml(format!(
                    "duplicate signature target {part_name}"
                )));
            }
            parts.push(part_name);
        }
    }
    Ok(parts)
}

fn verify_signature_part(
    package: &OpcPackage,
    signature_part: &str,
    xml: &[u8],
) -> Result<SignatureReport> {
    let root = parse_xml(xml)?;
    require_name(&root, DSIG_NS, "Signature")?;
    let signed_info = direct_child(&root, DSIG_NS, "SignedInfo")?;
    require_algorithm(
        direct_child(signed_info, DSIG_NS, "CanonicalizationMethod")?,
        "canonicalization",
        C14N_EXCLUSIVE,
    )?;
    require_algorithm(
        direct_child(signed_info, DSIG_NS, "SignatureMethod")?,
        "signature method",
        SIGNATURE_RSA_SHA256,
    )?;

    let certificate_element = descendant(&root, DSIG_NS, "X509Certificate")
        .ok_or_else(|| invalid_signature("missing X509Certificate"))?;
    let certificate_der = decode_base64(&element_text(certificate_element), "X509Certificate")?;
    let certificate = Certificate::from_der(&certificate_der)
        .map_err(|error| OpcError::InvalidSigningCertificate(error.to_string()))?;
    let subject_public_key_info = certificate.tbs_certificate().subject_public_key_info();
    let key_algorithm = subject_public_key_info.algorithm.oid.to_string();
    if key_algorithm != "1.2.840.113549.1.1.1" {
        return Err(unsupported("key algorithm", &key_algorithm));
    }
    let public_key = subject_public_key_info
        .subject_public_key
        .as_bytes()
        .ok_or_else(|| {
            OpcError::InvalidSigningCertificate("non-octet-aligned public key".into())
        })?;
    let signer = SignerCertificateIdentity {
        subject: certificate.tbs_certificate().subject().to_string(),
        issuer: certificate.tbs_certificate().issuer().to_string(),
        serial_number: certificate.tbs_certificate().serial_number().to_string(),
    };

    let mut issues = Vec::new();
    let mut covered_parts = BTreeSet::new();
    let mut covered_relationships = BTreeSet::new();
    let mut seen_references = HashSet::new();
    let mut all_reference_digests_valid = true;
    let mut authenticated_roots = Vec::new();

    for reference_element in direct_children(signed_info, DSIG_NS, "Reference") {
        let reference = parse_reference(reference_element)?;
        if !seen_references.insert(reference.uri.clone()) {
            issues.push(SignatureIssue::DuplicateReference {
                uri: reference.uri.clone(),
            });
            all_reference_digests_valid = false;
            continue;
        }
        let actual = apply_reference(
            package,
            &root,
            &reference,
            &mut covered_parts,
            &mut covered_relationships,
            &mut issues,
        )?;
        let digest_valid = compare_digest(&reference, actual, &mut issues);
        all_reference_digests_valid &= digest_valid;
        if digest_valid && let Some(target) = same_document_target(&root, &reference.uri)? {
            authenticated_roots.push(target);
        }
    }

    let mut manifests = Vec::new();
    for target in authenticated_roots {
        collect_matching_elements(target, DSIG_NS, "Manifest", &mut manifests);
    }
    let mut processed_manifests = HashSet::new();
    let mut manifest_index = 0;
    while let Some(manifest) = manifests.get(manifest_index).copied() {
        manifest_index += 1;
        if !processed_manifests.insert(manifest as *const XmlElement) {
            continue;
        }
        for reference_element in direct_children(manifest, DSIG_NS, "Reference") {
            let reference = parse_reference(reference_element)?;
            if !seen_references.insert(reference.uri.clone()) {
                issues.push(SignatureIssue::DuplicateReference {
                    uri: reference.uri.clone(),
                });
                all_reference_digests_valid = false;
                continue;
            }
            let actual = apply_reference(
                package,
                &root,
                &reference,
                &mut covered_parts,
                &mut covered_relationships,
                &mut issues,
            )?;
            let digest_valid = compare_digest(&reference, actual, &mut issues);
            all_reference_digests_valid &= digest_valid;
            if digest_valid && let Some(target) = same_document_target(&root, &reference.uri)? {
                collect_matching_elements(target, DSIG_NS, "Manifest", &mut manifests);
            }
        }
    }

    let signed_info_bytes = canonicalize(signed_info);
    let signature_value = decode_base64(
        &element_text(direct_child(&root, DSIG_NS, "SignatureValue")?),
        "SignatureValue",
    )?;
    let signed_info_valid = UnparsedPublicKey::new(&RSA_PKCS1_2048_8192_SHA256, public_key)
        .verify(&signed_info_bytes, &signature_value)
        .is_ok();
    if !signed_info_valid {
        issues.push(SignatureIssue::SignatureValueMismatch);
    }

    add_uncovered_issues(package, &covered_parts, &covered_relationships, &mut issues);
    let coverage_complete = !issues.iter().any(is_coverage_issue);
    let cryptographically_valid = signed_info_valid && all_reference_digests_valid;

    Ok(SignatureReport {
        signature_part: signature_part.to_string(),
        signer: Some(signer),
        cryptographically_valid,
        coverage_complete,
        covered_parts: covered_parts.into_iter().collect(),
        covered_relationships: covered_relationships.into_iter().collect(),
        issues,
    })
}

fn parse_reference(element: &XmlElement) -> Result<ReferenceSpec> {
    let uri = attribute(element, "", "URI")
        .ok_or_else(|| invalid_signature("Reference has no URI"))?
        .to_string();
    let digest_method = direct_child(element, DSIG_NS, "DigestMethod")?;
    require_algorithm(digest_method, "digest method", DIGEST_SHA256)?;
    let digest = decode_base64(
        &element_text(direct_child(element, DSIG_NS, "DigestValue")?),
        "DigestValue",
    )?;
    let mut transforms = Vec::new();
    if let Some(container) = optional_direct_child(element, DSIG_NS, "Transforms")? {
        for transform in direct_children(container, DSIG_NS, "Transform") {
            let algorithm = attribute(transform, "", "Algorithm")
                .ok_or_else(|| invalid_signature("Transform has no Algorithm"))?;
            match algorithm {
                C14N_EXCLUSIVE => transforms.push(ReferenceTransform::ExclusiveCanonicalization),
                RELATIONSHIP_TRANSFORM => {
                    let mut ids = Vec::new();
                    for reference in element_children(transform) {
                        if !has_name(reference, OPC_SIGNATURE_NS, "RelationshipReference") {
                            return Err(unsupported(
                                "relationship transform child",
                                &reference.name.qname,
                            ));
                        }
                        let id = attribute(reference, "", "SourceId").ok_or_else(|| {
                            invalid_signature("RelationshipReference has no SourceId")
                        })?;
                        if ids.iter().any(|existing| existing == id) {
                            return Err(invalid_signature(&format!(
                                "duplicate relationship SourceId {id}"
                            )));
                        }
                        ids.push(id.to_string());
                    }
                    transforms.push(ReferenceTransform::Relationships(ids));
                }
                other => return Err(unsupported("transform", other)),
            }
        }
    }
    Ok(ReferenceSpec {
        uri,
        digest,
        transforms,
    })
}

fn apply_reference(
    package: &OpcPackage,
    root: &XmlElement,
    reference: &ReferenceSpec,
    covered_parts: &mut BTreeSet<String>,
    covered_relationships: &mut BTreeSet<CoveredRelationship>,
    issues: &mut Vec<SignatureIssue>,
) -> Result<Option<Vec<u8>>> {
    if let Some(element) = same_document_target(root, &reference.uri)? {
        if !matches!(
            reference.transforms.as_slice(),
            [ReferenceTransform::ExclusiveCanonicalization]
        ) {
            return Err(unsupported("same-document transform chain", &reference.uri));
        }
        return Ok(Some(canonicalize(element)));
    }

    let part_name = reference_part_name(&reference.uri)?;
    if let Some((source_part, relationships)) = relationships_for_uri(package, &part_name) {
        let ids = match reference.transforms.as_slice() {
            [
                ReferenceTransform::Relationships(ids),
                ReferenceTransform::ExclusiveCanonicalization,
            ]
            | [ReferenceTransform::Relationships(ids)] => ids,
            _ => return Err(unsupported("relationship transform chain", &reference.uri)),
        };
        let transformed = relationship_transform(
            package,
            &source_part,
            relationships,
            ids,
            covered_relationships,
            issues,
        );
        return Ok(transformed);
    }

    if !reference.transforms.is_empty() {
        return Err(unsupported("part transform chain", &reference.uri));
    }
    let is_content_types = part_name.eq_ignore_ascii_case("/[Content_Types].xml");
    let bytes = if is_content_types {
        Some(content_types_bytes(package)?)
    } else {
        package.get_part(&part_name).map(<[u8]>::to_vec)
    };
    match bytes {
        Some(bytes) => {
            let covered_name = if is_content_types {
                "/[Content_Types].xml".to_owned()
            } else {
                package
                    .stored_part_name(&part_name)
                    .unwrap_or(&part_name)
                    .to_owned()
            };
            covered_parts.insert(covered_name);
            Ok(Some(bytes))
        }
        None => {
            issues.push(SignatureIssue::MissingPart {
                uri: reference.uri.clone(),
            });
            Ok(None)
        }
    }
}

fn same_document_target<'a>(root: &'a XmlElement, uri: &str) -> Result<Option<&'a XmlElement>> {
    let Some(id) = uri.strip_prefix('#') else {
        return Ok(None);
    };
    let elements = find_all_by_id(root, id);
    match elements.as_slice() {
        [] => Err(invalid_signature(&format!(
            "missing same-document reference #{id}"
        ))),
        [element] => Ok(Some(*element)),
        _ => Err(invalid_signature(&format!(
            "duplicate same-document Id {id}"
        ))),
    }
}

fn relationship_transform(
    package: &OpcPackage,
    source_part: &str,
    relationships: &Relationships,
    ids: &[String],
    covered: &mut BTreeSet<CoveredRelationship>,
    issues: &mut Vec<SignatureIssue>,
) -> Option<Vec<u8>> {
    let mut selected = Vec::new();
    let mut valid = true;
    for id in ids {
        let matches: Vec<&Relationship> = relationships
            .items
            .iter()
            .filter(|relationship| relationship.id == *id)
            .collect();
        match matches.as_slice() {
            [] => {
                issues.push(SignatureIssue::MissingRelationship {
                    source_part: source_part.to_string(),
                    relationship_id: id.clone(),
                });
                valid = false;
            }
            [relationship] if is_external(relationship) => {
                issues.push(SignatureIssue::ExternalRelationship {
                    source_part: source_part.to_string(),
                    relationship_id: id.clone(),
                    target: relationship.target.clone(),
                });
                valid = false;
            }
            [relationship] => {
                let target = OpcPackage::resolve_rel_target(source_part, &relationship.target);
                if !package.contains_part(&target) {
                    issues.push(SignatureIssue::MissingRelationshipTarget {
                        source_part: source_part.to_string(),
                        relationship_id: id.clone(),
                        target,
                    });
                    valid = false;
                    continue;
                }
                selected.push(*relationship);
                covered.insert(CoveredRelationship {
                    source_part: source_part.to_string(),
                    relationship_id: id.clone(),
                });
            }
            _ => {
                issues.push(SignatureIssue::DuplicateRelationship {
                    source_part: source_part.to_string(),
                    relationship_id: id.clone(),
                });
                valid = false;
            }
        }
    }
    if !valid {
        return None;
    }
    selected.sort_by(|left, right| left.id.cmp(&right.id));
    let mut output = format!("<Relationships xmlns=\"{RELATIONSHIPS_NS}\">");
    for relationship in selected {
        output.push_str("<Relationship Id=\"");
        escape_attribute(&relationship.id, &mut output);
        output.push_str("\" Target=\"");
        escape_attribute(&relationship.target, &mut output);
        output.push_str("\" TargetMode=\"Internal\" Type=\"");
        escape_attribute(&relationship.rel_type, &mut output);
        output.push_str("\"></Relationship>");
    }
    output.push_str("</Relationships>");
    Some(output.into_bytes())
}

fn compare_digest(
    reference: &ReferenceSpec,
    actual_bytes: Option<Vec<u8>>,
    issues: &mut Vec<SignatureIssue>,
) -> bool {
    let Some(actual_bytes) = actual_bytes else {
        return false;
    };
    let actual = Sha256::digest(&actual_bytes).to_vec();
    if actual == reference.digest {
        return true;
    }
    issues.push(SignatureIssue::ChangedReference {
        uri: reference.uri.clone(),
        expected_digest: BASE64.encode(&reference.digest),
        actual_digest: BASE64.encode(actual),
    });
    false
}

fn add_uncovered_issues(
    package: &OpcPackage,
    covered_parts: &BTreeSet<String>,
    covered_relationships: &BTreeSet<CoveredRelationship>,
    issues: &mut Vec<SignatureIssue>,
) {
    let infrastructure = signature_infrastructure_parts(package);
    if !covered_parts.contains("/[Content_Types].xml") {
        issues.push(SignatureIssue::UncoveredPart {
            part_name: "/[Content_Types].xml".to_string(),
        });
    }
    let mut parts: Vec<&String> = package
        .parts
        .keys()
        .filter(|name| !infrastructure.contains(&crate::package::part_identity(name)))
        .collect();
    parts.sort();
    for part_name in parts {
        if !covered_parts.contains(part_name) {
            issues.push(SignatureIssue::UncoveredPart {
                part_name: part_name.clone(),
            });
        }
    }

    add_uncovered_relationships("/", &package.package_rels, covered_relationships, issues);
    let mut sources: Vec<&String> = package.part_rels.keys().collect();
    sources.sort();
    for source in sources {
        add_uncovered_relationships(
            source,
            &package.part_rels[source],
            covered_relationships,
            issues,
        );
    }
}

fn signature_infrastructure_parts(package: &OpcPackage) -> BTreeSet<String> {
    let mut parts = BTreeSet::new();
    for origin in package
        .package_rels
        .get_all_by_type(rel_types::DIGITAL_SIGNATURE_ORIGIN)
    {
        if is_external(origin) {
            continue;
        }
        let origin_part = OpcPackage::resolve_rel_target("/", &origin.target);
        if let Some(stored_origin) = package.stored_part_name(&origin_part) {
            parts.insert(crate::package::part_identity(stored_origin));
        }
        if let Some(relationships) = package.get_part_rels(&origin_part) {
            for signature in relationships.get_all_by_type(rel_types::DIGITAL_SIGNATURE) {
                if !is_external(signature) {
                    parts.insert(crate::package::part_identity(
                        &OpcPackage::resolve_rel_target(&origin_part, &signature.target),
                    ));
                }
            }
        }
    }
    parts
}

fn add_uncovered_relationships(
    source: &str,
    relationships: &Relationships,
    covered: &BTreeSet<CoveredRelationship>,
    issues: &mut Vec<SignatureIssue>,
) {
    let mut items: Vec<&Relationship> = relationships
        .items
        .iter()
        .filter(|relationship| {
            relationship.rel_type != rel_types::DIGITAL_SIGNATURE_ORIGIN
                && relationship.rel_type != rel_types::DIGITAL_SIGNATURE
        })
        .collect();
    items.sort_by(|left, right| left.id.cmp(&right.id));
    for relationship in items {
        let item = CoveredRelationship {
            source_part: source.to_string(),
            relationship_id: relationship.id.clone(),
        };
        if !covered.contains(&item) {
            issues.push(SignatureIssue::UncoveredRelationship {
                source_part: source.to_string(),
                relationship_id: relationship.id.clone(),
            });
        }
    }
}

pub(crate) fn content_types_bytes(package: &OpcPackage) -> Result<Vec<u8>> {
    package.content_types_bytes()
}

fn relationships_for_uri<'a>(
    package: &'a OpcPackage,
    part_name: &str,
) -> Option<(String, &'a Relationships)> {
    if part_name.eq_ignore_ascii_case("/_rels/.rels") {
        return Some(("/".to_string(), &package.package_rels));
    }
    let marker = "/_rels/";
    let lowercase = part_name.to_ascii_lowercase();
    let marker_index = lowercase.rfind(marker)?;
    let tail = part_name.get(marker_index + marker.len()..)?;
    if !tail.to_ascii_lowercase().ends_with(".rels") {
        return None;
    }
    let filename = tail.get(..tail.len().checked_sub(".rels".len())?)?;
    let directory = &part_name[..marker_index + 1];
    let source = format!("{directory}{filename}");
    let stored_source = package
        .stored_part_rels_owner(&source)
        .unwrap_or(&source)
        .to_owned();
    package
        .get_part_rels(&stored_source)
        .map(|relationships| (stored_source, relationships))
}

fn reference_part_name(uri: &str) -> Result<String> {
    let path = uri.split('?').next().unwrap_or(uri);
    if path.is_empty() || !path.starts_with('/') {
        return Err(invalid_signature(&format!(
            "non-package reference URI {uri}"
        )));
    }
    percent_decode(path)
}

fn percent_decode(value: &str) -> Result<String> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] == b'%' {
            let high = bytes.get(index + 1).and_then(|byte| hex(*byte));
            let low = bytes.get(index + 2).and_then(|byte| hex(*byte));
            let (Some(high), Some(low)) = (high, low) else {
                return Err(invalid_signature(&format!(
                    "invalid percent escape in {value}"
                )));
            };
            decoded.push((high << 4) | low);
            index += 3;
        } else {
            decoded.push(bytes[index]);
            index += 1;
        }
    }
    String::from_utf8(decoded).map_err(|_| invalid_signature("reference URI is not UTF-8"))
}

fn hex(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

fn parse_xml(xml: &[u8]) -> Result<XmlElement> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(false);
    let mut stack: Vec<XmlElement> = Vec::new();
    let mut root = None;
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(start)) => {
                let inherited = stack.last().map(|element| &element.namespaces);
                stack.push(parse_element_start(&reader, &start, inherited)?);
            }
            Ok(Event::Empty(start)) => {
                let inherited = stack.last().map(|element| &element.namespaces);
                let element = parse_element_start(&reader, &start, inherited)?;
                attach_element(element, &mut stack, &mut root)?;
            }
            Ok(Event::End(_)) => {
                let element = stack
                    .pop()
                    .ok_or_else(|| invalid_signature("unexpected closing element"))?;
                attach_element(element, &mut stack, &mut root)?;
            }
            Ok(Event::Text(text)) => {
                if let Some(parent) = stack.last_mut() {
                    let decoded = text
                        .xml_content(quick_xml::XmlVersion::Implicit1_0)
                        .map_err(|error| invalid_signature(&error.to_string()))?;
                    let unescaped = quick_xml::escape::unescape(&decoded)
                        .map_err(|error| invalid_signature(&error.to_string()))?;
                    parent.children.push(XmlNode::Text(unescaped.into_owned()));
                }
            }
            Ok(Event::CData(text)) => {
                if let Some(parent) = stack.last_mut() {
                    let decoded = text
                        .xml_content(quick_xml::XmlVersion::Implicit1_0)
                        .map_err(|error| invalid_signature(&error.to_string()))?;
                    parent.children.push(XmlNode::Text(decoded.into_owned()));
                }
            }
            Ok(Event::PI(instruction)) => {
                if let Some(parent) = stack.last_mut() {
                    let target = reader
                        .decoder()
                        .decode(instruction.target())
                        .map_err(|error| invalid_signature(&error.to_string()))?;
                    let content = reader
                        .decoder()
                        .decode(instruction.content())
                        .map_err(|error| invalid_signature(&error.to_string()))?;
                    parent.children.push(XmlNode::ProcessingInstruction {
                        target: normalize_xml_line_endings(&target),
                        content: normalize_xml_line_endings(&content),
                    });
                }
            }
            Ok(Event::Eof) => break,
            Ok(Event::Decl(_) | Event::Comment(_) | Event::DocType(_)) => {}
            Ok(Event::GeneralRef(reference)) => {
                return Err(invalid_signature(&format!(
                    "unresolved general entity {}",
                    String::from_utf8_lossy(reference.as_ref())
                )));
            }
            Err(error) => return Err(error.into()),
        }
        buffer.clear();
    }
    if !stack.is_empty() {
        return Err(invalid_signature("unclosed XML element"));
    }
    root.ok_or_else(|| invalid_signature("signature XML has no root element"))
}

fn parse_element_start(
    reader: &Reader<&[u8]>,
    start: &BytesStart<'_>,
    inherited: Option<&BTreeMap<String, String>>,
) -> Result<XmlElement> {
    let mut namespaces = inherited.cloned().unwrap_or_default();
    namespaces.insert("xml".to_string(), XML_NS.to_string());
    let mut decoded_attributes = Vec::new();
    for attribute in start.attributes() {
        let attribute = attribute?;
        let key = std::str::from_utf8(attribute.key.as_ref())?.to_string();
        let value = attribute
            .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())?
            .into_owned();
        if key == "xmlns" {
            namespaces.insert(String::new(), value);
        } else if let Some(prefix) = key.strip_prefix("xmlns:") {
            namespaces.insert(prefix.to_string(), value);
        } else {
            decoded_attributes.push((key, value));
        }
    }
    let qname = std::str::from_utf8(start.name().as_ref())?.to_string();
    let name = resolve_name(&qname, &namespaces, true)?;
    let attributes = decoded_attributes
        .into_iter()
        .map(|(qname, value)| {
            Ok(XmlAttribute {
                name: resolve_name(&qname, &namespaces, false)?,
                value,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    Ok(XmlElement {
        name,
        attributes,
        children: Vec::new(),
        namespaces,
    })
}

fn resolve_name(
    qname: &str,
    namespaces: &BTreeMap<String, String>,
    use_default: bool,
) -> Result<XmlName> {
    let (prefix, local) = qname
        .split_once(':')
        .map_or(("", qname), |(prefix, local)| (prefix, local));
    let namespace = if prefix.is_empty() && !use_default {
        String::new()
    } else {
        namespaces.get(prefix).cloned().unwrap_or_default()
    };
    if !prefix.is_empty() && namespace.is_empty() {
        return Err(invalid_signature(&format!(
            "unbound namespace prefix {prefix}"
        )));
    }
    Ok(XmlName {
        qname: qname.to_string(),
        prefix: prefix.to_string(),
        local: local.to_string(),
        namespace,
    })
}

fn attach_element(
    element: XmlElement,
    stack: &mut [XmlElement],
    root: &mut Option<XmlElement>,
) -> Result<()> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(XmlNode::Element(element));
    } else if root.replace(element).is_some() {
        return Err(invalid_signature("signature XML has multiple roots"));
    }
    Ok(())
}

fn canonicalize(element: &XmlElement) -> Vec<u8> {
    let mut output = String::new();
    write_canonical(element, &BTreeMap::new(), &mut output);
    output.into_bytes()
}

fn write_canonical(
    element: &XmlElement,
    rendered_namespaces: &BTreeMap<String, String>,
    output: &mut String,
) {
    output.push('<');
    output.push_str(&element.name.qname);

    let mut visible = BTreeSet::new();
    visible.insert(element.name.prefix.clone());
    for attribute in &element.attributes {
        if !attribute.name.prefix.is_empty() && attribute.name.prefix != "xml" {
            visible.insert(attribute.name.prefix.clone());
        }
    }
    let mut child_namespaces = rendered_namespaces.clone();
    for prefix in visible {
        let uri = if prefix.is_empty() {
            element.name.namespace.as_str()
        } else {
            element
                .namespaces
                .get(&prefix)
                .map(String::as_str)
                .unwrap_or("")
        };
        if rendered_namespaces.get(&prefix).map(String::as_str) != Some(uri) {
            if prefix.is_empty() {
                output.push_str(" xmlns=\"");
            } else {
                output.push_str(" xmlns:");
                output.push_str(&prefix);
                output.push_str("=\"");
            }
            escape_attribute(uri, output);
            output.push('"');
            child_namespaces.insert(prefix, uri.to_string());
        }
    }

    let mut attributes: Vec<&XmlAttribute> = element.attributes.iter().collect();
    attributes.sort_by(|left, right| {
        (left.name.namespace.as_str(), left.name.local.as_str())
            .cmp(&(right.name.namespace.as_str(), right.name.local.as_str()))
    });
    for attribute in attributes {
        output.push(' ');
        output.push_str(&attribute.name.qname);
        output.push_str("=\"");
        escape_attribute(&attribute.value, output);
        output.push('"');
    }
    output.push('>');
    for child in &element.children {
        match child {
            XmlNode::Element(child) => write_canonical(child, &child_namespaces, output),
            XmlNode::Text(text) => escape_text(text, output),
            XmlNode::ProcessingInstruction { target, content } => {
                output.push_str("<?");
                output.push_str(target);
                escape_processing_instruction(content, output);
                output.push_str("?>");
            }
        }
    }
    output.push_str("</");
    output.push_str(&element.name.qname);
    output.push('>');
}

fn escape_processing_instruction(value: &str, output: &mut String) {
    for character in value.chars() {
        if character == '\r' {
            output.push_str("&#xD;");
        } else {
            output.push(character);
        }
    }
}

fn normalize_xml_line_endings(value: &str) -> String {
    let mut normalized = String::with_capacity(value.len());
    let mut characters = value.chars().peekable();
    while let Some(character) = characters.next() {
        if character == '\r' {
            if characters.peek() == Some(&'\n') {
                characters.next();
            }
            normalized.push('\n');
        } else {
            normalized.push(character);
        }
    }
    normalized
}

fn escape_attribute(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '"' => output.push_str("&quot;"),
            '\t' => output.push_str("&#x9;"),
            '\n' => output.push_str("&#xA;"),
            '\r' => output.push_str("&#xD;"),
            _ => output.push(character),
        }
    }
}

fn escape_text(value: &str, output: &mut String) {
    for character in value.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '\r' => output.push_str("&#xD;"),
            _ => output.push(character),
        }
    }
}

fn require_algorithm(element: &XmlElement, kind: &'static str, expected: &str) -> Result<()> {
    let actual = attribute(element, "", "Algorithm")
        .ok_or_else(|| invalid_signature(&format!("{} has no Algorithm", element.name.qname)))?;
    if actual != expected {
        return Err(unsupported(kind, actual));
    }
    Ok(())
}

fn require_name(element: &XmlElement, namespace: &str, local: &str) -> Result<()> {
    if has_name(element, namespace, local) {
        Ok(())
    } else {
        Err(invalid_signature(&format!(
            "expected {{{namespace}}}{local}, found {}",
            element.name.qname
        )))
    }
}

fn direct_child<'a>(
    element: &'a XmlElement,
    namespace: &str,
    local: &str,
) -> Result<&'a XmlElement> {
    optional_direct_child(element, namespace, local)?
        .ok_or_else(|| invalid_signature(&format!("missing {{{namespace}}}{local}")))
}

fn optional_direct_child<'a>(
    element: &'a XmlElement,
    namespace: &str,
    local: &str,
) -> Result<Option<&'a XmlElement>> {
    let matches: Vec<&XmlElement> = element_children(element)
        .filter(|child| has_name(child, namespace, local))
        .collect();
    match matches.as_slice() {
        [] => Ok(None),
        [child] => Ok(Some(*child)),
        _ => Err(invalid_signature(&format!(
            "duplicate {{{namespace}}}{local}"
        ))),
    }
}

fn direct_children<'a>(
    element: &'a XmlElement,
    namespace: &'a str,
    local: &'a str,
) -> impl Iterator<Item = &'a XmlElement> {
    element_children(element).filter(move |child| has_name(child, namespace, local))
}

fn element_children(element: &XmlElement) -> impl Iterator<Item = &XmlElement> {
    element.children.iter().filter_map(|child| match child {
        XmlNode::Element(element) => Some(element),
        XmlNode::Text(_) | XmlNode::ProcessingInstruction { .. } => None,
    })
}

fn descendant<'a>(element: &'a XmlElement, namespace: &str, local: &str) -> Option<&'a XmlElement> {
    element_children(element).find_map(|child| {
        has_name(child, namespace, local)
            .then_some(child)
            .or_else(|| descendant(child, namespace, local))
    })
}

fn collect_matching_elements<'a>(
    element: &'a XmlElement,
    namespace: &str,
    local: &str,
    found: &mut Vec<&'a XmlElement>,
) {
    if has_name(element, namespace, local) {
        found.push(element);
    }
    collect_descendants(element, namespace, local, found);
}

fn collect_descendants<'a>(
    element: &'a XmlElement,
    namespace: &str,
    local: &str,
    found: &mut Vec<&'a XmlElement>,
) {
    for child in element_children(element) {
        if has_name(child, namespace, local) {
            found.push(child);
        }
        collect_descendants(child, namespace, local, found);
    }
}

fn find_all_by_id<'a>(element: &'a XmlElement, id: &str) -> Vec<&'a XmlElement> {
    let mut found = Vec::new();
    collect_by_id(element, id, &mut found);
    found
}

fn collect_by_id<'a>(element: &'a XmlElement, id: &str, found: &mut Vec<&'a XmlElement>) {
    if attribute(element, "", "Id") == Some(id) {
        found.push(element);
    }
    for child in element_children(element) {
        collect_by_id(child, id, found);
    }
}

fn has_name(element: &XmlElement, namespace: &str, local: &str) -> bool {
    element.name.namespace == namespace && element.name.local == local
}

fn attribute<'a>(element: &'a XmlElement, namespace: &str, local: &str) -> Option<&'a str> {
    element
        .attributes
        .iter()
        .find(|attribute| attribute.name.namespace == namespace && attribute.name.local == local)
        .map(|attribute| attribute.value.as_str())
}

fn element_text(element: &XmlElement) -> String {
    let mut text = String::new();
    collect_text(element, &mut text);
    text.trim().to_string()
}

fn collect_text(element: &XmlElement, output: &mut String) {
    for child in &element.children {
        match child {
            XmlNode::Element(child) => collect_text(child, output),
            XmlNode::Text(text) => output.push_str(text),
            XmlNode::ProcessingInstruction { .. } => {}
        }
    }
}

fn decode_base64(value: &str, field: &str) -> Result<Vec<u8>> {
    let compact: String = value
        .chars()
        .filter(|character| !character.is_whitespace())
        .collect();
    BASE64
        .decode(compact)
        .map_err(|error| invalid_signature(&format!("invalid {field}: {error}")))
}

fn is_external(relationship: &Relationship) -> bool {
    relationship
        .target_mode
        .as_deref()
        .is_some_and(|mode| mode.eq_ignore_ascii_case("external"))
}

fn is_signature_part(name: &str) -> bool {
    name.get(.."/_xmlsignatures/".len())
        .is_some_and(|prefix| prefix.eq_ignore_ascii_case("/_xmlsignatures/"))
}

fn is_signature_xml_part(name: &str) -> bool {
    is_signature_part(name)
        && name
            .get(name.len().saturating_sub(4)..)
            .is_some_and(|suffix| suffix.eq_ignore_ascii_case(".xml"))
}

fn is_coverage_issue(issue: &SignatureIssue) -> bool {
    !matches!(
        issue,
        SignatureIssue::SignatureValueMismatch | SignatureIssue::ChangedReference { .. }
    )
}

fn invalid_signature(message: &str) -> OpcError {
    OpcError::InvalidSignatureXml(message.to_string())
}

fn unsupported(kind: &'static str, algorithm: &str) -> OpcError {
    OpcError::UnsupportedSignatureAlgorithm {
        kind,
        algorithm: algorithm.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use std::io::Cursor;

    use super::*;

    const TEST_PRIVATE_KEY_PKCS8_BASE64: &str = "MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCXVziWlLINGRFmqL/FTgWy1/o2zEkmQOHt23YQZDm7PiiBjq+ib85OF6UH9nH9VTr3rE4njpTj0TgzbTaY0ftxQqQg+sdSLZu6+FGAo+erEQAEoa9HZ6ys/9cuFbeYLUVVSQ33wtdq+VMDWMZqXxXd53QX9aqxBg1cBY4XAHgqWQ50X3AQ/IBhMTpZ0Wte3wGZeC+i1NQ45Ws9HwpwTgjg1mX3BTUaPtZW801SBOC2b3ug1jGJXO/fGigdzDwGcJKAXt/NZwa+PIhKv8GyebstGwCEu/detSmaODbMs7cYb4liJK2NLCY99Y7BLO4xMNg2hwploUwAdu+RgEf5U8LvAgMBAAECggEAEg5S7wxAjfV+sPvTHWwom+TOsnj/BTRagDFdzajXhnJtDMAETmH+gCysAN4zTWE8zs3c6TVGqEOO6/vMtsDeue2UfWbOHwzX9p+nwaxMeIlnsiXELsW8wUso1hO7Osmz6u/zXar+XoHumIif65L6neX+YNlriwFI2MDE6hOhQpP8eqA+u8DX1g2Oq0AhpPX6g5ABC422GjB6Z/NsXnwrFb4SpUy7aQNCHjDBQsetc2uE9hStvGvfwg4qvQl4NtL2k2udU/wOO66jvLBMbsF6Blz5vpljGpDuYJk/9sTTmzH+cJNob8nlOWwXVwNrAPMHe0l4RK8+DUqWvdEqkxZBoQKBgQDLOmt9DPX80aY4QucFNOGjjqH+Gk+6jEI6LrKfXiCswCiziWhoVCSKkQSi4k2Ol0NEKfcKy5jUYxglwjLDF/MBqz0rCKX/XjZ/hO8f8DlQFeZdLxvsFVH/lUl98JIifrrpMnO9za/0SfcMsy3CCglz8ho9+LUch8fQx4mTX3MCfwKBgQC+o5VqOZIJJXuYq6nWdjj5BPWKC8kdNq+t2Z5FnlsDuwNmSzCJI8juUPyId3pI0fiSbmdTfp8z4OrV46FrkFJ/oY8JJ5X3s2sfvviSVgvgZ/G+wJ4fQpN9yRg4Dp/RCgT7xrqxXtLTpjl7X0RrKri2uxyhh7AU/9mEuviwmTInkQKBgBed8l/V4cA/nNFs9Ovl+VLIgIrHA/zpz8hzJM7gYWux6Qj0Lu3w2U5BDAjhw6GOcoK5XbwjbN9BpMy+hKenYNYQ0Erv9lp22F55VFCh2gc0hFDP6K7Gy4CoGKJKErFviMkQ0+J6xLfe4JbZO7gQ8ohG2kXZYTKvlMjuZ055CSSBAoGAdTgOgl9dzSPwCGLdLlJJG80R0U0H31+lzAb4S6RgID4YjAiFkn2fafIAJUUZurbo2djqzasY5wRQQS4TLhlysKm9Uoq1qrX2k3GQVCJ2cQhY28qCL4R3PiutKaLMX/OCNvHuD2vXxG38AEEGx8JgC3On2iadfXwH2pZAng3EihECgYEAn9DRQhgazqs35kZgEwqkpBKovhd32pDRkA6XEoNHB1uTbHTAMuo53SOZbS1bjGB2B2fKwvxMkAFgKAkEhVqOIXXzgxmn8cv4C40QZbrC8Mjrfax6wbNmAGfmKawxxJtRrx3EoAB2ti9D7JaQxgZiIxdvrv4bd/8Ejd7d8P1as7c=";
    const TEST_EC_PRIVATE_KEY_PKCS8_BASE64: &str = "MIGHAgEAMBMGByqGSM49AgEGCCqGSM49AwEHBG0wawIBAQQgSpNKQ0l3N9dnwz0s3C9J0Loybv4rGe+1MKTapaW2XCyhRANCAAQQldnc+dRJT0LLEXXnxzokgFBlWnlSzkLBoCEEmxS0icOE3gOeWMrjHTgSa80tME74QlEtKfsuiIaQcBWjH2JU";
    const TEST_CERTIFICATE_DER_BASE64: &str = "MIIDdDCCAlygAwIBAgIUFuFlT5/whakB7xH7oRLpOMWRK6swDQYJKoZIhvcNAQELBQAwQzEaMBgGA1UEAwwRRi0xNzIgVGVzdCBTaWduZXIxGDAWBgNVBAoMD1RlbnNvcmJlZSBUZXN0czELMAkGA1UEBhMCR0IwHhcNMjYwODIyMjIyOTQ4WhcNMzYwODE5MjIyOTQ4WjBDMRowGAYDVQQDDBFGLTE3MiBUZXN0IFNpZ25lcjEYMBYGA1UECgwPVGVuc29yYmVlIFRlc3RzMQswCQYDVQQGEwJHQjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJdXOJaUsg0ZEWaov8VOBbLX+jbMSSZA4e3bdhBkObs+KIGOr6Jvzk4XpQf2cf1VOvesTieOlOPRODNtNpjR+3FCpCD6x1Itm7r4UYCj56sRAAShr0dnrKz/1y4Vt5gtRVVJDffC12r5UwNYxmpfFd3ndBf1qrEGDVwFjhcAeCpZDnRfcBD8gGExOlnRa17fAZl4L6LU1Djlaz0fCnBOCODWZfcFNRo+1lbzTVIE4LZve6DWMYlc798aKB3MPAZwkoBe381nBr48iEq/wbJ5uy0bAIS79161KZo4NsyztxhviWIkrY0sJj31jsEs7jEw2DaHCmWhTAB275GAR/lTwu8CAwEAAaNgMF4wHQYDVR0OBBYEFN+iqzK8SA1T4nFG9QkuPzX2343DMB8GA1UdIwQYMBaAFN+iqzK8SA1T4nFG9QkuPzX2343DMAwGA1UdEwEB/wQCMAAwDgYDVR0PAQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4IBAQBqjXqb8Wm/QzMHl7jLk42TJCwpTBW86nK/KuQb6XLfjEQ3hy4nSchLGZHlb5yladTu7KFp2DIDHRRVMDaxSUWd0bVcwus3D8JCjuydeLsOiGlQTtrDMp0eEBR3NU//047vDLRodwv7q0cE+r9d6z8bXprXBKqKrQrVweaZP6eNxb7Cz6aALENNDrSd5sGy1CfOC5hZzBzFcyQUp+bBmc5XL5MzO3sm5yQebCfqiwBiBwkNEGkjfOIxU0O+LxHjOzXrU2zLbOcn5WvPl+UkD/OzD5NjobsZk4wUtGqSe4+KGgb1VxK8JSnJOBqnUyPc/4KNl8JQROopBKTUxENfWyKM";

    const SIGNATURE_DS_BASE64: &str = "RMIFfzT/GCclCbe0UiSNTXzGrz/cuxmfyniUJLDTgpmJBv7NBSQ0LMJY71qcPE5jRhkdC7AUeWKQrIdFAjHVRqYrBDWEaNdUmiRcbdEw67c3nVF33Zmuj24FULFC58/UMhP7m3Rvl2MfWzPY+9EPoEcCfyDkvzZYaurSxnYf5E2L/3rI7Exv/QBrfVwv7bC+I/ReuoRTc9rToOMQUowsLLgcDC5MFz0onvHGUaM/azClme+AJHPPYky2k5aKWmiN+mcBIUrlZYUnx+tdnzcf1DVdCk1YpNc9xqWndFSIjiDd0gfeDWq++VsJBXme02aVMLLXciVrOlvAYaMVdawUHw==";
    const SIGNATURE_SIG_BASE64: &str = "PoIHdPAtYTKOlVcxhWvUEHrlYf58I53jEJFbeiVb1BrQq3MLln4Bq8uHjY6gPs043hDausEDwuf44Olw/l2twOPmyiDPotpCNZe+yg6QeWPYMcWMM2PedUsMflLaYxOOFCFSAolH6MWR6N7w8iIgUSohay5B7KyRiuOOTu2nvlcxZJhei0EEiQFnPlfLhijZDZ97V0I+JAAPy8DDvFF7NOAUOz15yIN1AVgKtzUdgdsCvqR6wKV1+2XEfDftBt462QHOhkW5lrHcMg/THY0yEHjw1VXN6hwBPshzTORXeJexr0lmcJJt1p1vyyjPuDEiM8/JYpDDF8xP0B44/uE4Cg==";
    const CERTIFICATE_DER_BASE64: &str = "MIIDZzCCAk+gAwIBAgIUHP/5/GIqsdVlsrP3bYWND71j/aswDQYJKoZIhvcNAQELBQAwQzEaMBgGA1UEAwwRRi0xNzEgVGVzdCBTaWduZXIxGDAWBgNVBAoMD1RlbnNvcmJlZSBUZXN0czELMAkGA1UEBhMCR0IwHhcNMjYwODIyMDE1NTAxWhcNMzYwODE5MDE1NTAxWjBDMRowGAYDVQQDDBFGLTE3MSBUZXN0IFNpZ25lcjEYMBYGA1UECgwPVGVuc29yYmVlIFRlc3RzMQswCQYDVQQGEwJHQjCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALRYSR2vue+MSg8VlQKoYDwwZk+b3rs7F2j6u3e5Wbanf/ABy4378xFq1xB6M5B3T+kabTVeENWx+IE3d2r8Z7KxOxsHhZyhrcS4QMa2alL9ErE2JrrTNFsM+ciI02xgLIr+baSig2iV5uzAgcqe9jMMurpeyCmaZxIAhpy3bjKUgyWVBgmZsznfDLMbJfC4u4fifY5Lpsghb9ewaV4euP/8/TE5orJOC80kWyqgbkws+RzoA1KyyBf6EzNtUzBBhEH+iBkhZtYVXuZcaq5WlkcKvQLzhsxPWSjD5CPU0tohtJkGP3ugDczifJRXH5i55GZjUPB9SGKdV6Gsi3Ow3x0CAwEAAaNTMFEwHQYDVR0OBBYEFG8Kul280npBFhtYiONVazzJqGNTMB8GA1UdIwQYMBaAFG8Kul280npBFhtYiONVazzJqGNTMA8GA1UdEwEB/wQFMAMBAf8wDQYJKoZIhvcNAQELBQADggEBAFd9eu72wlh9rtleRKvPGOjZN55GtQistxUqBWdNNSBEDIQDIvZ9gM3ix+IkM45zS5rX6ME8eqeAUIiP40izNnrOIk4PT6f68hL/dCrDmcxouU0RHI2OgHHz9Ir2n94Lz+dBWx+8Syafa48GfvfpKKXo4/qQQNk0idZ524ftYT7RpCL11z2QuzcBHnD7ZoubdRYo6i12WPGHo2mZu0BS1odKZajMvBPhdF4rH2TqGUgYyjeeHUb+8cnYz9OPgGB/6hvv+sBCqEHWDnsVYIJzhjDQS6l0v0YVeE7RIwHrxPvEF+dGAqLfDYoZ6mWe50jBtrOq9otI13ncnccEHm8Aovo=";
    const EC_CERTIFICATE_DER_BASE64: &str = "MIIB4DCCAYegAwIBAgIUf1e7Wg30URPA6mcy7mqRrxSGfkUwCgYIKoZIzj0EAwIwRjEdMBsGA1UEAwwURi0xNzEgRUMgVGVzdCBTaWduZXIxGDAWBgNVBAoMD1RlbnNvcmJlZSBUZXN0czELMAkGA1UEBhMCR0IwHhcNMjYwODIyMDIwMTE4WhcNMzYwODE5MDIwMTE4WjBGMR0wGwYDVQQDDBRGLTE3MSBFQyBUZXN0IFNpZ25lcjEYMBYGA1UECgwPVGVuc29yYmVlIFRlc3RzMQswCQYDVQQGEwJHQjBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABDzWARcxjZ++hPgwD2LB/Hz+cJ2fm0Fjb3uE3NNhH0gItnMAQJvS3KQBRUU28gVYra2iDQabYYcCGC5LNX2yXCmjUzBRMB0GA1UdDgQWBBTiKMGSnjWxKe/LNO0uLALvpD8QrTAfBgNVHSMEGDAWgBTiKMGSnjWxKe/LNO0uLALvpD8QrTAPBgNVHRMBAf8EBTADAQH/MAoGCCqGSM49BAMCA0cAMEQCIBd6DFidlSSE5BApuxqqpsXpx5bHWZJrpn89432E5LJoAiB0akoHnuZREqsmnrJ9GXehgUMXYqf7hiVdcB4JMXeE4w==";

    #[test]
    fn signing_material_strictness_and_failures_are_atomic() {
        let bytes = |package: &OpcPackage| {
            let mut output = Cursor::new(Vec::new());
            package.write_to(&mut output).unwrap();
            output.into_inner()
        };
        let der = test_private_key();
        assert_eq!(signing_material_der(&der, "PRIVATE KEY").unwrap(), der);
        let mut trailing = der.clone();
        trailing.push(b' ');
        assert_eq!(
            signing_material_der(&trailing, "PRIVATE KEY").unwrap(),
            trailing
        );
        let valid = format!(
            "-----BEGIN PRIVATE KEY-----\n{}\n-----END PRIVATE KEY-----",
            BASE64.encode(&der)
        );
        let invalid = [
            Vec::new(),
            vec![b'x'; SIGNING_MATERIAL_LIMIT + 1],
            valid.replace("PRIVATE KEY", "RSA PRIVATE KEY").into_bytes(),
            valid
                .replace("PRIVATE KEY", "ENCRYPTED PRIVATE KEY")
                .into_bytes(),
            valid
                .replace("END PRIVATE KEY", "END CERTIFICATE")
                .into_bytes(),
            format!("{valid}\n{valid}").into_bytes(),
            format!("{valid}trailer").into_bytes(),
            format!("é{valid}").into_bytes(),
            b"-----BEGIN PRIVATE KEY-----\nZh==\n-----END PRIVATE KEY-----".to_vec(),
            b"-----BEGIN PRIVATE KEY-----\nZg\n-----END PRIVATE KEY-----".to_vec(),
            b"-----BEGIN PRIVATE KEY-----\nProc-Type: 4,ENCRYPTED\nZg==\n-----END PRIVATE KEY-----"
                .to_vec(),
            b"-----BEGIN PRIVATE KEY-----\n-----END PRIVATE KEY-----".to_vec(),
        ];
        for key in invalid.iter().chain(std::iter::once(&trailing)) {
            let mut package = unsigned_package();
            let before = bytes(&package);
            assert!(package.sign(key, &test_certificate()).is_err());
            assert_eq!(before, bytes(&package));
            assert!(package.verify_signatures().unwrap().is_empty());
        }
        for cert in [
            &b"-----BEGIN PRIVATE KEY-----Zg==-----END PRIVATE KEY-----"[..],
            &vec![0; SIGNING_MATERIAL_LIMIT + 1],
        ] {
            let mut package = unsigned_package();
            let before = bytes(&package);
            assert!(matches!(
                package.sign(&der, cert),
                Err(OpcError::InvalidSigningCertificate(_))
            ));
            assert_eq!(before, bytes(&package));
        }
    }

    #[test]
    fn signing_accepts_single_strict_pem_blocks() {
        for newline in ["\n", "\r\n"] {
            let pem = |label: &str, bytes: &[u8]| {
                let encoded = BASE64.encode(bytes);
                let lines = encoded
                    .as_bytes()
                    .chunks(64)
                    .map(|line| std::str::from_utf8(line).unwrap())
                    .collect::<Vec<_>>()
                    .join(newline);
                format!(
                    " \t{newline}-----BEGIN {label}-----{newline}{lines}{newline}-----END {label}-----{newline} "
                )
            };
            let mut package = unsigned_package();
            let report = package
                .sign(
                    pem("PRIVATE KEY", &test_private_key()).as_bytes(),
                    pem("CERTIFICATE", &test_certificate()).as_bytes(),
                )
                .unwrap();
            assert!(report.cryptographically_valid && report.coverage_complete);
            assert_eq!(package.verify_signatures().unwrap(), vec![report]);
        }
    }

    #[test]
    fn signature_creation_uses_schema_order_and_complete_canonical_references() {
        let mut package = unsigned_package();
        package.set_part("/word/second.xml", b"second".to_vec());
        package
            .content_types
            .add_override("/word/second.xml", "application/xml");
        package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("rId2", "urn:test:second", "second.xml");
        package
            .get_or_create_part_rels("/word/document.xml")
            .add_with_id("rId1", "urn:test:self", "document.xml");
        let report = package
            .sign(&test_private_key(), &test_certificate())
            .unwrap();
        assert!(report.cryptographically_valid);
        assert!(report.coverage_complete);
        assert_eq!(
            report.covered_parts,
            vec![
                "/[Content_Types].xml",
                "/word/document.xml",
                "/word/second.xml"
            ]
        );
        assert_eq!(report.covered_relationships.len(), 3);

        let root = parse_xml(&package.parts["/_xmlsignatures/sig1.xml"]).unwrap();
        let children: Vec<&str> = element_children(&root)
            .map(|child| child.name.local.as_str())
            .collect();
        assert_eq!(
            children,
            ["SignedInfo", "SignatureValue", "KeyInfo", "Object"]
        );
        assert_eq!(attribute(&root, "", "Id"), Some("idPackageSignature"));
        let object = direct_child(&root, DSIG_NS, "Object").unwrap();
        let object_children: Vec<&str> = element_children(object)
            .map(|child| child.name.local.as_str())
            .collect();
        assert_eq!(object_children, ["Manifest", "SignatureProperties"]);
        let signature_property = descendant(object, DSIG_NS, "SignatureProperty").unwrap();
        assert_eq!(
            attribute(signature_property, "", "Target"),
            Some("#idPackageSignature")
        );
        let signature_time =
            descendant(signature_property, OPC_SIGNATURE_NS, "SignatureTime").unwrap();
        assert_eq!(
            element_text(direct_child(signature_time, OPC_SIGNATURE_NS, "Format").unwrap()),
            "YYYY-MM-DDThh:mm:ssTZD"
        );
        let signature_time_value =
            element_text(direct_child(signature_time, OPC_SIGNATURE_NS, "Value").unwrap());
        assert!(signature_time_value.ends_with('Z'));
        let manifest = descendant(&root, DSIG_NS, "Manifest").unwrap();
        let references: Vec<&XmlElement> =
            direct_children(manifest, DSIG_NS, "Reference").collect();
        let uris: Vec<&str> = references
            .iter()
            .map(|reference| attribute(reference, "", "URI").unwrap())
            .collect();
        assert_eq!(
            uris,
            [
                "/%5BContent_Types%5D.xml",
                "/word/document.xml?ContentType=application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "/word/second.xml?ContentType=application/xml",
                "/_rels/.rels?ContentType=application/vnd.openxmlformats-package.relationships+xml",
                "/word/_rels/document.xml.rels?ContentType=application/vnd.openxmlformats-package.relationships+xml",
            ]
        );
        let relationship_transform = direct_children(
            direct_child(references[4], DSIG_NS, "Transforms").unwrap(),
            DSIG_NS,
            "Transform",
        )
        .next()
        .unwrap();
        let relationship_ids: Vec<&str> = element_children(relationship_transform)
            .map(|reference| attribute(reference, "", "SourceId").unwrap())
            .collect();
        assert_eq!(relationship_ids, ["rId1", "rId2"]);

        let mut occupied = unsigned_package();
        occupied.set_part("/_xmlsignatures/origin.sigs", b"occupied".to_vec());
        occupied
            .content_types
            .add_override("/_xmlsignatures/origin.sigs", "application/octet-stream");
        occupied
            .content_types
            .add_override("/_xmlsignatures/sig7.xml", "application/octet-stream");
        let report = occupied
            .sign(&test_private_key(), &test_certificate())
            .unwrap();
        assert_eq!(report.signature_part, "/_xmlsignatures/sig8.xml");
        assert!(occupied.parts.contains_key("/_xmlsignatures/origin1.sigs"));
        assert_eq!(*occupied.parts["/_xmlsignatures/origin.sigs"], b"occupied");
        assert!(!occupied.parts.contains_key("/_xmlsignatures/sig7.xml"));

        let mut already_signed = unsigned_package();
        already_signed
            .sign(&test_private_key(), &test_certificate())
            .unwrap();
        assert_signing_failure_is_atomic(&mut already_signed);
    }

    #[test]
    fn signature_creation_rejects_mismatched_or_unsupported_key_material() {
        let mut package = unsigned_package();
        let before = package_bytes(&package);
        let mismatched_certificate = BASE64.decode(CERTIFICATE_DER_BASE64).unwrap();
        let ec_key = BASE64.decode(TEST_EC_PRIVATE_KEY_PKCS8_BASE64).unwrap();
        let ec_certificate = BASE64.decode(EC_CERTIFICATE_DER_BASE64).unwrap();
        for result in [
            package.sign(b"not-pkcs8", &test_certificate()),
            package.sign(&test_private_key(), b"not-x509"),
            package.sign(&test_private_key(), &mismatched_certificate),
            package.sign(&ec_key, &test_certificate()),
            package.sign(&test_private_key(), &ec_certificate),
        ] {
            assert!(result.is_err());
            assert_eq!(package_bytes(&package), before);
        }
    }

    #[test]
    fn signed_package_verifies_with_complete_coverage() {
        let mut package = unsigned_package();
        package
            .sign(&test_private_key(), &test_certificate())
            .unwrap();
        let bytes = package_bytes(&package);
        let reopened = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
        let report = reopened.verify_signatures().unwrap().remove(0);
        assert!(report.cryptographically_valid);
        assert!(report.coverage_complete);
    }

    #[test]
    fn every_signature_creation_failure_leaves_live_package_unchanged() {
        let mut package = unsigned_package();
        let before = package_bytes(&package);
        assert!(package.sign(b"not-pkcs8", &test_certificate()).is_err());
        assert_eq!(package_bytes(&package), before);

        let mut external = unsigned_package();
        external
            .package_rels
            .add_external("urn:test:external", "https://example.invalid");
        assert_signing_failure_is_atomic(&mut external);

        let mut missing_target = unsigned_package();
        missing_target
            .package_rels
            .add("urn:test:missing", "word/missing.xml");
        assert_signing_failure_is_atomic(&mut missing_target);

        let mut future_origin_target = unsigned_package();
        future_origin_target
            .package_rels
            .add("urn:test:dangling", "_xmlsignatures/origin.sigs");
        assert_signing_failure_is_atomic(&mut future_origin_target);

        let mut future_signature_target = unsigned_package();
        future_signature_target
            .get_or_create_part_rels("/word/document.xml")
            .add("urn:test:dangling", "../_xmlsignatures/sig1.xml");
        assert_signing_failure_is_atomic(&mut future_signature_target);

        let mut misplaced_signature = unsigned_package();
        misplaced_signature
            .package_rels
            .add(rel_types::DIGITAL_SIGNATURE, "_xmlsignatures/missing.xml");
        assert_signing_failure_is_atomic(&mut misplaced_signature);

        let mut misplaced_origin = unsigned_package();
        misplaced_origin
            .get_or_create_part_rels("/word/document.xml")
            .add(
                rel_types::DIGITAL_SIGNATURE_ORIGIN,
                "../_xmlsignatures/origin.sigs",
            );
        assert_signing_failure_is_atomic(&mut misplaced_origin);

        let mut orphan_signature = unsigned_package();
        orphan_signature.set_part("/_xmlsignatures/sig7.xml", b"not-a-signature".to_vec());
        orphan_signature
            .content_types
            .add_override("/_xmlsignatures/sig7.xml", SIGNATURE_XML_CONTENT_TYPE);
        assert_signing_failure_is_atomic(&mut orphan_signature);

        let mut duplicate = unsigned_package();
        duplicate.package_rels.items.push(Relationship {
            id: "rId1".to_string(),
            rel_type: "urn:test:duplicate".to_string(),
            target: "word/document.xml".to_string(),
            target_mode: None,
        });
        let before = duplicate.clone();
        assert!(
            duplicate
                .sign(&test_private_key(), &test_certificate())
                .is_err()
        );
        assert_eq!(duplicate.parts, before.parts);
        assert_eq!(duplicate.content_types, before.content_types);
        assert_eq!(duplicate.package_rels.items, before.package_rels.items);
        assert_eq!(duplicate.part_rels.len(), before.part_rels.len());

        let mut orphan = unsigned_package();
        orphan
            .get_or_create_part_rels("/word/missing.xml")
            .add("urn:test:orphan", "document.xml");
        assert_signing_failure_is_atomic(&mut orphan);

        let mut empty_source = unsigned_package();
        empty_source
            .get_or_create_part_rels("")
            .add("urn:test:empty-source", "word/document.xml");
        let before = format!("{empty_source:?}");
        assert!(
            empty_source
                .sign(&test_private_key(), &test_certificate())
                .is_err()
        );
        assert_eq!(format!("{empty_source:?}"), before);
    }

    #[test]
    #[ignore = "validated manually with the pinned Microsoft Word for Mac oracle"]
    fn word_for_mac_recognizes_and_protects_the_created_signature() {
        let output = std::env::var("RDOCX_F172_ORACLE_OUTPUT")
            .expect("set RDOCX_F172_ORACLE_OUTPUT to the Word oracle DOCX path");
        let mut package = unsigned_package();
        package
            .sign(&test_private_key(), &test_certificate())
            .unwrap();
        let bytes = package_bytes(&package);
        std::fs::write(output, &bytes).unwrap();
        let reopened = OpcPackage::from_reader(Cursor::new(bytes)).unwrap();
        let report = reopened.verify_signatures().unwrap().remove(0);
        assert!(report.cryptographically_valid);
        assert!(report.coverage_complete);
        assert_eq!(
            std::env::var("RDOCX_F172_WORD_MAC_EVIDENCE").as_deref(),
            Ok("recognized-and-protected"),
            "set RDOCX_F172_WORD_MAC_EVIDENCE=recognized-and-protected only after Word for Mac recognizes the digital signature and protects the document from editing"
        );
    }

    #[test]
    fn unix_days_convert_to_utc_civil_dates() {
        assert_eq!(civil_date_from_unix_days(0), (1970, 1, 1));
        assert_eq!(civil_date_from_unix_days(20_323), (2025, 8, 23));
        assert_eq!(civil_date_from_unix_days(20_819), (2027, 1, 1));
    }

    #[test]
    fn signature_parser_is_prefix_tolerant_and_algorithm_strict() {
        let package = signed_package("sig");
        let report = package.verify_signatures().unwrap().remove(0);
        assert!(report.cryptographically_valid);

        for (from, to, kind) in [
            (DIGEST_SHA256, "urn:test:sha1", "digest method"),
            (
                SIGNATURE_RSA_SHA256,
                "urn:test:rsa-sha1",
                "signature method",
            ),
            (C14N_EXCLUSIVE, "urn:test:c14n", "canonicalization"),
        ] {
            let mut changed = signed_package("sig");
            mutate_signature_xml(&mut changed, |xml| xml.replacen(from, to, 1));
            assert!(matches!(
                changed.verify_signatures(),
                Err(OpcError::UnsupportedSignatureAlgorithm { kind: actual, .. }) if actual == kind
            ));
        }

        let unsupported_transform = format!(
            r#"<ds:Reference xmlns:ds="{DSIG_NS}" URI="/word/document.xml"><ds:Transforms><ds:Transform Algorithm="urn:test:relationship"></ds:Transform></ds:Transforms><ds:DigestMethod Algorithm="{DIGEST_SHA256}"></ds:DigestMethod><ds:DigestValue>{}</ds:DigestValue></ds:Reference>"#,
            BASE64.encode([0_u8; 32])
        );
        let reference = parse_xml(unsupported_transform.as_bytes()).unwrap();
        assert!(matches!(
            parse_reference(&reference),
            Err(OpcError::UnsupportedSignatureAlgorithm {
                kind: "transform",
                ..
            })
        ));

        let mut ec = signed_package("sig");
        mutate_signature_xml(&mut ec, |xml| {
            xml.replace(CERTIFICATE_DER_BASE64, EC_CERTIFICATE_DER_BASE64)
        });
        assert!(matches!(
            ec.verify_signatures(),
            Err(OpcError::UnsupportedSignatureAlgorithm {
                kind: "key algorithm",
                ..
            })
        ));
    }

    #[test]
    fn referenced_object_processing_instruction_is_canonical_and_mutation_sensitive() {
        let xml = format!(
            r#"<ds:Signature xmlns:ds="{DSIG_NS}"><ds:Object Id="pi-object"><ds:Value>before</ds:Value><?audit alpha&beta <ok?><ds:Value>after</ds:Value></ds:Object></ds:Signature>"#
        );
        let root = parse_xml(xml.as_bytes()).unwrap();
        let object = same_document_target(&root, "#pi-object").unwrap().unwrap();
        let expected = format!(
            r#"<ds:Object xmlns:ds="{DSIG_NS}" Id="pi-object"><ds:Value>before</ds:Value><?audit alpha&beta <ok?><ds:Value>after</ds:Value></ds:Object>"#
        );
        assert_eq!(canonicalize(object), expected.as_bytes());

        let reference = ReferenceSpec {
            uri: "#pi-object".to_string(),
            digest: Sha256::digest(expected.as_bytes()).to_vec(),
            transforms: vec![ReferenceTransform::ExclusiveCanonicalization],
        };
        let actual = apply_reference(
            &OpcPackage::new(),
            &root,
            &reference,
            &mut BTreeSet::new(),
            &mut BTreeSet::new(),
            &mut Vec::new(),
        )
        .unwrap();
        assert!(compare_digest(&reference, actual, &mut Vec::new()));

        let changed = parse_xml(xml.replace("alpha", "changed").as_bytes()).unwrap();
        let changed_actual = apply_reference(
            &OpcPackage::new(),
            &changed,
            &reference,
            &mut BTreeSet::new(),
            &mut BTreeSet::new(),
            &mut Vec::new(),
        )
        .unwrap();
        let mut issues = Vec::new();
        assert!(!compare_digest(&reference, changed_actual, &mut issues));
        assert!(matches!(
            issues.as_slice(),
            [SignatureIssue::ChangedReference { uri, .. }] if uri == "#pi-object"
        ));

        let mut escaped = String::new();
        escape_processing_instruction(" data\rmore", &mut escaped);
        assert_eq!(escaped, " data&#xD;more");
    }

    #[test]
    fn relationship_transform_selects_exact_ids_in_canonical_order() {
        let mut package = OpcPackage::new();
        package.set_part("/a.xml", Vec::new());
        package.set_part("/b.xml", Vec::new());
        let relationships = Relationships::from_xml(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="urn:b" Target="b.xml"/><Relationship Id="rId1" Type="urn:a" Target="a.xml"/></Relationships>"#,
        )
        .unwrap();
        let mut covered = BTreeSet::new();
        let mut issues = Vec::new();
        let transformed = relationship_transform(
            &package,
            "/",
            &relationships,
            &["rId2".to_string(), "rId1".to_string()],
            &mut covered,
            &mut issues,
        )
        .unwrap();
        assert_eq!(
            String::from_utf8(transformed).unwrap(),
            format!(
                "<Relationships xmlns=\"{RELATIONSHIPS_NS}\"><Relationship Id=\"rId1\" Target=\"a.xml\" TargetMode=\"Internal\" Type=\"urn:a\"></Relationship><Relationship Id=\"rId2\" Target=\"b.xml\" TargetMode=\"Internal\" Type=\"urn:b\"></Relationship></Relationships>"
            )
        );
        assert!(issues.is_empty());

        let mut duplicate = Relationships::new();
        duplicate.items.push(Relationship {
            id: "rId1".to_owned(),
            rel_type: "urn:a".to_owned(),
            target: "a.xml".to_owned(),
            target_mode: None,
        });
        duplicate.items.push(Relationship {
            id: "rId1".to_owned(),
            rel_type: "urn:b".to_owned(),
            target: "b.xml".to_owned(),
            target_mode: None,
        });
        let mut duplicate_issues = Vec::new();
        assert!(
            relationship_transform(
                &package,
                "/",
                &duplicate,
                &["rId1".to_string()],
                &mut BTreeSet::new(),
                &mut duplicate_issues,
            )
            .is_none()
        );
        assert!(matches!(
            duplicate_issues.as_slice(),
            [SignatureIssue::DuplicateRelationship { .. }]
        ));

        let external = Relationships::from_xml(
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:a" Target="https://example.invalid" TargetMode="External"/></Relationships>"#,
        )
        .unwrap();
        let mut external_issues = Vec::new();
        assert!(
            relationship_transform(
                &package,
                "/",
                &external,
                &["rId1".to_string()],
                &mut BTreeSet::new(),
                &mut external_issues,
            )
            .is_none()
        );
        assert!(matches!(
            external_issues.as_slice(),
            [SignatureIssue::ExternalRelationship { .. }]
        ));
    }

    #[test]
    fn valid_signature_reports_complete_declared_coverage() {
        let report = signed_package("ds").verify_signatures().unwrap().remove(0);
        assert!(report.cryptographically_valid);
        assert!(report.coverage_complete);
        assert!(report.issues.is_empty());
        assert_eq!(
            report.covered_parts,
            vec!["/[Content_Types].xml", "/word/document.xml"]
        );
        assert_eq!(
            report.covered_relationships,
            vec![CoveredRelationship {
                source_part: "/".to_string(),
                relationship_id: "rId1".to_string(),
            }]
        );
        let signer = report.signer.unwrap();
        assert!(signer.subject.contains("F-171 Test Signer"));

        let mut relocated = signed_package("ds");
        let origin = relocated
            .parts
            .remove("/_xmlsignatures/origin.sigs")
            .unwrap();
        let signature = relocated.parts.remove("/_xmlsignatures/sig1.xml").unwrap();
        let mut origin_rels = relocated
            .part_rels
            .remove("/_xmlsignatures/origin.sigs")
            .unwrap();
        origin_rels.items[0].target = "signature.xml".to_string();
        relocated.set_part("/custom/security/origin.bin", origin);
        relocated.set_part("/custom/security/signature.xml", signature);
        relocated
            .part_rels
            .insert("/custom/security/origin.bin".to_string(), origin_rels);
        relocated
            .package_rels
            .items
            .iter_mut()
            .find(|relationship| relationship.rel_type == rel_types::DIGITAL_SIGNATURE_ORIGIN)
            .unwrap()
            .target = "custom/security/origin.bin".to_string();
        let relocated_report = relocated.verify_signatures().unwrap().remove(0);
        assert_eq!(
            relocated_report.signature_part,
            "/custom/security/signature.xml"
        );
        assert!(relocated_report.cryptographically_valid);
        assert!(relocated_report.coverage_complete);
    }

    #[test]
    fn signature_discovery_references_and_coverage_resolve_case_equivalent_parts() {
        let mut package = signed_package("ds");
        for (original, case_variant) in [
            ("/word/document.xml", "/WORD/DOCUMENT.XML"),
            ("/_xmlsignatures/origin.sigs", "/_XMLSIGNATURES/ORIGIN.SIGS"),
            ("/_xmlsignatures/sig1.xml", "/_XMLSIGNATURES/SIG1.XML"),
        ] {
            let bytes = package.parts.remove(original).unwrap();
            package.parts.insert(case_variant.to_owned(), bytes);
        }
        let origin_relationships = package
            .part_rels
            .remove("/_xmlsignatures/origin.sigs")
            .unwrap();
        package.part_rels.insert(
            "/_XMLSIGNATURES/ORIGIN.SIGS".to_owned(),
            origin_relationships,
        );

        let report = package.verify_signatures().unwrap().remove(0);
        assert_eq!(report.signature_part, "/_XMLSIGNATURES/SIG1.XML");
        assert!(report.cryptographically_valid, "{:?}", report.issues);
        assert!(report.coverage_complete, "{:?}", report.issues);
        assert!(
            report
                .covered_parts
                .contains(&"/WORD/DOCUMENT.XML".to_owned())
        );
    }

    #[test]
    fn modified_signed_part_is_named() {
        let mut package = signed_package("ds");
        package.set_part("/word/document.xml", b"changed".to_vec());
        let report = package.verify_signatures().unwrap().remove(0);
        assert!(!report.cryptographically_valid);
        assert!(report.coverage_complete);
        assert!(matches!(
            report.issues.as_slice(),
            [SignatureIssue::ChangedReference { uri, .. }] if uri == "/word/document.xml"
        ));
    }

    #[test]
    fn partial_or_malformed_coverage_never_reports_success() {
        let mut uncovered = signed_package("ds");
        uncovered.set_part("/word/uncovered.xml", Vec::new());
        let report = uncovered.verify_signatures().unwrap().remove(0);
        assert!(report.cryptographically_valid);
        assert!(!report.coverage_complete);
        assert!(report.issues.contains(&SignatureIssue::UncoveredPart {
            part_name: "/word/uncovered.xml".to_string(),
        }));

        let mut missing = signed_package("ds");
        missing.parts.remove("/word/document.xml");
        let report = missing.verify_signatures().unwrap().remove(0);
        assert!(!report.cryptographically_valid);
        assert!(report.issues.iter().any(|issue| matches!(
            issue,
            SignatureIssue::MissingPart { uri } if uri == "/word/document.xml"
        )));

        let mut duplicate = signed_package("ds");
        mutate_signature_xml(&mut duplicate, |xml| {
            let reference_start = xml
                .find("<ds:Reference URI=\"/word/document.xml\"")
                .unwrap();
            let relative_end = xml[reference_start..].find("</ds:Reference>").unwrap();
            let reference_end = reference_start + relative_end + "</ds:Reference>".len();
            let reference = xml[reference_start..reference_end].to_string();
            xml.replacen("</ds:Manifest>", &format!("{reference}</ds:Manifest>"), 1)
        });
        let report = duplicate.verify_signatures().unwrap().remove(0);
        assert!(!report.cryptographically_valid);
        assert!(report.issues.iter().any(|issue| matches!(
            issue,
            SignatureIssue::ChangedReference { uri, .. } if uri == "#idPackageObject"
        )));

        let mut duplicate_id = signed_package("ds");
        mutate_signature_xml(&mut duplicate_id, |xml| {
            xml.replacen(
                "</ds:Signature>",
                "<ds:Object Id=\"idPackageObject\"></ds:Object></ds:Signature>",
                1,
            )
        });
        assert!(matches!(
            duplicate_id.verify_signatures(),
            Err(OpcError::InvalidSignatureXml(message))
                if message.contains("duplicate same-document Id")
        ));

        let mut duplicate_signed_info = signed_package("ds");
        mutate_signature_xml(&mut duplicate_signed_info, |xml| {
            let start = xml.find("<ds:SignedInfo>").unwrap();
            let relative_end = xml[start..].find("</ds:SignedInfo>").unwrap();
            let end = start + relative_end + "</ds:SignedInfo>".len();
            let signed_info = xml[start..end].to_string();
            xml.replacen(&signed_info, &format!("{signed_info}{signed_info}"), 1)
        });
        assert!(matches!(
            duplicate_signed_info.verify_signatures(),
            Err(OpcError::InvalidSignatureXml(message)) if message.contains("duplicate")
        ));

        let mut infrastructure_relationship = signed_package("ds");
        let expected_relationship_id = infrastructure_relationship
            .part_rels
            .get_mut("/_xmlsignatures/origin.sigs")
            .unwrap()
            .add_external("urn:test:unsigned", "https://example.invalid");
        let report = infrastructure_relationship
            .verify_signatures()
            .unwrap()
            .remove(0);
        assert!(report.cryptographically_valid);
        assert!(!report.coverage_complete);
        assert!(report.issues.iter().any(|issue| matches!(
            issue,
            SignatureIssue::UncoveredRelationship {
                source_part,
                relationship_id,
            } if source_part == "/_xmlsignatures/origin.sigs"
                && relationship_id == &expected_relationship_id
        )));

        let mut external = signed_package("ds");
        external.package_rels.items[0].target_mode = Some("External".to_string());
        let report = external.verify_signatures().unwrap().remove(0);
        assert!(!report.cryptographically_valid);
        assert!(report.issues.iter().any(|issue| matches!(
            issue,
            SignatureIssue::ExternalRelationship { relationship_id, .. } if relationship_id == "rId1"
        )));

        let mut absent_target = signed_package("ds");
        absent_target.package_rels.items[0].target = "word/missing.xml".to_string();
        let report = absent_target.verify_signatures().unwrap().remove(0);
        assert!(!report.cryptographically_valid);
        assert!(report.issues.iter().any(|issue| matches!(
            issue,
            SignatureIssue::MissingRelationshipTarget { target, .. } if target == "/word/missing.xml"
        )));

        let mut orphan = signed_package("ds");
        orphan
            .package_rels
            .items
            .retain(|relationship| relationship.rel_type != rel_types::DIGITAL_SIGNATURE_ORIGIN);
        let report = orphan.verify_signatures().unwrap().remove(0);
        assert!(!report.cryptographically_valid);
        assert!(!report.coverage_complete);
        assert!(matches!(
            report.issues.as_slice(),
            [SignatureIssue::OrphanSignaturePart { .. }]
        ));
    }

    #[test]
    fn unsigned_manifest_cannot_complete_partial_coverage() {
        let mut package = signed_package("ds");
        let unsigned_part = b"unsigned but digest-correct".to_vec();
        package.set_part("/word/unsigned.xml", unsigned_part.clone());
        let unsigned_reference = reference_xml("ds", "/word/unsigned.xml", &unsigned_part, "");
        mutate_signature_xml(&mut package, |xml| {
            xml.replacen(
                "</ds:Signature>",
                &format!(
                    "<ds:Object><ds:Manifest>{unsigned_reference}</ds:Manifest></ds:Object></ds:Signature>"
                ),
                1,
            )
        });

        let report = package.verify_signatures().unwrap().remove(0);
        assert!(report.cryptographically_valid);
        assert!(!report.coverage_complete);
        assert!(
            !report
                .covered_parts
                .contains(&"/word/unsigned.xml".to_string())
        );
        assert!(report.issues.contains(&SignatureIssue::UncoveredPart {
            part_name: "/word/unsigned.xml".to_string(),
        }));
    }

    #[test]
    fn verification_does_not_change_package_bytes() {
        let fixture = signed_package("ds");
        let mut source = Cursor::new(Vec::new());
        fixture.write_to(&mut source).unwrap();
        let source_bytes = source.into_inner();
        if let Ok(path) = std::env::var("RDOCX_F171_ORACLE_OUTPUT") {
            std::fs::write(path, &source_bytes).unwrap();
        }
        let package = OpcPackage::from_reader(Cursor::new(source_bytes)).unwrap();
        let signature_before = package.parts["/_xmlsignatures/sig1.xml"].clone();
        let document_before = package.parts["/word/document.xml"].clone();
        let first = package.verify_signatures().unwrap();
        let second = package.verify_signatures().unwrap();
        assert_eq!(first, second);
        let mut saved = Cursor::new(Vec::new());
        package.write_to(&mut saved).unwrap();
        let reopened = OpcPackage::from_reader(Cursor::new(saved.into_inner())).unwrap();
        assert_eq!(reopened.parts["/_xmlsignatures/sig1.xml"], signature_before);
        assert_eq!(reopened.parts["/word/document.xml"], document_before);
        let report = reopened.verify_signatures().unwrap().remove(0);
        assert!(report.cryptographically_valid);
        assert!(report.coverage_complete);
    }

    fn unsigned_package() -> OpcPackage {
        let mut package = OpcPackage::with_main_part(
            "/word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        );
        package.set_part(
            "/word/document.xml",
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec(),
        );
        package
    }

    fn test_private_key() -> Vec<u8> {
        BASE64.decode(TEST_PRIVATE_KEY_PKCS8_BASE64).unwrap()
    }

    fn test_certificate() -> Vec<u8> {
        BASE64.decode(TEST_CERTIFICATE_DER_BASE64).unwrap()
    }

    fn assert_signing_failure_is_atomic(package: &mut OpcPackage) {
        let before = package_bytes(package);
        assert!(
            package
                .sign(&test_private_key(), &test_certificate())
                .is_err()
        );
        assert_eq!(package_bytes(package), before);
    }

    fn package_bytes(package: &OpcPackage) -> Vec<u8> {
        let mut bytes = Cursor::new(Vec::new());
        package.write_to(&mut bytes).unwrap();
        bytes.into_inner()
    }

    fn signed_package(prefix: &str) -> OpcPackage {
        let mut package = OpcPackage::with_main_part(
            "word/document.xml",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        );
        package.set_part(
            "/word/document.xml",
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec(),
        );
        package.set_part("/_xmlsignatures/origin.sigs", Vec::new());
        package.content_types.add_override(
            "/_xmlsignatures/origin.sigs",
            "application/vnd.openxmlformats-package.digital-signature-origin",
        );
        package.content_types.add_override(
            "/_xmlsignatures/sig1.xml",
            "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
        );
        package.package_rels.add(
            rel_types::DIGITAL_SIGNATURE_ORIGIN,
            "_xmlsignatures/origin.sigs",
        );
        package
            .get_or_create_part_rels("/_xmlsignatures/origin.sigs")
            .add(rel_types::DIGITAL_SIGNATURE, "sig1.xml");

        let content_types = package.content_types.to_xml().unwrap();
        let document = package.parts["/word/document.xml"].clone();
        let relationship_bytes = format!(
            "<Relationships xmlns=\"{RELATIONSHIPS_NS}\"><Relationship Id=\"rId1\" Target=\"word/document.xml\" TargetMode=\"Internal\" Type=\"{}\"></Relationship></Relationships>",
            rel_types::DOCUMENT
        );
        let part_references = format!(
            "{}{}{}",
            reference_xml(prefix, "/%5BContent_Types%5D.xml", &content_types, ""),
            reference_xml(prefix, "/word/document.xml", &document, ""),
            reference_xml(
                prefix,
                "/_rels/.rels",
                relationship_bytes.as_bytes(),
                &format!(
                    "<{prefix}:Transforms><{prefix}:Transform Algorithm=\"{RELATIONSHIP_TRANSFORM}\"><mdssi:RelationshipReference SourceId=\"rId1\"></mdssi:RelationshipReference></{prefix}:Transform><{prefix}:Transform Algorithm=\"{C14N_EXCLUSIVE}\"></{prefix}:Transform></{prefix}:Transforms>"
                ),
            ),
        );
        let object = format!(
            "<{prefix}:Object Id=\"idPackageObject\"><{prefix}:Manifest>{part_references}</{prefix}:Manifest></{prefix}:Object>"
        );
        let canonical_object = object
            .replacen(
                &format!("<{prefix}:Object "),
                &format!("<{prefix}:Object xmlns:{prefix}=\"{DSIG_NS}\" "),
                1,
            )
            .replace(
                "<mdssi:RelationshipReference ",
                "<mdssi:RelationshipReference xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\" ",
            );
        let object_digest = BASE64.encode(Sha256::digest(canonical_object.as_bytes()));
        let signed_info = format!(
            "<{prefix}:SignedInfo><{prefix}:CanonicalizationMethod Algorithm=\"{C14N_EXCLUSIVE}\"></{prefix}:CanonicalizationMethod><{prefix}:SignatureMethod Algorithm=\"{SIGNATURE_RSA_SHA256}\"></{prefix}:SignatureMethod><{prefix}:Reference URI=\"#idPackageObject\"><{prefix}:Transforms><{prefix}:Transform Algorithm=\"{C14N_EXCLUSIVE}\"></{prefix}:Transform></{prefix}:Transforms><{prefix}:DigestMethod Algorithm=\"{DIGEST_SHA256}\"></{prefix}:DigestMethod><{prefix}:DigestValue>{object_digest}</{prefix}:DigestValue></{prefix}:Reference></{prefix}:SignedInfo>"
        );
        let signature_value = match prefix {
            "ds" => SIGNATURE_DS_BASE64,
            "sig" => SIGNATURE_SIG_BASE64,
            _ => panic!("fixture has no precomputed signature for prefix {prefix}"),
        };
        let xml = format!(
            "<?xml version=\"1.0\"?><{prefix}:Signature xmlns:{prefix}=\"{DSIG_NS}\" xmlns:mdssi=\"http://schemas.openxmlformats.org/package/2006/digital-signature\">{signed_info}<{prefix}:SignatureValue>{signature_value}</{prefix}:SignatureValue><{prefix}:KeyInfo><{prefix}:X509Data><{prefix}:X509Certificate>{CERTIFICATE_DER_BASE64}</{prefix}:X509Certificate></{prefix}:X509Data></{prefix}:KeyInfo>{object}</{prefix}:Signature>"
        );
        package.set_part("/_xmlsignatures/sig1.xml", xml.into_bytes());
        package
    }

    fn reference_xml(prefix: &str, uri: &str, bytes: &[u8], transforms: &str) -> String {
        let digest = BASE64.encode(Sha256::digest(bytes));
        format!(
            "<{prefix}:Reference URI=\"{uri}\">{transforms}<{prefix}:DigestMethod Algorithm=\"{DIGEST_SHA256}\"></{prefix}:DigestMethod><{prefix}:DigestValue>{digest}</{prefix}:DigestValue></{prefix}:Reference>"
        )
    }

    fn mutate_signature_xml(package: &mut OpcPackage, mutate: impl FnOnce(String) -> String) {
        let xml = String::from_utf8(package.parts["/_xmlsignatures/sig1.xml"].to_vec()).unwrap();
        package.set_part("/_xmlsignatures/sig1.xml", mutate(xml).into_bytes());
    }
}
