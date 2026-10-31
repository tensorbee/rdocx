#!/usr/bin/env python3
"""Validate public from-scratch DOCX authoring and private local evidence."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
from pathlib import Path, PurePosixPath
import re
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
import tomllib
import unittest
from unittest import mock
import xml.etree.ElementTree as ET
import zipfile

from golden_png_harness import decode_png
from pptx_ssim_harness import SOFFICE, assert_tool_versions, structural_similarity


REPO_ROOT = Path(__file__).resolve().parent.parent
PRIVATE_ROOT = REPO_ROOT / "corpus" / "private-docx"
PRIVATE_MANIFEST = PRIVATE_ROOT / "manifest.json"
PRIVATE_EVIDENCE = PRIVATE_ROOT / "evidence"
PRIVATE_GENERATORS = PRIVATE_ROOT / "generators"
PUBLIC_DPI = 150
PRIVATE_CASES = ("P1", "P2", "P3", "P4", "P5")
MAX_PACKAGE_MEMBERS = 2048
MAX_EXPANDED_BYTES = 256 * 1024 * 1024
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WORD_MAIN_CONTENT_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.document.main+xml",
    "docm": "application/vnd.ms-word.document.macroEnabled.main+xml",
    "dotx": "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.template.main+xml",
    "dotm": "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
}
PRIVATE_SKIP = "private conformance: skipped (local inputs unavailable)"

PUBLIC_CONSUMER_SOURCE = r'''use rdocx::{
    BodyItemRef, Document, WordCreationProfile, WordPackageClass,
};
use std::env;
use std::fs;
use std::path::PathBuf;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    eprintln!("public authoring checkpoint: main-entry");
    let mut arguments = env::args_os().skip(1);
    let output = PathBuf::from(arguments.next().ok_or("missing output directory")?);
    if arguments.next().is_some() {
        return Err("unexpected input argument".into());
    }
    fs::create_dir_all(&output)?;

    for (extension, class) in [
        ("docx", WordPackageClass::Document),
        ("docm", WordPackageClass::MacroEnabledDocument),
        ("dotx", WordPackageClass::Template),
        ("dotm", WordPackageClass::MacroEnabledTemplate),
    ] {
        let path = output.join(format!("profile.{extension}"));
        let mut profile = Document::new_with_profile(
            WordCreationProfile::WordCompatible(class),
        );
        profile.save(&path)?;
        if Document::open(&path)?.package_class()? != class {
            return Err(format!("{extension} profile changed package class").into());
        }
    }

    let authored_path = output.join("authored.docx");
    let reopened_path = output.join("reopened.docx");
    let first_png = output.join("render-first.png");
    let second_png = output.join("render-second.png");
    let report_path = output.join("report.txt");

    let mut document = Document::new();
    document.set_title("Public Authoring Conformance Fixture");
    document.set_subject("Public facade package and render validation");
    document.set_header("Public conformance header");
    document.set_footer("Public conformance footer");
    document.add_paragraph("Public authoring starts from a blank document.");
    document.add_numbered_list_item("First modeled item", 0);
    document.add_numbered_list_item("Second modeled item", 0);
    let mut table = document.add_table(2, 2);
    table.cell(0, 0).ok_or("missing table cell")?.set_text("Capability");
    table.cell(0, 1).ok_or("missing table cell")?.set_text("Status");
    table.cell(1, 0).ok_or("missing table cell")?.set_text("Public facade");
    table.cell(1, 1).ok_or("missing table cell")?.set_text("Pass");
    document.save(&authored_path)?;

    let mut reopened = Document::open(&authored_path)?;
    if reopened.title() != Some("Public Authoring Conformance Fixture") {
        return Err("title did not survive reopen".into());
    }
    if reopened.subject() != Some("Public facade package and render validation")
        || reopened.header_text().as_deref() != Some("Public conformance header")
        || reopened.footer_text().as_deref() != Some("Public conformance footer")
    {
        return Err("package metadata or stories did not survive reopen".into());
    }
    if reopened.table_count() != 1 || !reopened.text().contains("Public facade") {
        return Err("modeled state did not survive reopen".into());
    }
    let numbered = reopened
        .paragraphs()
        .iter()
        .filter(|paragraph| paragraph.numbering().is_some())
        .count();
    if numbered != 2 {
        return Err("numbering did not survive reopen".into());
    }
    let unsupported = reopened
        .body_items()
        .filter(|item| matches!(item, BodyItemRef::UnsupportedXml(_)))
        .count();
    eprintln!("public authoring checkpoint: before-render-first");
    let first = reopened
        .render_page_to_png_deterministic(0, 150.0)?
        .ok_or("fixture rendered no first page")?;
    eprintln!("public authoring checkpoint: before-render-second");
    let second = reopened
        .render_page_to_png_deterministic(0, 150.0)?
        .ok_or("fixture rendered no repeated first page")?;
    if first != second {
        return Err("deterministic render changed between calls".into());
    }
    fs::write(&first_png, &first)?;
    fs::write(&second_png, &second)?;
    reopened.save(&reopened_path)?;
    let final_document = Document::open(&reopened_path)?;
    if final_document.title() != Some("Public Authoring Conformance Fixture")
        || final_document.table_count() != 1
        || final_document.text() != reopened.text()
    {
        return Err("save and reopen changed modeled state".into());
    }
    fs::write(
        report_path,
        format!(
            "paragraphs={}\ntables={}\nnumbered={}\nheader_footer={}\nunsupported={}\ndpi=150\n",
            final_document.paragraph_count(),
            final_document.table_count(),
            numbered,
            usize::from(final_document.has_header_footer_content()),
            unsupported,
        ),
    )?;
    Ok(())
}
'''


class ConformanceError(ValueError):
    """A sanitized conformance failure safe to print."""


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def safe_member_name(name: str) -> None:
    path = PurePosixPath(name)
    if (
        not name
        or name.startswith("/")
        or "\\" in name
        or any(part in ("", ".", "..") for part in path.parts)
    ):
        raise ConformanceError("package contains an unsafe member path")


def package_members(path: Path) -> dict[str, bytes]:
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if not infos or len(infos) > MAX_PACKAGE_MEMBERS:
                raise ConformanceError("package member inventory is invalid")
            names: set[str] = set()
            expanded = 0
            members: dict[str, bytes] = {}
            for info in infos:
                safe_member_name(info.filename)
                if info.is_dir() or info.filename in names:
                    raise ConformanceError("package member inventory is invalid")
                names.add(info.filename)
                expanded += info.file_size
                if expanded > MAX_EXPANDED_BYTES:
                    raise ConformanceError("package expansion exceeds the safety ceiling")
                members[info.filename] = archive.read(info)
    except (OSError, zipfile.BadZipFile, RuntimeError) as error:
        raise ConformanceError("DOCX package cannot be read") from error
    return members


def parse_xml(data: bytes, label: str) -> ET.Element:
    try:
        return ET.fromstring(data)
    except ET.ParseError as error:
        raise ConformanceError(f"{label} is not well-formed XML") from error


def content_type_map(members: dict[str, bytes]) -> dict[str, str]:
    data = members.get("[Content_Types].xml")
    if data is None:
        raise ConformanceError("package has no content type manifest")
    root = parse_xml(data, "content type manifest")
    if root.tag != f"{{{CONTENT_TYPES_NS}}}Types":
        raise ConformanceError("content type manifest has the wrong root")
    defaults: dict[str, str] = {}
    overrides: dict[str, str] = {}
    for child in root:
        if child.tag == f"{{{CONTENT_TYPES_NS}}}Default":
            extension = child.attrib.get("Extension", "").lower()
            content_type = child.attrib.get("ContentType", "")
            if not extension or not content_type or extension in defaults:
                raise ConformanceError("content type defaults are invalid")
            defaults[extension] = content_type
        elif child.tag == f"{{{CONTENT_TYPES_NS}}}Override":
            part = child.attrib.get("PartName", "")
            content_type = child.attrib.get("ContentType", "")
            if not part.startswith("/") or not content_type or part in overrides:
                raise ConformanceError("content type overrides are invalid")
            overrides[part] = content_type
        else:
            raise ConformanceError("content type manifest contains an unknown child")
    mapped: dict[str, str] = {}
    for name in members:
        if name == "[Content_Types].xml":
            continue
        part = f"/{name}"
        if part in overrides:
            mapped[part] = overrides[part]
            continue
        extension = (
            "rels"
            if name.endswith(".rels")
            else PurePosixPath(name).suffix.removeprefix(".").lower()
        )
        if extension not in defaults:
            raise ConformanceError("package part has no declared content type")
        mapped[part] = defaults[extension]
    return mapped


def relationship_records(
    members: dict[str, bytes],
) -> tuple[tuple[str, str, str, str, str], ...]:
    records = []
    for name, data in members.items():
        if not name.endswith(".rels"):
            continue
        root = parse_xml(data, "relationship part")
        if root.tag != f"{{{RELATIONSHIPS_NS}}}Relationships":
            raise ConformanceError("relationship part has the wrong root")
        seen: set[str] = set()
        for child in root:
            if child.tag != f"{{{RELATIONSHIPS_NS}}}Relationship":
                raise ConformanceError("relationship part contains an unknown child")
            relationship_id = child.attrib.get("Id", "")
            relationship_type = child.attrib.get("Type", "")
            target = child.attrib.get("Target", "")
            target_mode = child.attrib.get("TargetMode", "")
            if (
                not relationship_id
                or relationship_id in seen
                or not relationship_type
                or not target
            ):
                raise ConformanceError("relationship inventory is invalid")
            if target_mode not in ("", "External"):
                raise ConformanceError("relationship target mode is invalid")
            if target_mode != "External" and "\\" in target:
                raise ConformanceError("relationship target is unsafe")
            seen.add(relationship_id)
            records.append(
                (name, relationship_id, relationship_type, target, target_mode)
            )
    return tuple(sorted(records))


def validate_document_order(members: dict[str, bytes]) -> None:
    data = members.get("word/document.xml")
    if data is None:
        raise ConformanceError("package has no main document part")
    root = parse_xml(data, "main document part")
    if root.tag != f"{{{WORD_NS}}}document":
        raise ConformanceError("main document part has the wrong root")
    body = root.find(f"{{{WORD_NS}}}body")
    if body is None:
        raise ConformanceError("main document part has no body")
    section_positions = [
        index
        for index, child in enumerate(body)
        if child.tag == f"{{{WORD_NS}}}sectPr"
    ]
    if len(section_positions) != 1 or section_positions[0] != len(body) - 1:
        raise ConformanceError("main document children violate schema order")


def normalized_package_graph(
    path: Path, expected_main_content_type: str = WORD_MAIN_CONTENT_TYPES["docx"],
) -> tuple[
    tuple[tuple[str, str], ...],
    tuple[tuple[str, str, str, str, str], ...],
]:
    members = package_members(path)
    types = content_type_map(members)
    relationships = relationship_records(members)
    validate_document_order(members)
    if types.get("/word/document.xml") != expected_main_content_type:
        raise ConformanceError("main document content type is invalid")
    return tuple(sorted(types.items())), relationships


def validate_generated_package(path: Path) -> None:
    members = package_members(path)
    _, relationships = normalized_package_graph(path)
    office_document_targets = []
    for relationship_part, _, relationship_type, target, target_mode in relationships:
        if target_mode == "External":
            continue
        if relationship_part == "_rels/.rels":
            source_directory = ""
        else:
            relationship_path = PurePosixPath(relationship_part)
            source_directory = relationship_path.parent.parent.as_posix()
        resolved = posixpath.normpath(posixpath.join(source_directory, target)).lstrip("/")
        if resolved == ".." or resolved.startswith("../") or resolved not in members:
            raise ConformanceError("generated package has a missing relationship target")
        if (
            relationship_part == "_rels/.rels"
            and relationship_type
            == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        ):
            office_document_targets.append(resolved)
    if office_document_targets != ["word/document.xml"]:
        raise ConformanceError("generated package has an invalid office document relationship")


def validate_word_compatible_profile(path: Path, extension: str) -> None:
    members = package_members(path)
    expected_members = {
        "[Content_Types].xml",
        "_rels/.rels",
        "docProps/app.xml",
        "docProps/core.xml",
        "word/_rels/document.xml.rels",
        "word/document.xml",
        "word/fontTable.xml",
        "word/settings.xml",
        "word/styles.xml",
        "word/theme/theme1.xml",
    }
    if set(members) != expected_members:
        raise ConformanceError("fresh profile part inventory is incomplete")
    content_types, relationships = normalized_package_graph(
        path, WORD_MAIN_CONTENT_TYPES[extension]
    )
    required_content_types = {
        "/docProps/app.xml": "application/vnd.openxmlformats-officedocument."
        "extended-properties+xml",
        "/docProps/core.xml": "application/vnd.openxmlformats-package."
        "core-properties+xml",
        "/word/fontTable.xml": "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.fontTable+xml",
        "/word/settings.xml": "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.settings+xml",
        "/word/styles.xml": "application/vnd.openxmlformats-officedocument."
        "wordprocessingml.styles+xml",
        "/word/theme/theme1.xml": "application/vnd.openxmlformats-officedocument."
        "theme+xml",
    }
    mapped = dict(content_types)
    if any(mapped.get(part) != kind for part, kind in required_content_types.items()):
        raise ConformanceError("fresh profile content types are incomplete")
    expected_relationships = {
        (
            "_rels/.rels",
            "rId1",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
            "word/document.xml",
            "",
        ),
        (
            "_rels/.rels",
            "rId2",
            "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
            "docProps/core.xml",
            "",
        ),
        (
            "_rels/.rels",
            "rId3",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
            "docProps/app.xml",
            "",
        ),
        (
            "word/_rels/document.xml.rels",
            "rId0",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            "styles.xml",
            "",
        ),
        (
            "word/_rels/document.xml.rels",
            "rdocxSettings",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
            "settings.xml",
            "",
        ),
        (
            "word/_rels/document.xml.rels",
            "rdocxTheme",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            "theme/theme1.xml",
            "",
        ),
        (
            "word/_rels/document.xml.rels",
            "rdocxFontTable",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable",
            "fontTable.xml",
            "",
        ),
    }
    if set(relationships) != expected_relationships:
        raise ConformanceError("fresh profile relationship graph is incomplete")
    core = members["docProps/core.xml"]
    if b"dcterms:created" in core or b"dcterms:modified" in core:
        raise ConformanceError("fresh profile invented timestamps")
    if any(b"vbaProject" in data for data in members.values()):
        raise ConformanceError("fresh profile invented a VBA project")


def assert_preserved_opaque_parts(before: Path, after: Path) -> None:
    first = package_members(before)
    second = package_members(after)
    mutable = {
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
        "word/_rels/document.xml.rels",
        "docProps/core.xml",
    }
    opaque = sorted(set(first) - mutable)
    if not opaque:
        raise ConformanceError("public package has no preservation probe")
    if any(name not in second or first[name] != second[name] for name in opaque):
        raise ConformanceError("save and reopen changed an unmodeled package part")


def public_consumer_manifest() -> str:
    crate = (REPO_ROOT / "crates" / "rdocx").as_posix()
    return (
        "[package]\n"
        'name = "rdocx-authoring-conformance"\n'
        'version = "0.0.0"\n'
        'edition = "2024"\n'
        "publish = false\n\n"
        "[dependencies]\n"
        f'rdocx = {{ path = "{crate}" }}\n'
    )


def validate_public_boundary(manifest: str, source: str) -> None:
    try:
        parsed = tomllib.loads(manifest)
    except tomllib.TOMLDecodeError as error:
        raise ConformanceError("public consumer manifest is invalid") from error
    dependencies = parsed.get("dependencies")
    if not isinstance(dependencies, dict):
        raise ConformanceError("public consumer has no dependency boundary")
    rdocx_dependency = dependencies.get("rdocx")
    if (
        set(dependencies) != {"rdocx"}
        or not isinstance(rdocx_dependency, dict)
        or set(rdocx_dependency) != {"path"}
        or rdocx_dependency["path"]
        != (REPO_ROOT / "crates" / "rdocx").as_posix()
    ):
        raise ConformanceError("public consumer depends on a private crate")
    forbidden = (
        "rdocx_oxml",
        "rdocx-oxml",
        "oxml_",
        "oxml-",
        "OpcPackage",
        "set_raw_",
        "RawXml",
        "include_bytes!",
        "include_str!",
        "Document::from_bytes",
    )
    if any(token in source for token in forbidden):
        raise ConformanceError("public fixture bypasses the rdocx authoring boundary")
    if source.count("Document::new()") != 1:
        raise ConformanceError("public fixture must start from Document::new()")
    if "Document::open(&authored_path)" not in source:
        raise ConformanceError("public fixture does not reopen its own output")


def validate_private_generator_boundary(manifest: str, source: str) -> None:
    try:
        parsed = tomllib.loads(manifest)
    except tomllib.TOMLDecodeError as error:
        raise ConformanceError("private generator manifest is invalid") from error
    dependencies = parsed.get("dependencies")
    rdocx_dependency = dependencies.get("rdocx") if isinstance(dependencies, dict) else None
    if (
        not isinstance(dependencies, dict)
        or set(dependencies) != {"rdocx"}
        or not isinstance(rdocx_dependency, dict)
        or set(rdocx_dependency) != {"path"}
        or rdocx_dependency["path"]
        != (REPO_ROOT / "crates" / "rdocx").as_posix()
    ):
        raise ConformanceError("private generator depends on a private crate")
    forbidden = (
        "rdocx_oxml",
        "rdocx-oxml",
        "oxml_",
        "oxml-",
        "OpcPackage",
        "set_raw_",
        "RawXml",
        "include_bytes!",
        "Document::from_bytes",
        "Document::open_mhtml",
        "Document::from_mhtml_bytes",
        "Document::open_html",
        "Document::from_html",
        "insert_html_fragment",
        "insert_html_fragment_with_resources",
        "reference-",
        "candidate-",
        "manifest.json",
        "private-docx",
        "std::process",
        "process::Command",
        "Command::new",
        "std::fs::read",
        "fs::read",
        "File::open",
        "OpenOptions",
        ".read_to_end(",
        ".read_to_string(",
        "std::fs::copy",
        "fs::copy",
        "std::fs::rename",
        "fs::rename",
        "hard_link",
        "symlink",
        "std::net",
        'extern "C"',
        "#[link",
        '"soffice"',
        '"libreoffice"',
        '"pandoc"',
    )
    if any(token in source for token in forbidden):
        raise ConformanceError("private generator bypasses the rdocx authoring boundary")
    if source.count("Document::new()") != 1:
        raise ConformanceError("private generator must start from Document::new()")
    if source.count("Document::open(") != 1 or "Document::open(&output)" not in source:
        raise ConformanceError("private generator must reopen only its own output")


def parse_public_report(path: Path) -> dict[str, int]:
    try:
        values = {
            key: int(value)
            for key, value in (
                line.split("=", 1)
                for line in path.read_text(encoding="utf-8").splitlines()
            )
        }
    except (OSError, ValueError) as error:
        raise ConformanceError("public consumer report is invalid") from error
    if (
        values.get("paragraphs", 0) < 3
        or values.get("tables") != 1
        or values.get("numbered") != 2
        or values.get("header_footer") != 1
    ):
        raise ConformanceError("public modeled-state report is incomplete")
    if values.get("unsupported") != 0 or values.get("dpi") != PUBLIC_DPI:
        raise ConformanceError("public diagnostics or render resolution is invalid")
    return values


def public_consumer_failure(
    result: subprocess.CompletedProcess[str] | subprocess.TimeoutExpired,
    arguments: tuple[str, ...],
    cwd: Path,
    environment: dict[str, str],
) -> str:
    """Keep exact public child diagnostics before its temporary project is removed."""
    timed_out = isinstance(result, subprocess.TimeoutExpired)
    command = result.cmd if timed_out else result.args
    configuration = {
        name: environment.get(name)
        for name in (
            "CARGO_TARGET_DIR", "CARGO_BUILD_TARGET", "CARGO_BUILD_TARGET_DIR",
            "CARGO_BUILD_BUILD_DIR", "CARGO_HOME", "RUSTUP_HOME", "RUSTUP_TOOLCHAIN",
            "RUSTC", "RUSTDOC", "RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER",
            "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "CARGO_BUILD_RUSTFLAGS",
            "RUSTDOCFLAGS", "CARGO_ENCODED_RUSTDOCFLAGS", "CARGO_BUILD_RUSTDOCFLAGS",
        )
    }
    for setting in (
        "OPT_LEVEL", "DEBUG", "SPLIT_DEBUGINFO", "STRIP", "DEBUG_ASSERTIONS",
        "OVERFLOW_CHECKS", "LTO", "PANIC", "INCREMENTAL", "CODEGEN_UNITS", "RPATH",
    ):
        for prefix in ("CARGO_PROFILE_DEV_", "CARGO_PROFILE_DEV_BUILD_OVERRIDE_"):
            name = prefix + setting
            configuration[name] = environment.get(name)
    fields = {
        "argv": arguments,
        "command": command,
        "timeout" if timed_out else "returncode": (
            result.timeout if timed_out else result.returncode
        ),
        "stdout": result.stdout,
        "stderr": result.stderr,
    }
    entry_marker = "public authoring checkpoint: main-entry"
    if isinstance(result.stderr, bytes):
        entry_observed = entry_marker.encode("ascii") in result.stderr
    else:
        entry_observed = isinstance(result.stderr, str) and entry_marker in result.stderr
    return "\n".join(
        (
            "public rdocx-only consumer timed out"
            if timed_out else "public rdocx-only consumer failed",
            "phase='public consumer cargo run'",
            f"cwd={str(cwd)!r}",
            f"host_platform={sys.platform!r}",
            "command_profile='dev'",
            "compiler_version='not observed at this call site'",
            f"inherited_configuration={configuration!r} "
            "(None means unspecified in child environment)",
            "client entry observed in captured stderr"
            if entry_observed else "client entry not observed in captured stderr",
            "wait cause is not inferred from marker presence or absence",
            *(f"{name}({type(value).__name__})={value!r}" for name, value in fields.items()),
        )
    )


def run_public() -> None:
    with TemporaryDirectory(prefix="rdocx-authoring-public-") as directory:
        root = Path(directory)
        source_root = root / "src"
        output = root / "output"
        source_root.mkdir()
        manifest = public_consumer_manifest()
        validate_public_boundary(manifest, PUBLIC_CONSUMER_SOURCE)
        (root / "Cargo.toml").write_text(manifest, encoding="utf-8")
        (source_root / "main.rs").write_text(PUBLIC_CONSUMER_SOURCE, encoding="utf-8")
        environment = os.environ.copy()
        environment.setdefault(
            "CARGO_TARGET_DIR",
            str(REPO_ROOT / "target" / "docx-authoring-conformance"),
        )
        arguments = ("cargo", "run", "--quiet", "--offline", "--", str(output))
        try:
            completed = subprocess.run(
                arguments,
                cwd=root,
                env=environment,
                capture_output=True,
                text=True,
                timeout=300,
            )
        except subprocess.TimeoutExpired as error:
            raise ConformanceError(
                public_consumer_failure(error, arguments, root, environment)
            ) from error
        if completed.returncode != 0:
            raise ConformanceError(
                public_consumer_failure(completed, arguments, root, environment)
            )

        authored = output / "authored.docx"
        reopened = output / "reopened.docx"
        first_png = output / "render-first.png"
        second_png = output / "render-second.png"
        for path in (authored, reopened, first_png, second_png, output / "report.txt"):
            if not path.is_file() or path.stat().st_size == 0:
                raise ConformanceError("public consumer evidence is incomplete")
        for extension in WORD_MAIN_CONTENT_TYPES:
            profile = output / f"profile.{extension}"
            if not profile.is_file() or profile.stat().st_size == 0:
                raise ConformanceError("public profile evidence is incomplete")
            validate_word_compatible_profile(profile, extension)
        if normalized_package_graph(authored) != normalized_package_graph(reopened):
            raise ConformanceError("save and reopen changed the normalized package graph")
        assert_preserved_opaque_parts(authored, reopened)
        parse_public_report(output / "report.txt")
        try:
            first = decode_png(first_png)
            second = decode_png(second_png)
        except ValueError as error:
            raise ConformanceError("public deterministic PNG evidence is invalid") from error
        if first != second or first[0] <= 0 or first[1] <= 0:
            raise ConformanceError("repeated deterministic public render changed")
    print("public authoring conformance: pass")


def git_paths(*arguments: str) -> tuple[str, ...]:
    completed = subprocess.run(
        ("git", *arguments),
        cwd=REPO_ROOT,
        check=True,
        capture_output=True,
        text=True,
        timeout=30,
    )
    return tuple(line for line in completed.stdout.splitlines() if line)


def private_artifact_path(path: str) -> bool:
    lowered = path.lower()
    return (
        lowered.startswith("corpus/private-docx/")
        or lowered.endswith(".docx")
        or lowered.endswith(".docm")
        or lowered.endswith(".dotx")
        or lowered.endswith(".dotm")
        or "private-docx" in lowered
        or "private_docx" in lowered
    )


def reject_tracked_or_staged_private_artifacts(
    tracked: tuple[str, ...] | None = None,
    staged: tuple[str, ...] | None = None,
) -> None:
    tracked = git_paths("ls-files") if tracked is None else tracked
    staged = (
        git_paths("diff", "--cached", "--name-only", "--diff-filter=ACMR")
        if staged is None
        else staged
    )
    if any(private_artifact_path(path) for path in (*tracked, *staged)):
        raise ConformanceError("tracked or staged private artifact detected")


def load_private_manifest(root: Path) -> dict[str, object]:
    path = root / "manifest.json"
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise ConformanceError("private manifest is unavailable or invalid") from error
    if not isinstance(payload, dict) or set(payload) != {
        "schema",
        "dpi",
        "tools",
        "cases",
    }:
        raise ConformanceError("private manifest fields are invalid")
    if payload["schema"] != 1 or payload["dpi"] != PUBLIC_DPI:
        raise ConformanceError("private manifest schema or resolution is invalid")
    tools = payload["tools"]
    if not isinstance(tools, dict) or set(tools) != {"libreoffice", "pdftoppm"}:
        raise ConformanceError("private tool identity fields are invalid")
    cases = payload["cases"]
    if not isinstance(cases, list) or len(cases) != len(PRIVATE_CASES):
        raise ConformanceError("private case inventory is incomplete")
    expected_keys = {
        "id",
        "reference",
        "candidate",
        "reference_sha256",
        "candidate_sha256",
        "minimum_ssim",
        "reference_pages",
        "candidate_pages",
    }
    for index, case in enumerate(cases, 1):
        if not isinstance(case, dict) or set(case) != expected_keys:
            raise ConformanceError("private case fields are invalid")
        alias = PRIVATE_CASES[index - 1]
        if case["id"] != alias:
            raise ConformanceError("private aliases are missing or out of order")
        if case["reference"] != f"reference-{index:02d}.docx":
            raise ConformanceError("private reference aliases are invalid")
        if case["candidate"] != f"candidate-{alias}.docx":
            raise ConformanceError("private candidate aliases are invalid")
        for digest_key in ("reference_sha256", "candidate_sha256"):
            if not isinstance(case[digest_key], str) or not re.fullmatch(
                r"[0-9a-f]{64}", case[digest_key]
            ):
                raise ConformanceError("private digest evidence is invalid")
        threshold = case["minimum_ssim"]
        if type(threshold) not in (int, float) or not 0.0 <= threshold <= 1.0:
            raise ConformanceError("private visual threshold is invalid")
        for page_key in ("reference_pages", "candidate_pages"):
            if type(case[page_key]) is not int or case[page_key] < 1:
                raise ConformanceError("private page expectation is invalid")
    return payload


def private_inventory(root: Path) -> tuple[str, ...]:
    inventory = []
    for path in root.iterdir():
        if path.name in {"evidence", "generators"} and path.is_dir():
            continue
        if not path.is_file():
            raise ConformanceError("private input inventory is missing or unexpected")
        inventory.append(path.name)
    return tuple(sorted(inventory))


def validate_private_inventory(root: Path, payload: dict[str, object]) -> None:
    cases = payload["cases"]
    assert isinstance(cases, list)
    expected = {"manifest.json"}
    for case in cases:
        assert isinstance(case, dict)
        expected.add(str(case["reference"]))
        expected.add(str(case["candidate"]))
    if set(private_inventory(root)) != expected:
        raise ConformanceError("private input inventory is missing or unexpected")
    for case in cases:
        assert isinstance(case, dict)
        reference = root / str(case["reference"])
        candidate = root / str(case["candidate"])
        if sha256(reference) != case["reference_sha256"]:
            raise ConformanceError("private reference identity changed")
        if sha256(candidate) != case["candidate_sha256"]:
            raise ConformanceError("private candidate identity changed")


def generate_private_candidates(
    root: Path, payload: dict[str, object], output_root: Path
) -> dict[str, Path]:
    generator_root = root / "generators"
    expected = {f"{alias}.rs" for alias in PRIVATE_CASES}
    try:
        actual = {
            path.name
            for path in generator_root.iterdir()
            if path.is_file()
        }
    except OSError as error:
        raise ConformanceError("private generator inventory is unavailable") from error
    if actual != expected or any(path.is_dir() for path in generator_root.iterdir()):
        raise ConformanceError("private generator inventory is incomplete")

    manifest = public_consumer_manifest()
    cases = payload["cases"]
    assert isinstance(cases, list)
    generated = {}
    output_root.mkdir(exist_ok=True)
    with TemporaryDirectory(prefix="rdocx-private-build-") as directory:
        temporary = Path(directory)
        environment = os.environ.copy()
        environment.setdefault(
            "CARGO_TARGET_DIR",
            str(REPO_ROOT / "target" / "docx-authoring-conformance"),
        )
        for case in cases:
            assert isinstance(case, dict)
            alias = str(case["id"])
            try:
                source = (generator_root / f"{alias}.rs").read_text(encoding="utf-8")
            except OSError as error:
                raise ConformanceError(f"{alias} private generator is unavailable") from error
            validate_private_generator_boundary(manifest, source)
            project = temporary / alias
            (project / "src").mkdir(parents=True)
            (project / "Cargo.toml").write_text(manifest, encoding="utf-8")
            (project / "src" / "main.rs").write_text(source, encoding="utf-8")
            outputs = []
            for attempt in ("first", "second"):
                output = output_root / f"{alias}-{attempt}.docx"
                completed = subprocess.run(
                    ("cargo", "run", "--quiet", "--offline", "--", str(output)),
                    cwd=project,
                    env=environment,
                    capture_output=True,
                    text=True,
                    timeout=300,
                )
                if completed.returncode != 0 or not output.is_file():
                    raise ConformanceError(f"{alias} private generator failed")
                outputs.append(output)
            if outputs[0].read_bytes() != outputs[1].read_bytes():
                raise ConformanceError(f"{alias} private generator is not deterministic")
            generated[alias] = outputs[0]
    return generated


def render_private_alias(source: Path, alias: str, output: Path) -> list[Path]:
    with TemporaryDirectory(prefix="rdocx-private-render-") as directory:
        temporary = Path(directory)
        generic_input = temporary / f"{alias}.docx"
        render_root = temporary / "rendered"
        shutil.copyfile(source, generic_input)
        completed = subprocess.run(
            (
                "cargo",
                "run",
                "--quiet",
                "--locked",
                "--offline",
                "-p",
                "rdocx-cli",
                "--",
                "render",
                "--dpi",
                str(PUBLIC_DPI),
                "--output-dir",
                str(render_root),
                str(generic_input),
            ),
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            timeout=300,
        )
        if completed.returncode != 0:
            raise ConformanceError(f"{alias} deterministic render failed")
        pages = sorted(
            render_root.glob(f"{alias}_page*.png"),
            key=lambda path: int(path.stem.rsplit("page", 1)[1]),
        )
        if not pages:
            raise ConformanceError(f"{alias} deterministic render produced no pages")
        output.mkdir(parents=True, exist_ok=True)
        published = []
        for index, page in enumerate(pages, 1):
            target = output / f"{alias}-page-{index}.png"
            shutil.copyfile(page, target)
            published.append(target)
        return published


def modeled_projection(source: Path, alias: str) -> dict[str, object]:
    with TemporaryDirectory(prefix="rdocx-private-inspect-") as directory:
        generic_input = Path(directory) / f"{alias}.docx"
        shutil.copyfile(source, generic_input)
        completed = subprocess.run(
            (
                "cargo",
                "run",
                "--quiet",
                "--locked",
                "--offline",
                "-p",
                "rdocx-cli",
                "--",
                "inspect",
                str(generic_input),
                "--json",
            ),
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            timeout=300,
        )
        if completed.returncode != 0:
            raise ConformanceError(f"{alias} modeled inspection failed")
        try:
            projection = json.loads(completed.stdout)
        except json.JSONDecodeError as error:
            raise ConformanceError(f"{alias} modeled inspection is invalid") from error
        if not isinstance(projection, dict) or projection.pop("file", None) is None:
            raise ConformanceError(f"{alias} modeled inspection is incomplete")
        return projection


def libreoffice_repair_evidence(candidate: Path, alias: str) -> bool:
    with TemporaryDirectory(prefix="rdocx-private-libreoffice-") as directory:
        temporary = Path(directory)
        input_root = temporary / "input"
        output_root = temporary / "output"
        profile_root = temporary / "profile"
        input_root.mkdir()
        output_root.mkdir()
        generic_input = input_root / f"{alias}.docx"
        shutil.copyfile(candidate, generic_input)
        profile = f"-env:UserInstallation={profile_root.resolve().as_uri()}"
        completed = subprocess.run(
            (
                SOFFICE,
                "--headless",
                profile,
                "--convert-to",
                "docx",
                "--outdir",
                str(output_root),
                str(generic_input),
            ),
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            timeout=300,
        )
        diagnostics = f"{completed.stdout}\n{completed.stderr}".lower()
        if completed.returncode != 0:
            raise ConformanceError(f"{alias} LibreOffice conversion failed")
        if re.search(r"\b(?:repair(?:ed|ing)?|corrupt(?:ed|ion)?)\b", diagnostics):
            raise ConformanceError(f"{alias} LibreOffice reported package repair")
        reopened = output_root / generic_input.name
        if not reopened.is_file():
            raise ConformanceError(f"{alias} LibreOffice did not reopen the package")
        normalized_package_graph(reopened)
        validate_generated_package(reopened)
    return True


def private_case_evidence(
    root: Path,
    case: dict[str, object],
    evidence: Path,
    generated_candidate: Path | None = None,
) -> dict[str, object]:
    alias = str(case["id"])
    reference = root / str(case["reference"])
    candidate = generated_candidate or root / str(case["candidate"])
    normalized_package_graph(reference)
    validate_generated_package(candidate)
    if modeled_projection(reference, f"{alias}-reference") != modeled_projection(
        candidate, f"{alias}-candidate"
    ):
        raise ConformanceError(f"{alias} modeled projection differs")
    repair_free = libreoffice_repair_evidence(candidate, alias)
    reference_pages = render_private_alias(reference, f"{alias}-reference", evidence)
    candidate_pages = render_private_alias(candidate, f"{alias}-candidate", evidence)
    if len(reference_pages) != case["reference_pages"]:
        raise ConformanceError(f"{alias} reference page count differs")
    if len(candidate_pages) != case["candidate_pages"]:
        raise ConformanceError(f"{alias} candidate page count differs")
    if len(reference_pages) != len(candidate_pages):
        raise ConformanceError(f"{alias} page counts differ")
    scores = []
    dimensions = []
    for reference_page, candidate_page in zip(reference_pages, candidate_pages):
        try:
            first = decode_png(reference_page)
            second = decode_png(candidate_page)
        except ValueError as error:
            raise ConformanceError(f"{alias} rendered PNG evidence is invalid") from error
        if first[:2] != second[:2]:
            raise ConformanceError(f"{alias} rendered page geometry differs")
        score = structural_similarity(first, second)
        if score < float(case["minimum_ssim"]):
            raise ConformanceError(f"{alias} rendered similarity is below its threshold")
        scores.append(round(score, 9))
        dimensions.append([first[0], first[1]])
    return {
        "id": alias,
        "reference_pages": len(reference_pages),
        "candidate_pages": len(candidate_pages),
        "dimensions": dimensions,
        "scores": scores,
        "minimum_ssim": case["minimum_ssim"],
        "libreoffice_repair_free": repair_free,
    }


def run_private(
    required: bool, root: Path = PRIVATE_ROOT, generate: bool = False
) -> bool:
    reject_tracked_or_staged_private_artifacts()
    if not root.is_dir() or not (root / "manifest.json").is_file():
        if required:
            raise ConformanceError("private corpus is unavailable")
        print(PRIVATE_SKIP)
        return False
    payload = load_private_manifest(root)
    validate_private_inventory(root, payload)
    try:
        expected_soffice, expected_pdftoppm = assert_tool_versions()
    except (OSError, subprocess.SubprocessError, ValueError) as error:
        raise ConformanceError("required private renderer versions are unavailable") from error
    tools = payload["tools"]
    assert isinstance(tools, dict)
    if tools["libreoffice"] != expected_soffice or tools["pdftoppm"] != expected_pdftoppm:
        raise ConformanceError("private manifest tool identities changed")
    with TemporaryDirectory(prefix="rdocx-private-generated-") as directory:
        generated = (
            generate_private_candidates(root, payload, Path(directory))
            if generate
            else {}
        )
        evidence = root / "evidence"
        if evidence.exists():
            shutil.rmtree(evidence)
        evidence.mkdir()
        records = []
        cases = payload["cases"]
        assert isinstance(cases, list)
        for case in cases:
            assert isinstance(case, dict)
            alias = str(case["id"])
            records.append(
                private_case_evidence(root, case, evidence, generated.get(alias))
            )
    (evidence / "results.json").write_text(
        json.dumps(
            {
                "schema": 1,
                "dpi": PUBLIC_DPI,
                "tools": tools,
                "cases": records,
            },
            indent=2,
            sort_keys=True,
        )
        + "\n",
        encoding="utf-8",
    )
    print("private authoring conformance: pass (P1 through P5)")
    return True


class HarnessSelfTests(unittest.TestCase):
    def test_public_timeout_retains_typed_child_diagnostics(self) -> None:
        cases = (
            (
                ["cargo", "run", "literal $(no-command); quote'"],
                300, b"prefix\xff\x00\n", "text stderr\n",
            ),
            (("cargo", "run"), 12.5, "text stdout", b"binary stderr\xfe\n"),
            (["cargo"], 300, b"", ""),
            (("cargo",), 300, "", b""),
            (["cargo"], 300, None, None),
            (("cargo",), 300, None, b"public authoring checkpoint: main-entry\n"),
        )
        inherited = {
            "CARGO_TARGET_DIR": "relative-target",
            "RUSTUP_TOOLCHAIN": "1.93-test-fixture",
            "CARGO_BUILD_TARGET": "fixture-target",
            "CARGO_PROFILE_DEV_OPT_LEVEL": "1",
            "CARGO_PROFILE_DEV_DEBUG_ASSERTIONS": "true",
            "CARGO_PROFILE_DEV_OVERFLOW_CHECKS": "true",
            "RUSTFLAGS": "--cfg fixture",
            "UNRELATED_SECRET": "must-not-appear",
        }
        for command, timeout, stdout, stderr in cases:
            with self.subTest(command=command, stdout=stdout, stderr=stderr):
                error = subprocess.TimeoutExpired(
                    command, timeout, output=stdout, stderr=stderr
                )
                with mock.patch.dict(os.environ, inherited, clear=True), mock.patch(
                    f"{__name__}.subprocess.run", side_effect=error
                ) as child:
                    with self.assertRaises(
                        (ConformanceError, subprocess.TimeoutExpired)
                    ) as caught:
                        run_public()
                self.assertIsInstance(caught.exception, ConformanceError)
                diagnostic = str(caught.exception)
                arguments = child.call_args.args[0]
                options = child.call_args.kwargs
                self.assertEqual(child.call_count, 1)
                self.assertEqual(
                    arguments[:-1], ("cargo", "run", "--quiet", "--offline", "--")
                )
                self.assertEqual(arguments[-1], str(options["cwd"] / "output"))
                self.assertEqual(options["timeout"], 300)
                self.assertIs(options["capture_output"], True)
                self.assertIs(options["text"], True)
                self.assertEqual(options["env"], inherited)
                self.assertFalse(options["cwd"].exists())
                for fragment in (
                    "phase='public consumer cargo run'",
                    f"argv(tuple)={arguments!r}",
                    f"command({type(command).__name__})={command!r}",
                    f"timeout({type(timeout).__name__})={timeout!r}",
                    f"stdout({type(stdout).__name__})={stdout!r}",
                    f"stderr({type(stderr).__name__})={stderr!r}",
                    f"cwd={str(options['cwd'])!r}",
                    "'CARGO_TARGET_DIR': 'relative-target'",
                    "'RUSTUP_TOOLCHAIN': '1.93-test-fixture'",
                    "'CARGO_BUILD_TARGET': 'fixture-target'",
                    "'CARGO_PROFILE_DEV_OPT_LEVEL': '1'",
                    "'CARGO_PROFILE_DEV_DEBUG_ASSERTIONS': 'true'",
                    "'CARGO_PROFILE_DEV_OVERFLOW_CHECKS': 'true'",
                    "'RUSTC': None",
                    "'RUSTFLAGS': '--cfg fixture'",
                    "command_profile='dev'",
                    "compiler_version='not observed at this call site'",
                ):
                    self.assertIn(fragment, diagnostic)
                self.assertNotIn("UNRELATED_SECRET", diagnostic)
                self.assertNotIn("must-not-appear", diagnostic)
                self.assertNotIn("compilation proven", diagnostic)
                self.assertNotIn("lock wait proven", diagnostic)
                if stderr and (
                    b"main-entry" in stderr
                    if isinstance(stderr, bytes) else "main-entry" in stderr
                ):
                    self.assertIn("client entry observed in captured stderr", diagnostic)
                else:
                    self.assertIn("client entry not observed in captured stderr", diagnostic)

    def test_public_nonzero_child_retains_streams_and_stays_failure(self) -> None:
        def fail(arguments: tuple[str, ...], **_: object):
            return subprocess.CompletedProcess(arguments, 17, "stdout\n", "stderr\n")

        with mock.patch.dict(os.environ, {}, clear=True), mock.patch(
            f"{__name__}.subprocess.run", side_effect=fail
        ) as child:
            with self.assertRaises(ConformanceError) as caught:
                run_public()
        options = child.call_args.kwargs
        arguments = child.call_args.args[0]
        diagnostic = str(caught.exception)
        self.assertEqual(child.call_count, 1)
        self.assertEqual(options["timeout"], 300)
        self.assertEqual(
            options["env"],
            {"CARGO_TARGET_DIR": str(REPO_ROOT / "target" / "docx-authoring-conformance")},
        )
        for fragment in (
            "phase='public consumer cargo run'", "returncode(int)=17",
            f"command(tuple)={arguments!r}",
            "stdout(str)='stdout\\n'", "stderr(str)='stderr\\n'",
            "'RUSTUP_TOOLCHAIN': None", "'CARGO_PROFILE_DEV_OPT_LEVEL': None",
        ):
            self.assertIn(fragment, diagnostic)
        self.assertFalse(options["cwd"].exists())
        with mock.patch(
            f"{__name__}.parse_arguments",
            return_value=argparse.Namespace(
                self_test=False, public=True,
                private_required=False, private_generate_required=False,
            ),
        ), mock.patch(f"{__name__}.reject_tracked_or_staged_private_artifacts"), mock.patch(
            f"{__name__}.run_public", side_effect=ConformanceError(diagnostic)
        ), mock.patch(f"{__name__}.print") as reported:
            self.assertEqual(main(), 1)
        reported.assert_called_once_with(
            f"docx_authoring_conformance: {diagnostic}", file=sys.stderr
        )

    def test_public_progress_markers_preserve_order_and_render_calls(self) -> None:
        markers = (
            'eprintln!("public authoring checkpoint: main-entry");',
            'eprintln!("public authoring checkpoint: before-render-first");',
            'eprintln!("public authoring checkpoint: before-render-second");',
        )
        for marker in markers:
            self.assertEqual(PUBLIC_CONSUMER_SOURCE.count(marker), 1)
        main_start = PUBLIC_CONSUMER_SOURCE.index("fn main()")
        first_render = PUBLIC_CONSUMER_SOURCE.index("    let first = reopened")
        second_render = PUBLIC_CONSUMER_SOURCE.index("    let second = reopened")
        positions = [PUBLIC_CONSUMER_SOURCE.index(marker) for marker in markers]
        self.assertLess(main_start, positions[0])
        self.assertLess(positions[0], PUBLIC_CONSUMER_SOURCE.index("    let mut arguments"))
        self.assertLess(positions[0], positions[1])
        self.assertLess(positions[1], first_render)
        self.assertLess(first_render, positions[2])
        self.assertLess(positions[2], second_render)
        self.assertEqual(
            PUBLIC_CONSUMER_SOURCE.count(".render_page_to_png_deterministic(0, 150.0)?"), 2
        )
        validate_public_boundary(public_consumer_manifest(), PUBLIC_CONSUMER_SOURCE)

    def test_public_zero_child_without_evidence_stays_failure(self) -> None:
        with mock.patch(
            f"{__name__}.subprocess.run",
            return_value=subprocess.CompletedProcess(("cargo",), 0, "", ""),
        ) as child:
            with self.assertRaisesRegex(
                ConformanceError, "public consumer evidence is incomplete"
            ):
                run_public()
        self.assertEqual(child.call_count, 1)
        self.assertEqual(child.call_args.kwargs["timeout"], 300)

    def test_public_boundary_rejects_private_dependencies_and_raw_bypasses(self) -> None:
        manifest = public_consumer_manifest()
        validate_public_boundary(manifest, PUBLIC_CONSUMER_SOURCE)
        with self.assertRaisesRegex(ConformanceError, "private crate"):
            validate_public_boundary(manifest + "oxml-opc = \"0.1\"\n", PUBLIC_CONSUMER_SOURCE)
        aliased = manifest.replace(
            "rdocx = { path = ",
            'rdocx = { package = "rdocx-oxml", path = ',
        )
        with self.assertRaisesRegex(ConformanceError, "private crate"):
            validate_public_boundary(aliased, PUBLIC_CONSUMER_SOURCE)
        with self.assertRaisesRegex(ConformanceError, "authoring boundary"):
            validate_public_boundary(manifest, PUBLIC_CONSUMER_SOURCE + "\nOpcPackage")
        with self.assertRaisesRegex(ConformanceError, "Document::new"):
            validate_public_boundary(
                manifest, PUBLIC_CONSUMER_SOURCE.replace("Document::new()", "Document::default()")
            )

    def test_private_generator_boundary_requires_new_and_own_output_only(self) -> None:
        manifest = public_consumer_manifest()
        source = """use rdocx::Document;
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::path::PathBuf::from(std::env::args_os().nth(1).unwrap());
    let mut document = Document::new();
    document.save(&output)?;
    Document::open(&output)?;
    Ok(())
}
"""
        validate_private_generator_boundary(manifest, source)
        with self.assertRaisesRegex(ConformanceError, "authoring boundary"):
            validate_private_generator_boundary(
                manifest, source.replace("Document::new()", "Document::from_bytes(&[])?")
            )
        with self.assertRaisesRegex(ConformanceError, "authoring boundary"):
            validate_private_generator_boundary(
                manifest,
                source.replace(
                    "document.save(&output)?;",
                    "document.insert_html_fragment(todo!(), \"<p>shortcut</p>\", &[])?;\n"
                    "    document.save(&output)?;",
                ),
            )
        with self.assertRaisesRegex(ConformanceError, "reopen only"):
            validate_private_generator_boundary(
                manifest, source.replace("Document::open(&output)", "Document::open(\"input.docx\")")
            )
        with self.assertRaisesRegex(ConformanceError, "authoring boundary"):
            validate_private_generator_boundary(
                manifest,
                source.replace(
                    "document.save(&output)?;",
                    'std::process::Command::new("soffice").status()?;\n'
                    "    document.save(&output)?;",
                ),
            )
        with self.assertRaisesRegex(ConformanceError, "authoring boundary"):
            validate_private_generator_boundary(
                manifest,
                source.replace(
                    "document.save(&output)?;",
                    'let _source = std::fs::read("reference.docx")?;\n'
                    "    document.save(&output)?;",
                ),
            )

    def test_private_generation_uses_deterministic_outputs_as_evidence(self) -> None:
        source = """use rdocx::Document;
fn main() -> Result<(), Box<dyn std::error::Error>> {
    let output = std::path::PathBuf::from(std::env::args_os().nth(1).unwrap());
    let mut document = Document::new();
    document.save(&output)?;
    Document::open(&output)?;
    Ok(())
}
"""
        with TemporaryDirectory(prefix="rdocx-private-generator-test-") as directory:
            root = Path(directory)
            generators = root / "generators"
            generators.mkdir()
            for alias in PRIVATE_CASES:
                (generators / f"{alias}.rs").write_text(source, encoding="utf-8")
            payload = {"cases": [{"id": alias} for alias in PRIVATE_CASES]}

            def write_generated_output(arguments: tuple[str, ...], **_: object):
                Path(arguments[-1]).write_bytes(b"generated-from-public-api")
                return subprocess.CompletedProcess(arguments, 0)

            with mock.patch(
                f"{__name__}.subprocess.run", side_effect=write_generated_output
            ):
                generated = generate_private_candidates(
                    root, payload, root / "generated"
                )
            self.assertEqual(set(generated), set(PRIVATE_CASES))
            self.assertTrue(
                all(
                    path.read_bytes() == b"generated-from-public-api"
                    for path in generated.values()
                )
            )

    def test_generated_package_requires_resolved_targets_and_one_document_root(self) -> None:
        content_types = f'''<Types xmlns="{CONTENT_TYPES_NS}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="{WORD_MAIN_CONTENT_TYPES['docx']}"/>
</Types>'''.encode()
        document = f'''<w:document xmlns:w="{WORD_NS}"><w:body><w:sectPr/></w:body></w:document>'''.encode()

        def write_package(path: Path, relationship_type: str, target: str) -> None:
            relationships = f'''<Relationships xmlns="{RELATIONSHIPS_NS}">
<Relationship Id="rId1" Type="{relationship_type}" Target="{target}"/>
</Relationships>'''.encode()
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("[Content_Types].xml", content_types)
                archive.writestr("_rels/.rels", relationships)
                archive.writestr("word/document.xml", document)

        office_document = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        )
        with TemporaryDirectory(prefix="rdocx-generated-package-test-") as directory:
            root = Path(directory)
            valid = root / "valid.docx"
            write_package(valid, office_document, "word/document.xml")
            validate_generated_package(valid)

            def write_libreoffice_output(arguments: tuple[str, ...], **_: object):
                output_root = Path(arguments[arguments.index("--outdir") + 1])
                source = Path(arguments[-1])
                shutil.copyfile(source, output_root / source.name)
                return subprocess.CompletedProcess(arguments, 0, "converted", "")

            with mock.patch(
                f"{__name__}.subprocess.run", side_effect=write_libreoffice_output
            ):
                self.assertTrue(libreoffice_repair_evidence(valid, "P1"))
            with mock.patch(
                f"{__name__}.subprocess.run",
                return_value=subprocess.CompletedProcess(
                    (SOFFICE,), 0, "package repaired", ""
                ),
            ):
                with self.assertRaisesRegex(ConformanceError, "reported package repair"):
                    libreoffice_repair_evidence(valid, "P1")

            missing = root / "missing.docx"
            write_package(missing, office_document, "word/missing.xml")
            with self.assertRaisesRegex(ConformanceError, "missing relationship target"):
                validate_generated_package(missing)

            wrong_root = root / "wrong-root.docx"
            write_package(wrong_root, "urn:not-an-office-document", "word/document.xml")
            with self.assertRaisesRegex(ConformanceError, "invalid office document"):
                validate_generated_package(wrong_root)

    def test_normalized_relationship_records_retain_ids(self) -> None:
        first = b'''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="item.xml"/></Relationships>'''
        second = first.replace(b'rId1', b'rId2')
        self.assertNotEqual(
            relationship_records({"_rels/.rels": first}),
            relationship_records({"_rels/.rels": second}),
        )

    def test_private_artifact_scan_rejects_without_disclosing_paths(self) -> None:
        reject_tracked_or_staged_private_artifacts(("src/lib.rs",), ("README.md",))
        secret = "corpus/private-docx/customer-name.docx"
        with self.assertRaises(ConformanceError) as context:
            reject_tracked_or_staged_private_artifacts((secret,), ())
        self.assertNotIn(secret, str(context.exception))
        with mock.patch(
            f"{__name__}.parse_arguments",
            return_value=argparse.Namespace(
                self_test=False, public=False,
                private_required=True, private_generate_required=False,
            ),
        ), mock.patch(
            f"{__name__}.run_private",
            side_effect=subprocess.TimeoutExpired(
                ["private-command", secret], 120, output=secret, stderr=secret
            ),
        ), mock.patch(f"{__name__}.print") as reported:
            self.assertEqual(main(), 1)
        reported.assert_called_once_with(
            "docx_authoring_conformance: external command timed out", file=sys.stderr
        )

    def test_optional_and_required_private_modes_fail_closed(self) -> None:
        with TemporaryDirectory(prefix="rdocx-private-mode-test-") as directory:
            root = Path(directory) / "missing"
            with mock.patch(
                f"{__name__}.reject_tracked_or_staged_private_artifacts"
            ):
                self.assertFalse(run_private(False, root))
                with self.assertRaisesRegex(ConformanceError, "unavailable"):
                    run_private(True, root)

    def test_private_manifest_rejects_missing_extra_changed_and_incomplete_evidence(self) -> None:
        with TemporaryDirectory(prefix="rdocx-private-manifest-test-") as directory:
            root = Path(directory)
            (root / "manifest.json").write_text("[]", encoding="utf-8")
            with self.assertRaisesRegex(ConformanceError, "fields are invalid"):
                load_private_manifest(root)
            payload = {
                "schema": 1,
                "dpi": PUBLIC_DPI,
                "tools": {"libreoffice": "pinned", "pdftoppm": "pinned"},
                "cases": [],
            }
            for index, alias in enumerate(PRIVATE_CASES, 1):
                reference = root / f"reference-{index:02d}.docx"
                candidate = root / f"candidate-{alias}.docx"
                reference.write_bytes(f"reference-{alias}".encode())
                candidate.write_bytes(f"candidate-{alias}".encode())
                payload["cases"].append(
                    {
                        "id": alias,
                        "reference": reference.name,
                        "candidate": candidate.name,
                        "reference_sha256": sha256(reference),
                        "candidate_sha256": sha256(candidate),
                        "minimum_ssim": 0.5 + index / 100,
                        "reference_pages": 1,
                        "candidate_pages": 1,
                    }
                )
            (root / "manifest.json").write_text(json.dumps(payload), encoding="utf-8")
            loaded = load_private_manifest(root)
            validate_private_inventory(root, loaded)

            missing = root / "candidate-P2.docx"
            missing_bytes = missing.read_bytes()
            missing.unlink()
            with self.assertRaisesRegex(ConformanceError, "missing or unexpected"):
                validate_private_inventory(root, loaded)
            missing.write_bytes(missing_bytes)

            (root / "unexpected.bin").write_bytes(b"extra")
            with self.assertRaisesRegex(ConformanceError, "unexpected"):
                validate_private_inventory(root, loaded)
            (root / "unexpected.bin").unlink()

            (root / "candidate-P3.docx").write_bytes(b"changed")
            with self.assertRaisesRegex(ConformanceError, "identity changed"):
                validate_private_inventory(root, loaded)

            incomplete = dict(payload)
            incomplete["cases"] = payload["cases"][:-1]
            (root / "manifest.json").write_text(
                json.dumps(incomplete), encoding="utf-8"
            )
            with self.assertRaisesRegex(ConformanceError, "inventory is incomplete"):
                load_private_manifest(root)


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--self-test", action="store_true")
    mode.add_argument("--public", action="store_true")
    mode.add_argument("--private-required", action="store_true")
    mode.add_argument("--private-generate-required", action="store_true")
    return parser.parse_args()


def main() -> int:
    arguments = parse_arguments()
    try:
        if arguments.self_test:
            suite = unittest.defaultTestLoader.loadTestsFromTestCase(HarnessSelfTests)
            result = unittest.TextTestRunner(verbosity=2).run(suite)
            return 0 if result.wasSuccessful() else 1
        if arguments.public:
            reject_tracked_or_staged_private_artifacts()
            run_public()
            return 0
        if arguments.private_required:
            run_private(True)
            return 0
        if arguments.private_generate_required:
            run_private(True, generate=True)
            return 0
        reject_tracked_or_staged_private_artifacts()
        run_public()
        run_private(False)
        return 0
    except subprocess.TimeoutExpired:
        print("docx_authoring_conformance: external command timed out", file=sys.stderr)
        return 1
    except (ConformanceError, subprocess.CalledProcessError, OSError) as error:
        print(f"docx_authoring_conformance: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
