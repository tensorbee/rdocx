# oxml-opc

Format-neutral Open Packaging Conventions support for ZIP parts, relationships, content types, and package preservation.

## Use it when

Use this crate when implementing an OOXML package reader or writer. Use `rdocx` or `rpptx` for complete document APIs.

## Relationship

This is the shared successor to the deprecated `rdocx-opc` shim and is used by both document families.

## Example

```rust,no_run
use oxml_opc::ContentTypes;

let content_types = ContentTypes::from_xml(br#"
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="xml" ContentType="application/xml"/>
    </Types>
"#)?;
assert_eq!(content_types.defaults["xml"], "application/xml");
# Ok::<(), oxml_opc::OpcError>(())
```

Add `oxml-opc = "0.10.0"` to your dependencies. Start with the [package API documentation](https://docs.rs/oxml-opc).
