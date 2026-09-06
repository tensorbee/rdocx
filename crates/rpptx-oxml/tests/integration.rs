use std::collections::{HashMap, HashSet};
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::process::Command;

#[test]
fn media_model_reads_aliases_and_writes_schema_order_without_losing_raw_siblings() {
    use rpptx_oxml::picture::{MediaKind, MediaSource};
    use rpptx_oxml::timing::{MediaCommandKind, MediaDisplayPolicy, TimingNode, TimingTarget};

    let xml = br#"<q:pic xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:x="urn:f215"><q:nvPicPr><q:cNvPr id="2" name="media"/><q:cNvPicPr/><q:nvPr><x:before-media a='A&#x20;B' r:id="rId2"><a:audioFile r:link="rId99"/></x:before-media><a:audioFile r:link="rId2"/><q:extLst><q:ext uri="{00000000-0000-0000-0000-000000000000}"><x:wrapper><p14:media r:link="rId98"><p14:trim st="1" end="2"/></p14:media></x:wrapper></q:ext><q:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media xmlns:r="urn:shadow" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" rel:embed="rId1" rel:link="rId4"><p14:trim st="875.25" end="125.5"/><x:inside-media/></p14:media><x:metadata>keep &amp; stay</x:metadata></q:ext></q:extLst><x:after-media/></q:nvPr></q:nvPicPr><q:blipFill><a:blip r:embed="rId3"/></q:blipFill><q:spPr/></q:pic>"#;
    let mut picture = rpptx_oxml::picture::CT_Picture::from_xml(xml).unwrap();
    let media = picture.media.as_ref().unwrap();
    assert_eq!(media.kind, MediaKind::Audio);
    assert_eq!(
        media.source,
        MediaSource::Linked {
            relationship_id: "rId4".to_owned()
        }
    );
    assert_eq!(media.poster_relationship_id.as_deref(), Some("rId3"));
    assert_eq!(picture.standard_media_relationship_id(), Some("rId2"));
    assert_eq!(
        picture.office_media_relationship_ids(),
        &["rId1".to_owned(), "rId4".to_owned()]
    );
    assert_eq!(picture.media_trim_bounds(), (Some(875), Some(126)));
    let written = String::from_utf8(picture.to_xml().unwrap()).unwrap();
    assert!(written.starts_with("<p:pic"));
    assert!(
        written.contains(
            "<x:before-media a='A&#x20;B' r:id=\"rId2\"><a:audioFile r:link=\"rId99\"/></x:before-media>"
        )
    );
    assert!(written.contains("<x:metadata>keep &amp; stay</x:metadata>"));
    assert!(written.contains("<x:inside-media/>"));
    assert!(written.contains("<x:after-media/>"));
    assert!(written.contains(
        r#"<p14:media xmlns:r="urn:shadow" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships" rel:embed="rId1" rel:link="rId4"><p14:trim st="875.25" end="125.5"/>"#
    ));
    let unrelated_extension = r#"<q:ext uri="{00000000-0000-0000-0000-000000000000}"><x:wrapper><p14:media r:link="rId98"><p14:trim st="1" end="2"/></p14:media></x:wrapper></q:ext>"#;
    assert!(written.contains(unrelated_extension));
    let reparsed = rpptx_oxml::picture::CT_Picture::from_xml(written.as_bytes()).unwrap();
    assert_eq!(reparsed.media, picture.media);

    picture
        .replace_media_relationship_ids(
            "rId20",
            MediaSource::Embedded {
                relationship_id: "rId21".to_owned(),
            },
        )
        .unwrap();
    let embedded = String::from_utf8(picture.to_xml().unwrap()).unwrap();
    assert!(embedded.contains(r#"<a:audioFile r:link="rId20"/>"#));
    assert!(embedded.contains(r#"rel:embed="rId21">"#));
    assert!(!embedded.contains(r#"rel:link="rId4""#));
    assert!(!embedded.contains(r#"r:embed="rId21""#));
    assert!(embedded.contains(r#"<p14:trim st="875.25" end="125.5"/>"#));
    assert!(embedded.contains(r#"<a:audioFile r:link="rId99"/>"#));
    assert!(embedded.contains(r#"<x:before-media a='A&#x20;B' r:id="rId2">"#));
    assert!(embedded.contains(unrelated_extension));
    assert!(embedded.contains("<x:metadata>keep &amp; stay</x:metadata>"));
    assert_eq!(
        picture.media.as_ref().unwrap().source,
        MediaSource::Embedded {
            relationship_id: "rId21".to_owned()
        }
    );

    picture
        .replace_media_relationship_ids(
            "rId30",
            MediaSource::Linked {
                relationship_id: "rId31".to_owned(),
            },
        )
        .unwrap();
    let linked = String::from_utf8(picture.to_xml().unwrap()).unwrap();
    assert!(linked.contains(r#"<a:audioFile r:link="rId30"/>"#));
    assert!(linked.contains(r#"rel:link="rId31">"#));
    assert!(!linked.contains(r#"rel:embed="rId21""#));
    assert!(!linked.contains(r#"r:link="rId31""#));
    assert!(linked.contains(r#"<a:audioFile r:link="rId99"/>"#));
    assert!(linked.contains(r#"<x:before-media a='A&#x20;B' r:id="rId2">"#));
    assert!(linked.contains("<x:inside-media/>"));
    assert!(linked.contains(unrelated_extension));

    let standard_only_xml = br#"<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:f215"><p:nvPicPr><p:cNvPr id="5" name="standard only"/><p:cNvPicPr/><p:nvPr><a:videoFile r:link="rId40"/><x:keep/></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rId41"/></p:blipFill><p:spPr/></p:pic>"#;
    let mut standard_only = rpptx_oxml::picture::CT_Picture::from_xml(standard_only_xml).unwrap();
    assert!(standard_only.office_media_relationship_ids().is_empty());
    standard_only
        .replace_media_relationship_ids(
            "rId42",
            MediaSource::Embedded {
                relationship_id: "unused-office-id".to_owned(),
            },
        )
        .unwrap();
    let embedded_standard_only = String::from_utf8(standard_only.to_xml().unwrap()).unwrap();
    assert!(embedded_standard_only.contains(r#"<a:videoFile r:link="rId42"/>"#));
    assert!(embedded_standard_only.contains("<x:keep/>"));
    assert!(!embedded_standard_only.contains("p14:media"));
    assert!(!embedded_standard_only.contains("unused-office-id"));
    standard_only
        .replace_media_relationship_ids(
            "rId43",
            MediaSource::Linked {
                relationship_id: "another-unused-office-id".to_owned(),
            },
        )
        .unwrap();
    let linked_standard_only = String::from_utf8(standard_only.to_xml().unwrap()).unwrap();
    assert!(linked_standard_only.contains(r#"<a:videoFile r:link="rId43"/>"#));
    assert!(!linked_standard_only.contains("p14:media"));
    assert!(!linked_standard_only.contains("another-unused-office-id"));

    let timing_xml = br#"<q:timing xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:f215"><q:tnLst><x:before/><q:audio><q:cMediaNode vol="80000"><x:wrapper><x:cTn repeatCount="1" display="1"/></x:wrapper><q:cTn id="7" display="0" repeatCount="indefinite"/><q:tgtEl><q:spTgt spid="2"/></q:tgtEl><x:media-extra/></q:cMediaNode></q:audio><q:cmd type="call" cmd="seek(0.5)"><q:cBhvr><q:cTn id="8" dur="1"/><q:tgtEl><q:spTgt spid="2"/></q:tgtEl></q:cBhvr><x:command-extra/></q:cmd><x:after/></q:tnLst></q:timing>"#;
    let timing = rpptx_oxml::timing::CT_Timing::from_xml(timing_xml).unwrap();
    let TimingNode::Audio(audio) = &timing.nodes()[1] else {
        panic!("expected typed audio node");
    };
    assert_eq!(audio.common.target, TimingTarget::Shape(2));
    assert_eq!(audio.common.volume, Some(80_000));
    assert!(audio.common.looped);
    assert_eq!(
        audio.common.display,
        Some(MediaDisplayPolicy::HideWhenStopped)
    );
    let TimingNode::MediaCommand(command) = &timing.nodes()[2] else {
        panic!("expected typed media command");
    };
    assert_eq!(command.command, MediaCommandKind::Seek { position_ms: 500 });
    assert_eq!(timing.to_xml(), timing_xml);
}

#[test]
fn smartart_parts_read_aliases_write_schema_order_and_preserve_raw_children() {
    use rpptx_oxml::diagram::{
        CT_DiagramColorsDefinition, CT_DiagramData, CT_DiagramDrawing, CT_DiagramLayoutDefinition,
        CT_DiagramStyleDefinition,
    };

    let data_xml = br#"<q:dataModel xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:f219" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" x:root="kept"><?root-before keep?><!--root-before-->root&amp;text<x:before/><x:ptLst><x:pt modelId="foreign"/></x:ptLst><q:ptLst x:list="kept"><?points keep?><!--points-note-->points&amp;text<x:list-child/><q:pt modelId="n1" x:modelId="foreign-shadow" x:type="pres" x:point="kept"><q:prSet/><x:before-text/><?point keep?><!--point-note-->point&amp;text<q:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Before</a:t></a:r></a:p></q:t><x:after-text/></q:pt><x:after-point/></q:ptLst><x:between/><x:cxnLst><x:cxn modelId="foreign"/></x:cxnLst><q:cxnLst x:list="kept"><?connections keep?><!--connections-note-->connections&amp;text<x:list-child/><q:cxn modelId="c1" x:modelId="foreign-shadow" srcId="n1" x:srcId="foreign" destId="n1" type="parOf" srcOrd="2" destOrd="3"/><x:after-connection/></q:cxnLst><x:bg><x:foreign/></x:bg><q:bg><x:background/></q:bg><q:whole><x:whole/></q:whole><q:extLst><a:ext><dsp:dataModelExt relId="drawing-scope-id" x:relId="foreign-drawing"/></a:ext></q:extLst><x:after/><?root-after keep?><!--root-after-->tail&amp;text</q:dataModel>"#;
    let mut data = CT_DiagramData::from_xml(data_xml).unwrap();
    assert_eq!(data.to_xml().unwrap(), data_xml);
    assert_eq!(data.points().len(), 1);
    assert_eq!(data.connections().len(), 1);
    assert_eq!(data.points()[0].model_id, "n1");
    assert!(data.background_xml().unwrap().starts_with(b"<q:bg>"));
    assert_eq!(data.drawing_relationship_id(), Some("drawing-scope-id"));
    data.set_node_text("n1", "After").unwrap();
    let written = data.to_xml().unwrap();
    let written_text = String::from_utf8(written.clone()).unwrap();
    assert!(written_text.starts_with("<dgm:dataModel"));
    assert!(
        written_text
            .contains("xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"")
    );
    assert!(written_text.contains("<dgm:t><a:bodyPr/>"));
    assert!(written_text.contains("<a:t>After</a:t>"));
    assert!(written_text.contains("<x:before-text/>"));
    assert!(written_text.contains("<x:after-text/>"));
    assert!(written_text.contains("x:list=\"kept\""));
    assert!(
        written_text.contains("<?points keep?><!--points-note-->points&amp;text<x:list-child/>")
    );
    assert!(written_text.contains("<?point keep?><!--point-note-->point&amp;text"));
    assert!(written_text.contains(
        "<?connections keep?><!--connections-note-->connections&amp;text<x:list-child/>"
    ));
    assert!(written_text.contains("<x:after-point/>"));
    assert!(written_text.contains("<x:after-connection/>"));
    assert!(written_text.contains("<?root-before keep?><!--root-before-->root&amp;text"));
    assert!(written_text.contains("<?root-after keep?><!--root-after-->tail&amp;text"));
    assert!(written_text.contains("<x:ptLst><x:pt modelId=\"foreign\"/></x:ptLst>"));
    assert!(written_text.contains("<x:cxnLst><x:cxn modelId=\"foreign\"/></x:cxnLst>"));
    assert!(written_text.contains("<q:bg><x:background/></q:bg>"));
    assert_order(
        &written,
        &[
            "<dgm:ptLst",
            "<dgm:cxnLst",
            "<q:bg",
            "<q:whole",
            "<q:extLst",
        ],
    );
    let reparsed = CT_DiagramData::from_xml(&written).unwrap();
    assert_eq!(
        reparsed.points()[0].text.as_ref().unwrap().plain_text(),
        "After"
    );

    let conflicting_prefix_xml = br#"<q:dataModel xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:dgm="urn:producer" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><q:ptLst><q:pt modelId="n1"><q:t><a:bodyPr/><a:lstStyle/><a:p/></q:t></q:pt><q:pt modelId="n2"><q:t xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><a:bodyPr/><a:lstStyle/><a:p/></q:t></q:pt></q:ptLst><q:cxnLst/><dgm:preserved/></q:dataModel>"#;
    let mut conflicting = CT_DiagramData::from_xml(conflicting_prefix_xml).unwrap();
    assert!(conflicting.points()[0].text.is_none());
    assert!(conflicting.set_node_text("n1", "Unsafe").is_err());
    conflicting.set_node_text("n2", "Unsafe").unwrap();
    let error = conflicting.to_xml().unwrap_err().to_string();
    assert!(error.contains("xmlns:dgm conflicts"));
    assert!(CT_DiagramData::from_xml(br#"<d:dataModel xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram"><d:ptLst/><d:bg/><d:cxnLst/></d:dataModel>"#).is_err());
    assert!(CT_DiagramData::from_xml(br#"<d:dataModel xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram"><d:ptLst/><d:extLst/><d:whole/></d:dataModel>"#).is_err());

    let layout_xml = br#"<x:layoutDef xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:f="urn:foreign" uniqueId="urn:list" f:uniqueId="urn:matrix"><f:title val="Foreign Matrix"/><x:title f:val="Foreign" val="Vertical List"/><x:catLst><f:cat type="matrix"/><x:cat f:type="cycle" type="list"/></x:catLst><x:layoutNode><f:alg type="cycle"/><x:alg f:type="cycle" type="lin"/><x:constrLst><f:constr type="foreign"/><x:constr f:type="foreign" type="w"/></x:constrLst><x:extLst><f:raw/></x:extLst></x:layoutNode></x:layoutDef>"#;
    let style_xml = br#"<x:styleDef xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:f="urn:foreign" uniqueId="urn:style" f:uniqueId="foreign"><f:styleLbl name="foreign"><a:lnRef idx="99"/></f:styleLbl><x:styleLbl name="node0" f:name="foreign"><x:style><f:lnRef idx="99"/><a:lnRef idx="2" f:idx="99"/><a:fillRef idx="1"/><a:effectRef idx="0"/><a:fontRef idx="minor"/></x:style><f:raw/></x:styleLbl></x:styleDef>"#;
    let colors_xml = br#"<x:colorsDef xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:f="urn:foreign" uniqueId="urn:colors" f:uniqueId="foreign"><f:styleLbl name="foreign"><f:fillClrLst><f:srgbClr val="FFFFFF"/></f:fillClrLst></f:styleLbl><x:styleLbl name="node0" f:name="foreign"><f:fillClrLst><a:srgbClr val="FFFFFF"/></f:fillClrLst><x:fillClrLst><f:srgbClr val="FFFFFF"/><a:srgbClr f:val="FFFFFF" val="112233"/></x:fillClrLst></x:styleLbl></x:colorsDef>"#;
    let drawing_xml = br#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:f="urn:foreign"><dsp:spTree><f:sp/><dsp:sp/><f:raw/></dsp:spTree></dsp:drawing>"#;
    let layout = CT_DiagramLayoutDefinition::from_xml(layout_xml).unwrap();
    let style = CT_DiagramStyleDefinition::from_xml(style_xml).unwrap();
    let colors = CT_DiagramColorsDefinition::from_xml(colors_xml).unwrap();
    let drawing = CT_DiagramDrawing::from_xml(drawing_xml).unwrap();
    assert_eq!(layout.to_xml(), layout_xml);
    assert_eq!(style.to_xml(), style_xml);
    assert_eq!(colors.to_xml(), colors_xml);
    assert_eq!(drawing.to_xml(), drawing_xml);
    assert_eq!(layout.categories, ["list"]);
    let shape_style = style.labels[0].shape_style.as_ref().unwrap();
    assert_eq!(shape_style.line_reference, Some(2));
    assert_eq!(shape_style.fill_reference, Some(1));
    assert_eq!(shape_style.font_reference.as_deref(), Some("minor"));
    assert_eq!(colors.labels[0].fill_colors, ["112233"]);
    assert_eq!(style.labels.len(), 1);
    assert_eq!(colors.labels.len(), 1);
    assert_eq!(drawing.shape_count, 1);
    assert!(
        CT_DiagramDrawing::from_xml(br#"<f:drawing xmlns:f="urn:foreign"><f:sp/></f:drawing>"#)
            .is_err()
    );
    assert!(CT_DiagramData::from_xml(br#"<d:dataModel xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram"><d:ptLst/></d:dataModel><d:dataModel/>"#).is_err());
    assert!(CT_DiagramLayoutDefinition::from_xml(br#"<d:layoutDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram"><d:layoutNode/></d:layoutDef>trailing"#).is_err());
    assert!(CT_DiagramStyleDefinition::from_xml(br#"<d:styleDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram"><d:styleLbl>"#).is_err());
    assert!(CT_DiagramColorsDefinition::from_xml(br#"<d:colorsDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram"><d:styleLbl/></d:colorsDef><d:colorsDef/>"#).is_err());
    assert!(CT_DiagramDrawing::from_xml(br#"<d:drawing xmlns:d="http://schemas.microsoft.com/office/drawing/2008/diagram"/>after"#).is_err());
}

#[test]
fn smartart_checked_edits_reject_duplicate_graph_identities_without_mutation() {
    use rpptx_oxml::diagram::CT_DiagramData;

    let duplicate_points = br#"<d:dataModel xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><d:ptLst><d:pt modelId="same"><d:t><a:bodyPr/><a:lstStyle/><a:p/></d:t></d:pt><d:pt modelId="same"><d:t><a:bodyPr/><a:lstStyle/><a:p/></d:t></d:pt></d:ptLst><d:cxnLst/></d:dataModel>"#;
    let mut data = CT_DiagramData::from_xml(duplicate_points).unwrap();
    let error = data
        .set_node_text("same", "ambiguous")
        .unwrap_err()
        .to_string();
    assert!(
        error.contains("duplicate diagram point modelId same"),
        "{error}"
    );
    assert_eq!(data.to_xml().unwrap(), duplicate_points);

    let duplicate_connections = br#"<d:dataModel xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><d:ptLst><d:pt modelId="one"><d:t><a:bodyPr/><a:lstStyle/><a:p/></d:t></d:pt></d:ptLst><d:cxnLst><d:cxn modelId="same" srcId="one" destId="one"/><d:cxn modelId="same" srcId="one" destId="one"/></d:cxnLst></d:dataModel>"#;
    let mut data = CT_DiagramData::from_xml(duplicate_connections).unwrap();
    let error = data
        .set_node_text("one", "ambiguous")
        .unwrap_err()
        .to_string();
    assert!(
        error.contains("duplicate diagram connection modelId same"),
        "{error}"
    );
    assert_eq!(data.to_xml().unwrap(), duplicate_connections);
}

#[test]
fn smartart_projections_ignore_same_namespace_lookalikes_outside_schema_positions() {
    use std::collections::HashMap;

    use rpptx_oxml::diagram::{
        CT_DiagramColorsDefinition, CT_DiagramData, CT_DiagramDrawing, CT_DiagramLayoutDefinition,
        CT_DiagramStyleDefinition, DiagramLayoutFamily,
    };

    let data_xml = br#"<d:dataModel xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><d:ptLst/><d:extLst><a:ext><d:opaque><dsp:dataModelExt relId="lookalike"/></d:opaque><dsp:dataModelExt relId="owned"/></a:ext></d:extLst></d:dataModel>"#;
    let mut data = CT_DiagramData::from_xml(data_xml).unwrap();
    assert_eq!(data.drawing_relationship_id(), Some("owned"));
    data.remap_drawing_relationship(&HashMap::from([
        ("owned".to_owned(), "copied".to_owned()),
        ("lookalike".to_owned(), "must-not-change".to_owned()),
    ]))
    .unwrap();
    let remapped = String::from_utf8(data.to_xml().unwrap()).unwrap();
    assert!(remapped.contains("relId=\"lookalike\""));
    assert!(remapped.contains("relId=\"copied\""));
    assert!(!remapped.contains("must-not-change"));

    let layout = CT_DiagramLayoutDefinition::from_xml(br#"<d:layoutDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram"><d:title val="Vertical List"/><d:catLst><d:cat type="list"/></d:catLst><d:layoutNode><d:alg type="lin"/><d:constrLst><d:constr type="w"/></d:constrLst><d:extLst><d:alg type="pyra"/><d:constr type="fake"/></d:extLst></d:layoutNode><d:extLst><d:title val="Pyramid"/><d:cat type="pyramid"/><d:layoutNode><d:alg type="pyra"/></d:layoutNode></d:extLst></d:layoutDef>"#).unwrap();
    assert_eq!(layout.title.as_deref(), Some("Vertical List"));
    assert_eq!(layout.categories, ["list"]);
    assert_eq!(layout.algorithms, ["lin"]);
    assert_eq!(layout.constraints, ["w"]);
    assert_eq!(
        layout.family,
        DiagramLayoutFamily::Unsupported("unknown".to_owned())
    );

    let style = CT_DiagramStyleDefinition::from_xml(br#"<d:styleDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><d:styleLbl name="owned"><d:style><a:lnRef idx="2"/><d:extLst><a:fillRef idx="99"/></d:extLst></d:style></d:styleLbl><d:extLst><d:styleLbl name="lookalike"><d:style><a:lnRef idx="99"/></d:style></d:styleLbl></d:extLst></d:styleDef>"#).unwrap();
    assert_eq!(style.labels.len(), 1);
    assert_eq!(style.labels[0].name, "owned");
    let shape_style = style.labels[0].shape_style.as_ref().unwrap();
    assert_eq!(shape_style.line_reference, Some(2));
    assert_eq!(shape_style.fill_reference, None);

    let colors = CT_DiagramColorsDefinition::from_xml(br#"<d:colorsDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><d:styleLbl name="owned"><d:fillClrLst><a:srgbClr val="112233"/><d:opaque><a:srgbClr val="FFFFFF"/></d:opaque></d:fillClrLst></d:styleLbl><d:extLst><d:styleLbl name="lookalike"><d:fillClrLst><a:srgbClr val="FFFFFF"/></d:fillClrLst></d:styleLbl></d:extLst></d:colorsDef>"#).unwrap();
    assert_eq!(colors.labels.len(), 1);
    assert_eq!(colors.labels[0].name, "owned");
    assert_eq!(colors.labels[0].fill_colors, ["112233"]);

    let drawing_xml = br#"<x:drawing xmlns:x="http://schemas.microsoft.com/office/drawing/2008/diagram"><x:sp/><x:opaque><x:sp/></x:opaque><x:spTree><x:sp/><x:opaque><x:sp/></x:opaque><x:sp/></x:spTree><x:extLst><x:sp/><x:spTree><x:sp/></x:spTree></x:extLst></x:drawing>"#;
    let drawing = CT_DiagramDrawing::from_xml(drawing_xml).unwrap();
    assert_eq!(drawing.shape_count, 2);
    assert_eq!(drawing.to_xml(), drawing_xml);
}

#[test]
fn smartart_render_projection_reads_only_schema_owned_values_and_preserves_raw_xml() {
    use oxml_drawing::color::{ColorChoice, ColorTransform};
    use rpptx_oxml::diagram::{CT_DiagramColorsDefinition, CT_DiagramLayoutDefinition};

    let xml = br#"<d:layoutDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:x="urn:producer" x:keep='A&#x20;B' uniqueId="urn:f220:cycle"><d:layoutNode><d:alg type="cycle" rev="0"><d:param type="stAng" val="90"/><d:param type="spanAng" val="360"/></d:alg><d:constrLst><d:constr type="w" for="ch" val="0.5"/><d:constr type="h" for="ch" fact="0.8"/></d:constrLst><d:ruleLst><d:rule type="maxVal" for="ch" val="6"/></d:ruleLst><d:extLst><d:alg type="pyramid"><d:param type="stAng" val="must-not-project"/></d:alg><d:constrLst><d:constr type="fake" val="99"/></d:constrLst><d:ruleLst><d:rule type="fake" val="99"/></d:ruleLst></d:extLst><x:alg type="foreign"/></d:layoutNode><d:extLst xmlns:d="urn:shadow"><d:layoutNode><d:alg type="shadow"/></d:layoutNode></d:extLst></d:layoutDef>"#;
    let layout = CT_DiagramLayoutDefinition::from_xml(xml).unwrap();

    assert_eq!(
        layout.render_projection().algorithms,
        [(
            "cycle".to_owned(),
            vec![
                ("rev".to_owned(), "0".to_owned()),
                ("stAng".to_owned(), "90".to_owned()),
                ("spanAng".to_owned(), "360".to_owned()),
            ],
        )]
    );
    assert_eq!(
        layout.render_projection().constraints,
        [
            vec![
                ("type".to_owned(), "w".to_owned()),
                ("for".to_owned(), "ch".to_owned()),
                ("val".to_owned(), "0.5".to_owned()),
            ],
            vec![
                ("type".to_owned(), "h".to_owned()),
                ("for".to_owned(), "ch".to_owned()),
                ("fact".to_owned(), "0.8".to_owned()),
            ],
        ]
    );
    assert_eq!(
        layout.render_projection().rules,
        [vec![
            ("type".to_owned(), "maxVal".to_owned()),
            ("for".to_owned(), "ch".to_owned()),
            ("val".to_owned(), "6".to_owned()),
        ]]
    );
    assert_eq!(layout.to_xml(), xml);

    let colors_xml = br#"<d:colorsDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a2="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><d:styleLbl name="node"><d:fillClrLst><a2:schemeClr val="accent1"><a2:tint val="25000"/><x:shade val="must-not-project"/><a2:lumMod val="80000"/><a:tint xmlns:a="urn:shadow" val="must-not-project"/></a2:schemeClr><x:srgbClr val="FFFFFF"/></d:fillClrLst><d:linClrLst><a2:srgbClr val="112233"><a2:alpha val="50000"/></a2:srgbClr></d:linClrLst><d:extLst><d:fillClrLst><a2:srgbClr val="AABBCC"/></d:fillClrLst></d:extLst></d:styleLbl></d:colorsDef>"#;
    let colors = CT_DiagramColorsDefinition::from_xml(colors_xml).unwrap();
    let projected = &colors.render_projection().labels[0];
    assert_eq!(projected.name, "node");
    assert_eq!(projected.fill.len(), 1);
    let ColorChoice::Scheme {
        value, transforms, ..
    } = &projected.fill[0]
    else {
        panic!("expected aliased scheme colour")
    };
    assert_eq!(value, "accent1");
    assert!(matches!(
        transforms.as_slice(),
        [
            ColorTransform::Tint(_),
            ColorTransform::LuminanceModulation(_)
        ]
    ));
    assert_eq!(projected.line.len(), 1);
    assert!(
        matches!(&projected.line[0], ColorChoice::Srgb { transforms, .. } if matches!(transforms.as_slice(), [ColorTransform::Alpha(_)]))
    );
    assert!(projected.effect.is_empty());
    assert!(projected.text_fill.is_empty());
    assert_eq!(colors.to_xml(), colors_xml);
}

#[test]
fn smartart_layout_family_requires_one_of_the_six_exact_authentic_identities() {
    use rpptx_oxml::diagram::{CT_DiagramLayoutDefinition, DiagramLayoutFamily};

    for (identity, expected) in [
        (
            "urn:microsoft.com/office/officeart/2005/8/layout/list1",
            DiagramLayoutFamily::List,
        ),
        (
            "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1",
            DiagramLayoutFamily::Hierarchy,
        ),
        (
            "urn:microsoft.com/office/officeart/2005/8/layout/cycle1",
            DiagramLayoutFamily::Cycle,
        ),
        (
            "urn:microsoft.com/office/officeart/2009/3/layout/CircleRelationship",
            DiagramLayoutFamily::Relationship,
        ),
        (
            "urn:microsoft.com/office/officeart/2005/8/layout/matrix1",
            DiagramLayoutFamily::Matrix,
        ),
        (
            "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1",
            DiagramLayoutFamily::Pyramid,
        ),
        (
            "urn:f220:list",
            DiagramLayoutFamily::Unsupported("urn:f220:list".to_owned()),
        ),
        (
            "urn:producer:unrelated-list-layout",
            DiagramLayoutFamily::Unsupported("urn:producer:unrelated-list-layout".to_owned()),
        ),
    ] {
        let xml = format!(
            r#"<d:layoutDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="{identity}"><d:title val="Authentic identity"/><d:catLst><d:cat type="list"/></d:catLst><d:layoutNode><d:forEach axis="ch"><d:layoutNode><d:alg type="lin"/></d:layoutNode></d:forEach></d:layoutNode></d:layoutDef>"#
        );
        let layout = CT_DiagramLayoutDefinition::from_xml(xml.as_bytes()).unwrap();
        assert_eq!(layout.family, expected, "{identity}");
        assert_eq!(layout.to_xml(), xml.as_bytes());
    }
}

#[test]
fn smartart_render_projection_keeps_nested_instruction_ownership_and_excludes_lookalikes() {
    use rpptx_oxml::diagram::{CT_DiagramLayoutDefinition, DiagramRenderInstructionKind};

    let xml = br#"<q:layoutDef xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:dgm="urn:shadow" xmlns:x="urn:extension" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/list1"><q:layoutNode styleLbl="root"><q:forEach axis="ch" ptType="node"><q:layoutNode styleLbl="child"><q:alg type="lin"><q:param type="linDir" val="fromL"/></q:alg><q:shape type="rect"/><q:presOf axis="self" ptType="node"/></q:layoutNode></q:forEach><q:choose name="owned"><q:if func="cnt" op="gte" val="2"><q:layoutNode name="many"><q:shape type="ellipse"/></q:layoutNode></q:if><q:else><q:layoutNode name="one"><q:shape type="rect"/></q:layoutNode></q:else></q:choose><x:forEach axis="must-not-project"><x:alg type="must-not-project"/></x:forEach><q:extLst><q:layoutNode styleLbl="must-not-project"><q:alg type="must-not-project"/></q:layoutNode></q:extLst><dgm:layoutNode styleLbl="shadow"><dgm:alg type="shadow"/></dgm:layoutNode></q:layoutNode></q:layoutDef>"#;
    let layout = CT_DiagramLayoutDefinition::from_xml(xml).unwrap();
    let projection = layout.render_projection();
    let root = projection.root.as_ref().expect("schema-owned layout root");
    assert_eq!(root.kind, DiagramRenderInstructionKind::LayoutNode);
    assert_eq!(
        root.attributes,
        [("styleLbl".to_owned(), "root".to_owned())]
    );
    assert_eq!(
        root.children
            .iter()
            .map(|child| &child.kind)
            .collect::<Vec<_>>(),
        [
            &DiagramRenderInstructionKind::ForEach,
            &DiagramRenderInstructionKind::Choose,
        ]
    );
    let loop_node = &root.children[0].children[0];
    assert_eq!(loop_node.kind, DiagramRenderInstructionKind::LayoutNode);
    assert_eq!(
        loop_node
            .children
            .iter()
            .map(|child| &child.kind)
            .collect::<Vec<_>>(),
        [
            &DiagramRenderInstructionKind::Algorithm,
            &DiagramRenderInstructionKind::Shape,
            &DiagramRenderInstructionKind::PresentationOf,
        ]
    );
    assert_eq!(
        root.children[1].children[0].kind,
        DiagramRenderInstructionKind::Condition
    );
    assert_eq!(
        root.children[1].children[1].kind,
        DiagramRenderInstructionKind::Else
    );
    assert_eq!(layout.to_xml(), xml);
}

#[test]
fn smartart_render_projection_rejects_excessive_layout_depth() {
    use rpptx_oxml::diagram::CT_DiagramLayoutDefinition;

    let mut xml = String::from(
        r#"<d:layoutDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:f220:list">"#,
    );
    for _ in 0..66 {
        xml.push_str("<d:layoutNode>");
    }
    xml.push_str(r#"<d:alg type="lin"/>"#);
    for _ in 0..66 {
        xml.push_str("</d:layoutNode>");
    }
    xml.push_str("</d:layoutDef>");
    let error = CT_DiagramLayoutDefinition::from_xml(xml.as_bytes())
        .unwrap_err()
        .to_string();
    assert!(error.contains("depth bound 64"), "{error}");
}

#[test]
fn smartart_text_with_nested_fixed_namespace_shadow_stays_opaque_and_uneditable() {
    use rpptx_oxml::diagram::CT_DiagramData;

    let unsafe_text = r#"<dgm:t><a:bodyPr/><a:lstStyle/><a:p xmlns:a="urn:producer"><a:r><a:t>Foreign</a:t></a:r></a:p></dgm:t>"#;
    let xml = format!(
        r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="unsafe">{unsafe_text}</dgm:pt><dgm:pt modelId="safe"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Before</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/></dgm:dataModel>"#
    );
    let mut data = CT_DiagramData::from_xml(xml.as_bytes()).unwrap();
    assert!(data.points()[0].text.is_none());
    assert_eq!(data.to_xml().unwrap(), xml.as_bytes());
    assert!(data.set_node_text("unsafe", "Must reject").is_err());
    assert_eq!(data.to_xml().unwrap(), xml.as_bytes());

    data.set_node_text("safe", "After").unwrap();
    let written = String::from_utf8(data.to_xml().unwrap()).unwrap();
    assert!(written.contains(unsafe_text));
    assert!(written.contains("<a:t>After</a:t>"));
    let reopened = CT_DiagramData::from_xml(written.as_bytes()).unwrap();
    assert!(reopened.points()[0].text.is_none());
    assert_eq!(
        reopened.points()[1].text.as_ref().unwrap().plain_text(),
        "After"
    );
}

#[test]
fn smartart_text_with_fixed_dgm_root_shadow_stays_opaque_and_uneditable() {
    use rpptx_oxml::diagram::CT_DiagramData;

    let unsafe_text = r#"<q:t xmlns:q="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:dgm="urn:producer" xmlns:x="urn:safe" x:keep="producer"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Foreign</a:t></a:r></a:p></q:t>"#;
    let xml = format!(
        r#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="shadow">{unsafe_text}</dgm:pt><dgm:pt modelId="safe"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Before</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/></dgm:dataModel>"#
    );
    let mut data = CT_DiagramData::from_xml(xml.as_bytes()).unwrap();
    assert!(data.points()[0].text.is_none());
    assert_eq!(data.to_xml().unwrap(), xml.as_bytes());

    assert!(data.set_node_text("shadow", "Must reject").is_err());
    assert_eq!(data.to_xml().unwrap(), xml.as_bytes());

    data.set_node_text("safe", "After").unwrap();
    let written = String::from_utf8(data.to_xml().unwrap()).unwrap();
    assert!(written.contains(unsafe_text));
    assert!(written.contains("<a:t>After</a:t>"));
}

#[test]
fn edited_smartart_text_retains_safe_root_attributes_and_namespaces() {
    use rpptx_oxml::diagram::CT_DiagramData;

    let xml = br#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="safe"><dgm:t xmlns:x="urn:producer" xmlns:y="urn:unused" x:keep="producer"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Before</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/></dgm:dataModel>"#;
    let mut data = CT_DiagramData::from_xml(xml).unwrap();
    data.set_node_text("safe", "After").unwrap();
    let written = String::from_utf8(data.to_xml().unwrap()).unwrap();
    assert!(
        written.contains(
            "<dgm:t xmlns:x=\"urn:producer\" xmlns:y=\"urn:unused\" x:keep=\"producer\">"
        )
    );
    assert!(written.contains("<a:t>After</a:t>"));

    let reopened = CT_DiagramData::from_xml(written.as_bytes()).unwrap();
    assert_eq!(
        reopened.points()[0].text.as_ref().unwrap().plain_text(),
        "After"
    );
}

#[test]
fn supported_smartart_nodes_remain_editable_after_save_and_reopen() {
    use rpptx_oxml::diagram::{
        CT_DiagramData, CT_DiagramLayoutDefinition, DiagramConnectionKind, DiagramLayoutFamily,
        DiagramPointKind, DiagramRelationshipIds,
    };
    use std::collections::HashMap;

    let xml = br#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="n1"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>One</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="a1" type="asst"><dgm:t><a:bodyPr/><a:lstStyle/><a:p/></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="c1" srcId="n1" destId="a1" type="presOf" srcOrd="4" destOrd="5"/></dgm:cxnLst></dgm:dataModel>"#;
    let mut data = CT_DiagramData::from_xml(xml).unwrap();
    assert_eq!(data.points()[0].kind, DiagramPointKind::Node);
    assert_eq!(data.points()[1].kind, DiagramPointKind::Assistant);
    assert_eq!(
        data.connections()[0].kind,
        DiagramConnectionKind::PresentationOf
    );
    assert_eq!(
        (
            data.connections()[0].source_order,
            data.connections()[0].destination_order
        ),
        (4, 5)
    );
    data.set_node_text("n1", "Edited").unwrap();
    data.move_point(1, 0).unwrap();
    let reopened = CT_DiagramData::from_xml(&data.to_xml().unwrap()).unwrap();
    assert_eq!(reopened.points()[0].model_id, "a1");
    assert_eq!(
        reopened.points()[1].text.as_ref().unwrap().plain_text(),
        "Edited"
    );

    let layout = CT_DiagramLayoutDefinition::from_xml(br#"<d:layoutDef xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:hierarchy4"><d:layoutNode><d:alg type="hierRoot"/></d:layoutNode></d:layoutDef>"#).unwrap();
    assert_eq!(
        layout.family,
        DiagramLayoutFamily::Unsupported("urn:hierarchy4".to_owned())
    );

    let mut ids = DiagramRelationshipIds::from_xml(br#"<d:relIds xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:x="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="urn:producer" p:keep='A&#x20;B' x:dm="same" x:lo="same-layout" x:qs="same-style" x:cs="same-colors"><p:child x:id="unmapped"/></d:relIds>"#).unwrap();
    ids.drawing = Some("same-drawing".to_owned());
    ids.remap(&HashMap::from([
        ("same".to_owned(), "rId9".to_owned()),
        ("same-drawing".to_owned(), "rId10".to_owned()),
    ]))
    .unwrap();
    assert_eq!(ids.data, "rId9");
    assert_eq!(ids.drawing.as_deref(), Some("rId10"));
    let remapped = String::from_utf8(ids.to_xml()).unwrap();
    assert!(remapped.contains("x:dm=\"rId9\""));
    assert!(remapped.contains("p:keep='A&#x20;B'"));
    assert!(remapped.contains("<p:child x:id=\"unmapped\"/>"));
    assert!(!remapped.contains("r:draw"));

    let frame_xml = br#"<p:graphicFrame xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="2" name="SmartArt"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><d:relIds rel:dm="inherited-data" rel:lo="inherited-layout" rel:qs="inherited-style" rel:cs="inherited-colors"><d:ext rel:id="unmapped"/></d:relIds></a:graphicData></a:graphic></p:graphicFrame>"#;
    let frame = rpptx_oxml::graphic_frame::CT_GraphicFrame::from_xml(frame_xml).unwrap();
    let rpptx_oxml::graphic_frame::GraphicDataPayload::SmartArt(inherited) =
        frame.graphic_data.payload()
    else {
        panic!("expected inherited SmartArt relationship ids");
    };
    let mut inherited = (**inherited).clone();
    inherited
        .remap(&HashMap::from([
            ("inherited-data".to_owned(), "copied-data".to_owned()),
            ("inherited-layout".to_owned(), "copied-layout".to_owned()),
        ]))
        .unwrap();
    let remapped = String::from_utf8(inherited.to_xml()).unwrap();
    assert!(remapped.contains("rel:dm=\"copied-data\""));
    assert!(remapped.contains("rel:lo=\"copied-layout\""));
    assert!(remapped.contains("rel:qs=\"inherited-style\""));
    assert!(remapped.contains("<d:ext rel:id=\"unmapped\"/>"));

    assert!(DiagramRelationshipIds::from_xml(br#"<d:relIds xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="a" r:lo="b" r:qs="c" r:cs="d">"#).is_err());
}

#[test]
fn every_corpus_smartart_part_projects_without_rewriting_source_bytes() {
    use rpptx_oxml::diagram::{
        CT_DiagramColorsDefinition, CT_DiagramData, CT_DiagramDrawing, CT_DiagramLayoutDefinition,
        CT_DiagramStyleDefinition,
    };

    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut coverage = [0usize; 5];
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            let Some(xml) = package.get_part(part) else {
                continue;
            };
            if content_type.contains("diagramData") {
                let parsed = CT_DiagramData::from_xml(xml)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                assert_eq!(parsed.to_xml().unwrap(), xml, "{} {part}", entry.path);
                coverage[0] += 1;
            } else if content_type.contains("diagramLayout") {
                let parsed = CT_DiagramLayoutDefinition::from_xml(xml)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                assert_eq!(parsed.to_xml(), xml, "{} {part}", entry.path);
                coverage[1] += 1;
            } else if content_type.contains("diagramStyle") {
                let parsed = CT_DiagramStyleDefinition::from_xml(xml)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                assert_eq!(parsed.to_xml(), xml, "{} {part}", entry.path);
                coverage[2] += 1;
            } else if content_type.contains("diagramColors") {
                let parsed = CT_DiagramColorsDefinition::from_xml(xml)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                assert_eq!(parsed.to_xml(), xml, "{} {part}", entry.path);
                coverage[3] += 1;
            } else if content_type.contains("diagramDrawing") {
                let parsed = CT_DiagramDrawing::from_xml(xml)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                assert_eq!(parsed.to_xml(), xml, "{} {part}", entry.path);
                coverage[4] += 1;
            }
        }
    }
    assert!(coverage.iter().all(|count| *count > 0), "{coverage:?}");
}

#[test]
fn media_trigger_uses_ancestor_click_sequence_context() {
    use rpptx_oxml::timing::{CT_Timing, MediaPlaybackTrigger};

    let xml = br#"<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst><p:par><p:cTn id="1" nodeType="clickEffect"><p:childTnLst><p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="2" dur="1"/><p:tgtEl><p:spTgt spid="9"/></p:tgtEl></p:cBhvr></p:cmd></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"#;
    let timing = CT_Timing::from_xml(xml).unwrap();
    assert_eq!(
        timing.media_playback_trigger_for_shape(9),
        Some(MediaPlaybackTrigger::InClickSequence)
    );
}

#[test]
fn media_insertion_uses_direct_presentation_list_and_precedes_later_children() {
    use rpptx_oxml::timing::{CT_Timing, MediaPlaybackTrigger};

    let without_list = br#"<q:timing xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:f215"><x:keep a='A&#x20;B'/><q:bldLst/><q:extLst><x:tail/></q:extLst></q:timing>"#;
    let mut timing = CT_Timing::from_xml(without_list).unwrap();
    timing
        .add_media(
            false,
            9,
            80_000,
            false,
            false,
            MediaPlaybackTrigger::Automatic,
        )
        .unwrap();
    let written = String::from_utf8(timing.to_xml()).unwrap();
    assert!(written.contains("<x:keep a='A&#x20;B'/><p:tnLst"));
    assert!(written.find("<p:tnLst").unwrap() < written.find("<q:bldLst").unwrap());
    assert!(written.find("<q:bldLst").unwrap() < written.find("<q:extLst").unwrap());
    CT_Timing::from_xml(written.as_bytes()).unwrap();

    let foreign_first = br#"<q:timing xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:f215"><x:tnLst><x:keep a='A&#x20;B'/></x:tnLst><q:tnLst><x:existing/></q:tnLst></q:timing>"#;
    let mut timing = CT_Timing::from_xml(foreign_first).unwrap();
    timing
        .add_media(
            true,
            10,
            90_000,
            false,
            false,
            MediaPlaybackTrigger::OnClick,
        )
        .unwrap();
    let written = String::from_utf8(timing.to_xml()).unwrap();
    assert!(written.contains("<x:tnLst><x:keep a='A&#x20;B'/></x:tnLst>"));
    assert!(written.contains("<q:tnLst><x:existing/><p:video"));
    CT_Timing::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn media_removal_preserves_media_shaped_content_in_unmodelled_wrappers() {
    use rpptx_oxml::timing::{CT_Timing, TimingNode};

    let retained = r#"<x:payload a='A&#x20;B'><p:cmd type="call" cmd="stop"><p:cBhvr><p:cTn id="90"/><p:tgtEl><p:spTgt spid="9"/></p:tgtEl></p:cBhvr></p:cmd></x:payload>"#;
    let owned_audio = r#"<p:audio><p:cMediaNode><p:cTn id="1"/><p:tgtEl><p:spTgt spid="9"/></p:tgtEl></p:cMediaNode></p:audio>"#;
    let owned_command = r#"<p:cmd type="call" cmd="play"><p:cBhvr><p:cTn id="2"/><p:tgtEl><p:spTgt spid="9"/></p:tgtEl></p:cBhvr></p:cmd>"#;
    let other_video = r#"<p:video><p:cMediaNode><p:cTn id="3"/><p:tgtEl><p:spTgt spid="10"/></p:tgtEl></p:cMediaNode></p:video>"#;
    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst>{retained}{owned_audio}{owned_command}{other_video}</p:tnLst></p:timing>"#
    );
    let mut timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    timing.remove_media(9).unwrap();
    let written = String::from_utf8(timing.to_xml()).unwrap();
    assert!(written.contains(retained));
    assert!(!written.contains(owned_audio));
    assert!(!written.contains(owned_command));
    assert!(written.contains(other_video));
    assert_eq!(written.matches("<p:cmd").count(), 1);
    assert_eq!(timing.nodes().len(), 2);
    assert!(matches!(timing.nodes()[0], TimingNode::Unsupported(_)));
    assert!(matches!(timing.nodes()[1], TimingNode::Video(_)));
}

#[test]
fn media_id_allocation_ignores_foreign_nonnumeric_time_nodes() {
    use rpptx_oxml::timing::{CT_Timing, MediaPlaybackTrigger};

    let retained = r#"<x:payload><x:cTn id="producer"/></x:payload>"#;
    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst>{retained}<p:par><p:cTn id="4"/></p:par></p:tnLst></p:timing>"#
    );
    let mut timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    timing
        .add_media(
            false,
            9,
            80_000,
            false,
            false,
            MediaPlaybackTrigger::Automatic,
        )
        .unwrap();
    let written = String::from_utf8(timing.to_xml()).unwrap();
    assert!(written.contains(retained));
    assert!(
        written.contains(r#"<p:cTn id="5" dur="indefinite"/>"#),
        "{written}"
    );
    assert!(written.contains(r#"<p:cTn id="6" dur="1" fill="hold""#));
}

#[test]
fn media_id_allocation_counts_unsupported_nodes_and_ignores_raw_maximum_ids() {
    use rpptx_oxml::timing::{CT_Timing, MediaPlaybackTrigger};

    let retained = r#"<x:payload><p:cTn id="4294967295"/></x:payload>"#;
    let unsupported = r#"<p:animClr><p:cBhvr><p:cTn id="11"/><p:tgtEl><p:spTgt spid="3"/></p:tgtEl></p:cBhvr></p:animClr>"#;
    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst>{retained}<p:par><p:cTn id="7"/></p:par>{unsupported}</p:tnLst></p:timing>"#
    );
    let mut timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    timing
        .add_media(
            true,
            10,
            90_000,
            false,
            false,
            MediaPlaybackTrigger::OnClick,
        )
        .unwrap();
    let written = String::from_utf8(timing.to_xml()).unwrap();
    assert!(written.contains(retained));
    assert!(written.contains(unsupported));
    assert!(written.contains(r#"<p:cTn id="12" dur="indefinite"/>"#));
    assert!(written.contains(r#"<p:cTn id="13" dur="1" fill="hold" nodeType="clickEffect""#));
}

#[test]
fn nonnumeric_media_command_offsets_remain_explicit_and_byte_preserved() {
    use rpptx_oxml::timing::{CT_Timing, MediaCommandKind, TimingNode};

    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst><p:cmd type="call" cmd="seek(bookmark)" x:keep='A&#x20;B'><p:cBhvr><p:cTn id="4"/><p:tgtEl><p:spTgt spid="9"/></p:tgtEl></p:cBhvr><x:tail/></p:cmd></p:tnLst></p:timing>"#
    );
    let timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    let TimingNode::MediaCommand(command) = &timing.nodes()[0] else {
        panic!("expected typed media command");
    };
    assert_eq!(
        command.command,
        MediaCommandKind::Other("seek(bookmark)".to_owned())
    );
    assert_eq!(timing.to_xml(), xml.as_bytes());
}

use oxml_core::raw_xml::{capture_element, capture_empty_element};
use oxml_core::units::Emu;
use oxml_drawing::color::{ColorMapSlot, ThemeColorSlot};
use oxml_drawing::shape_props::CT_ShapeProperties;
use oxml_drawing::table::CT_Table;
use oxml_drawing::text::CT_TextBody;
use oxml_drawing::xfrm::CT_Transform2D;
use oxml_opc::relationship::rel_types;
use oxml_opc::{OpcPackage, content_types};
use quick_xml::events::{BytesEnd, BytesStart, Event};
use quick_xml::{Reader, Writer, XmlVersion};
use rpptx_oxml::PRESENTATION_PART;
use rpptx_oxml::connector::CT_ConnectionShape;
use rpptx_oxml::graphic_frame::{CT_GraphicFrame, GraphicDataPayload};
use rpptx_oxml::namespace::{MC_NS, P_NS, P_PREFIX, R_NS, R_PREFIX};
use rpptx_oxml::notes_parts::{CT_NotesMaster, CT_NotesSlide};
use rpptx_oxml::picture::CT_Picture;
use rpptx_oxml::placeholder::{CT_Placeholder, PhType, PlaceholderKey};
use rpptx_oxml::presentation::{CT_Presentation, CT_SlideId};
use rpptx_oxml::relmap::{rewrite_exact_rel_ids, rewrite_rel_ids};
use rpptx_oxml::shape_tree::{
    CT_Shape, CT_ShapeTree, ShapeIdAllocator, ShapeTreeChild, rewrite_shape_ids,
};
use rpptx_oxml::slide_parts::{
    BackgroundRendering, CT_Slide, CT_SlideLayout, CT_SlideMaster, ColorMapOverrideKind,
};
use rpptx_oxml::timing::{
    CT_SlideTransition, CT_Timing, TimingDuration, TimingEvent, TimingNode, TimingNodeType,
    TimingTarget, TransitionEffect,
};

const MANIFEST: &str = include_str!("../../../scripts/pptx-corpus-manifest.tsv");
const EXPECTED_DECKS: usize = 50;
type DrawingElements = (Vec<Vec<u8>>, Vec<Vec<u8>>);

#[test]
fn timing_sequences_parallel_groups_triggers_and_effects_round_trip_in_schema_order() {
    let xml = format!(
        r#"<q:timing xmlns:q="{P_NS}" xmlns:r="{R_NS}" xmlns:x="urn:f213"><q:tnLst><q:par><q:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><q:stCondLst><q:cond evt="onClick" delay="0"><q:tgtEl><q:sldTgt/></q:tgtEl></q:cond></q:stCondLst><q:childTnLst><q:seq concurrent="1" nextAc="seek"><q:cTn id="2" grpId="2" dur="1000" nodeType="mainSeq"><q:childTnLst><q:animEffect transition="in" filter="fade"><q:cBhvr><q:cTn id="3" dur="250" presetID="1" presetClass="entr" presetSubtype="0"/><q:tgtEl><q:spTgt spid="7"/></q:tgtEl></q:cBhvr></q:animEffect><q:animMotion path="M 0 0 L 1 1 E" origin="layout"><q:cBhvr><q:cTn id="4" dur="750"/><q:tgtEl><q:spTgt spid="8"/></q:tgtEl></q:cBhvr></q:animMotion><q:set><q:cBhvr><q:cTn id="5" dur="1"/><q:tgtEl><q:spTgt spid="7"/></q:tgtEl><q:attrNameLst><q:attrName>style.visibility</q:attrName></q:attrNameLst></q:cBhvr><q:to><q:strVal val="visible"/></q:to></q:set><x:command r:id="rId9"/></q:childTnLst></q:cTn><q:prevCondLst><q:cond evt="onPrev" delay="0"><q:tn val="3"/></q:cond></q:prevCondLst><q:nextCondLst><q:cond evt="onNext" delay="0"/></q:nextCondLst></q:seq></q:childTnLst></q:cTn></q:par></q:tnLst><q:bldLst><q:bldP spid="7" grpId="2" build="p" bldLvl="2" rev="1" advAuto="500" animBg="0" autoUpdateAnimBg="1" uiExpand="1"/></q:bldLst></q:timing>"#
    );
    let timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(timing.builds().len(), 1);
    let build = &timing.builds()[0];
    assert_eq!(build.group_id, Some(2));
    assert_eq!(build.build_mode.as_deref(), Some("p"));
    assert_eq!(build.build_level, Some(2));
    assert_eq!(build.reverse, Some(true));
    assert_eq!(
        build.advance_automatically,
        Some(TimingDuration::Finite(500))
    );
    let TimingNode::Parallel(root) = &timing.nodes()[0] else {
        panic!("expected timing root parallel node");
    };
    assert_eq!(root.common.node_type, Some(TimingNodeType::TimingRoot));
    assert_eq!(
        root.common.start_conditions[0].event,
        Some(TimingEvent::OnClick)
    );
    assert_eq!(root.common.start_conditions[0].target, TimingTarget::Slide);
    let TimingNode::Sequence(sequence) = &root.common.children[0] else {
        panic!("expected main sequence");
    };
    assert_eq!(sequence.common.duration, TimingDuration::Finite(1000));
    assert_eq!(sequence.common.group_id, build.group_id);
    assert_eq!(
        sequence.previous_conditions[0].event,
        Some(TimingEvent::OnPrevious)
    );
    assert_eq!(
        sequence.previous_conditions[0].target,
        TimingTarget::TimeNode(3)
    );
    assert_eq!(sequence.next_conditions[0].event, Some(TimingEvent::OnNext));
    assert_eq!(
        sequence.next_conditions[0].target,
        TimingTarget::Unsupported
    );
    let TimingNode::Effect(effect) = &sequence.common.children[0] else {
        panic!("expected entrance effect");
    };
    assert_eq!(effect.target, TimingTarget::Shape(7));
    let TimingNode::Motion(motion) = &sequence.common.children[1] else {
        panic!("expected motion path");
    };
    assert_eq!(motion.target, TimingTarget::Shape(8));
    let TimingNode::Set(set) = &sequence.common.children[2] else {
        panic!("expected set behavior");
    };
    assert_eq!(set.value.as_deref(), Some("visible"));
    assert!(matches!(
        sequence.common.children[3],
        TimingNode::Unsupported(_)
    ));
    assert_eq!(timing.to_xml(), xml.as_bytes());
    assert_eq!(timing, CT_Timing::from_xml(&timing.to_xml()).unwrap());
}

#[test]
fn timing_mutation_preserves_every_unsupported_sibling_and_relationship_attribute() {
    let xml = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R_NS}" xmlns:x="urn:f213"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><x:before/><p:timing xmlns:a="urn:shadow-a" xmlns:r="urn:shadow-r"><p:tnLst><x:before-node r:id="rId4"><x:data/></x:before-node><p:par><p:cTn x:keep='A&#x20;B' r:id='rId&#x37;' id='1' dur='100'><p:childTnLst><x:media r:id="rId7"/></p:childTnLst></p:cTn></p:par><x:after-node/></p:tnLst><x:after-list/></p:timing><x:after/></p:sld>"#
    );
    let mut slide = CT_Slide::from_xml(xml.as_bytes()).unwrap();
    slide
        .timing
        .as_mut()
        .unwrap()
        .set_node_duration(1, TimingDuration::Finite(275))
        .unwrap();
    let written = String::from_utf8(slide.to_xml().unwrap()).unwrap();
    assert!(written.contains("dur='275'"));
    assert!(written.contains("<p:cTn x:keep='A&#x20;B' r:id='rId&#x37;' id='1' dur='275'>"));
    assert!(written.contains(r#"<x:before-node r:id="rId4"><x:data/></x:before-node>"#));
    assert!(written.contains(r#"<x:media r:id="rId7"/>"#));
    assert!(written.contains("<x:after-node/>"));
    assert!(written.contains("<x:after-list/>"));
    assert_order(
        written.as_bytes(),
        &["<x:before/>", "<p:timing", "<x:after/>"],
    );
    let reopened = CT_Slide::from_xml(written.as_bytes()).unwrap();
    let reopened_timing = reopened.timing.unwrap();
    let root = reopened_timing
        .nodes()
        .iter()
        .find_map(|node| match node {
            TimingNode::Parallel(root) => Some(root),
            _ => None,
        })
        .expect("expected parallel node");
    assert_eq!(root.common.duration, TimingDuration::Finite(275));
}

#[test]
fn timing_duration_mutation_never_searches_unsupported_node_subtrees() {
    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst><x:node><p:cTn id="2" dur="900" x:keep="yes"/></x:node><p:par><p:cTn id="1" dur="100"><p:childTnLst><p:set><p:cBhvr><p:cTn id="2" dur="125"/></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"#
    );
    let mut timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    timing
        .set_node_duration(2, TimingDuration::Finite(275))
        .unwrap();
    let written = String::from_utf8(timing.to_xml()).unwrap();
    assert!(written.contains(r#"<x:node><p:cTn id="2" dur="900" x:keep="yes"/></x:node>"#));
    assert!(written.contains(r#"<p:cBhvr><p:cTn id="2" dur="275"/></p:cBhvr>"#));

    let unsupported_only = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst><x:node><p:cTn id="9" dur="900"/></x:node></p:tnLst></p:timing>"#
    );
    let mut timing = CT_Timing::from_xml(unsupported_only.as_bytes()).unwrap();
    let before = timing.clone();
    assert!(
        timing
            .set_node_duration(9, TimingDuration::Finite(275))
            .is_err()
    );
    assert_eq!(timing, before);
}

#[test]
fn malformed_timing_does_not_partially_mutate_the_slide() {
    for timing in [
        r#"<p:timing><p:tnLst><p:par><p:cTn id="1" dur="broken"/></p:par></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:par><p:cTn dur="1"/></p:par></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst/><p:tnLst/></p:timing>"#,
        r#"<p:timing><p:bldLst/><p:tnLst/></p:timing>"#,
        r#"<p:timing><p:tnLst><p:par><p:cTn id="1"><p:childTnLst/><p:stCondLst/></p:cTn></p:par></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:par><p:cTn id="1"><p:stCondLst/><p:stCondLst/></p:cTn></p:par></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:set><p:cBhvr><p:tgtEl><p:spTgt spid="1"/></p:tgtEl><p:cTn id="2"/></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:set><p:cBhvr><p:cTn id="2"/><p:tgtEl><p:spTgt spid="1"/></p:tgtEl><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:set><p:to><p:strVal val="visible"/></p:to><p:cBhvr><p:cTn id="2"/></p:cBhvr></p:set></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:set><x:wrapper><p:cBhvr><p:cTn id="2"/></p:cBhvr></x:wrapper></p:set></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:set><p:cBhvr><p:cTn id="2"/></p:cBhvr><p:cBhvr><p:cTn id="3"/></p:cBhvr></p:set></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:par><x:wrapper><p:cTn id="1"/></x:wrapper></p:par></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:par><p:cTn id="1"/><p:cTn id="2"/></p:par></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:par><p:cTn id="1"><p:stCondLst><p:cond><p:tn val="2"/><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:stCondLst></p:cTn></p:par></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:set><p:cBhvr><p:cTn id="2"/></p:cBhvr><p:to><p:strVal val="visible"/></p:to><p:to><p:strVal val="hidden"/></p:to></p:set></p:tnLst></p:timing>"#,
        r#"<p:timing><p:tnLst><p:set><p:cBhvr><p:cTn id="2"/></p:cBhvr><p:to><p:strVal val="visible"/><p:intVal val="1"/></p:to></p:set></p:tnLst></p:timing>"#,
    ] {
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>{timing}</p:sld>"#
        );
        assert!(CT_Slide::from_xml(xml.as_bytes()).is_err(), "{timing}");
    }

    let valid = format!(
        r#"<p:timing xmlns:p="{P_NS}"><p:tnLst><p:par><p:cTn id="1" dur="100"/></p:par></p:tnLst></p:timing>"#
    );
    let mut timing = CT_Timing::from_xml(valid.as_bytes()).unwrap();
    let before = timing.clone();
    assert!(
        timing
            .set_node_duration(99, TimingDuration::Finite(200))
            .is_err()
    );
    assert_eq!(timing, before);
}

#[test]
fn foreign_and_nested_set_variants_do_not_become_typed_values() {
    for value in [
        r#"<p:to><x:strVal val="visible"/></p:to>"#,
        r#"<x:wrapper><p:to><p:strVal val="visible"/></p:to></x:wrapper>"#,
    ] {
        let xml = format!(
            r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst><p:set><p:cBhvr><p:cTn id="1"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl></p:cBhvr>{value}</p:set></p:tnLst></p:timing>"#
        );
        let timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
        let TimingNode::Set(set) = &timing.nodes()[0] else {
            panic!("expected set behavior");
        };
        assert_eq!(set.value, None, "{value}");
    }
}

#[test]
fn timing_targets_and_attribute_names_use_direct_expanded_name_boundaries() {
    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst><p:par><p:cTn id="1"><p:stCondLst><p:cond><x:wrapper><p:tn val="2"/></x:wrapper></p:cond></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="2"/><p:tgtEl><p:tn val="3"/><x:wrapper><p:spTgt spid="4"/></x:wrapper></p:tgtEl><p:attrNameLst>  <x:attrName>style.foreign</x:attrName>  <p:attrName> style.visibility </p:attrName>  </p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"#
    );
    let timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    let TimingNode::Parallel(root) = &timing.nodes()[0] else {
        panic!("expected parallel node");
    };
    assert_eq!(
        root.common.start_conditions[0].target,
        TimingTarget::Unsupported
    );
    assert_eq!(
        timing.condition_has_explicit_target(1, false, 0),
        Some(false)
    );
    let TimingNode::Set(set) = &root.common.children[0] else {
        panic!("expected set node");
    };
    assert_eq!(set.target, TimingTarget::Unsupported);
    assert_eq!(set.attribute_name.as_deref(), Some("style.visibility"));
    assert_eq!(timing.to_xml(), xml.as_bytes());
}

#[test]
fn timing_condition_target_presence_preserves_the_f213_projection() {
    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:q="{P_NS}"><p:tnLst><p:par><p:cTn id="7"><p:stCondLst><p:cond delay="0"/><q:cond delay="1"></q:cond><p:cond delay="2"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond><p:cond delay="3"><p:rtn val="all"/></p:cond><p:cond delay="4"><p:tgtEl><p:sndTgt/></p:tgtEl></p:cond></p:stCondLst><p:endCondLst><p:cond delay="5"/><p:cond delay="6"><p:rtn val="all"/></p:cond></p:endCondLst></p:cTn></p:par></p:tnLst></p:timing>"#
    );
    let timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    let TimingNode::Parallel(root) = &timing.nodes()[0] else {
        panic!("expected parallel timing node")
    };

    assert_eq!(
        root.common
            .start_conditions
            .iter()
            .map(|condition| &condition.target)
            .collect::<Vec<_>>(),
        vec![
            &TimingTarget::Unsupported,
            &TimingTarget::Unsupported,
            &TimingTarget::Slide,
            &TimingTarget::Unsupported,
            &TimingTarget::Unsupported,
        ]
    );
    assert_eq!(
        (0..5)
            .map(|index| timing.condition_has_explicit_target(7, false, index))
            .collect::<Vec<_>>(),
        vec![Some(false), Some(false), Some(true), Some(true), Some(true)]
    );
    assert_eq!(
        (0..2)
            .map(|index| timing.condition_has_explicit_target(7, true, index))
            .collect::<Vec<_>>(),
        vec![Some(false), Some(true)]
    );
    assert_eq!(timing.condition_has_explicit_target(7, false, 5), None);
    assert_eq!(timing.condition_has_explicit_target(8, false, 0), None);
    assert_eq!(timing.to_xml(), xml.as_bytes());
    let mut mutated = timing.clone();
    mutated
        .set_node_duration(7, TimingDuration::Finite(900))
        .unwrap();
    assert_eq!(
        mutated.condition_has_explicit_target(7, false, 3),
        Some(true)
    );
}

#[test]
fn unsupported_direct_targets_consume_choices_and_direct_cdata_names_are_typed() {
    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:r="{R_NS}" xmlns:x="urn:producer"><p:tnLst><p:par><p:cTn id="1"><p:stCondLst><p:cond><p:rtn val="all"/></p:cond></p:stCondLst><p:childTnLst><p:set><p:cBhvr><p:cTn id="2"/><p:tgtEl><p:sndTgt r:embed="rId7"/></p:tgtEl><p:attrNameLst><p:attrName>  <x:format><![CDATA[style.foreign]]></x:format><![CDATA[ style.visibility ]]></p:attrName></p:attrNameLst></p:cBhvr></p:set></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"#
    );
    let timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    let TimingNode::Parallel(root) = &timing.nodes()[0] else {
        panic!("expected parallel node");
    };
    assert_eq!(
        root.common.start_conditions[0].target,
        TimingTarget::Unsupported
    );
    assert_eq!(
        timing.condition_has_explicit_target(1, false, 0),
        Some(true)
    );
    let TimingNode::Set(set) = &root.common.children[0] else {
        panic!("expected set node");
    };
    assert_eq!(set.target, TimingTarget::Unsupported);
    assert_eq!(set.attribute_name.as_deref(), Some("style.visibility"));
    assert_eq!(timing.to_xml(), xml.as_bytes());

    for invalid in [
        r#"<p:par><p:cTn id="1"><p:stCondLst><p:cond><p:rtn val="all"/><p:tn val="2"/></p:cond></p:stCondLst></p:cTn></p:par>"#,
        r#"<p:set><p:cBhvr><p:cTn id="1"/><p:tgtEl><p:sndTgt r:embed="rId7"/><p:spTgt spid="4"/></p:tgtEl></p:cBhvr></p:set>"#,
        r#"<p:set><p:cBhvr><p:cTn id="1"/><p:tgtEl><p:inkTgt spid="3"/><p:sldTgt/></p:tgtEl></p:cBhvr></p:set>"#,
    ] {
        let xml = format!(
            r#"<p:timing xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:tnLst>{invalid}</p:tnLst></p:timing>"#
        );
        assert!(CT_Timing::from_xml(xml.as_bytes()).is_err(), "{invalid}");
    }
}

#[test]
fn direct_attribute_name_resolves_numeric_references_without_descending() {
    let xml = format!(
        r#"<p:timing xmlns:p="{P_NS}" xmlns:x="urn:producer"><p:tnLst><p:set><p:cBhvr><p:cTn id="1"/><p:attrNameLst><p:attrName><x:format>wrong&#x2E;name&amp;state</x:format>style&#x2E;visibility</p:attrName></p:attrNameLst></p:cBhvr></p:set></p:tnLst></p:timing>"#
    );
    let timing = CT_Timing::from_xml(xml.as_bytes()).unwrap();
    let TimingNode::Set(set) = &timing.nodes()[0] else {
        panic!("expected set node");
    };
    assert_eq!(set.attribute_name.as_deref(), Some("style.visibility"));
    assert_eq!(timing.to_xml(), xml.as_bytes());
}

#[test]
fn compatibility_transition_selects_supported_choice_or_fallback_only() {
    use rpptx_oxml::timing::TransitionSpeed;

    let xml = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="{MC_NS}" xmlns:u="urn:unsupported" xmlns:x="urn:producer"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><mc:AlternateContent><mc:Choice Requires="u"><p:transition spd="slow"><p:wipe dir="d"/></p:transition></mc:Choice><mc:Fallback><p:transition spd="fast"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent></p:sld>"#
    );
    let mut slide = CT_Slide::from_xml(xml.as_bytes()).unwrap();
    let transition = slide.transition.as_mut().unwrap();
    assert_eq!(transition.speed, Some(TransitionSpeed::Fast));
    assert_eq!(transition.effect, Some(TransitionEffect::Fade));
    transition.set_speed(TransitionSpeed::Medium).unwrap();
    let written = String::from_utf8(slide.to_xml().unwrap()).unwrap();
    assert!(written.contains(r#"<mc:Choice Requires="u"><p:transition spd="slow">"#));
    assert!(written.contains(r#"<mc:Fallback><p:transition spd="med">"#));

    let nested = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="{MC_NS}" xmlns:x="urn:producer"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><mc:AlternateContent><mc:Fallback><x:wrapper><p:transition/></x:wrapper></mc:Fallback></mc:AlternateContent></p:sld>"#
    );
    let slide = CT_Slide::from_xml(nested.as_bytes()).unwrap();
    assert!(slide.transition.is_none());
    assert!(
        String::from_utf8(slide.to_xml().unwrap())
            .unwrap()
            .contains("<x:wrapper><p:transition/></x:wrapper>")
    );
}

#[test]
fn selected_compatibility_branches_enforce_transition_and_effect_singletons() {
    let slide = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="{MC_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><mc:AlternateContent><mc:Fallback><p:transition/><p:transition/></mc:Fallback></mc:AlternateContent></p:sld>"#
    );
    assert!(CT_Slide::from_xml(slide.as_bytes()).is_err());

    let transition = format!(
        r#"<p:transition xmlns:p="{P_NS}" xmlns:mc="{MC_NS}"><mc:AlternateContent><mc:Fallback><p:fade/><p:wipe/></mc:Fallback></mc:AlternateContent></p:transition>"#
    );
    assert!(CT_SlideTransition::from_xml(transition.as_bytes()).is_err());
}

#[test]
fn transition_effect_choice_and_following_children_enforce_schema_sequence() {
    for transition in [
        r#"<p:transition><p:wipe/><p:push/></p:transition>"#,
        r#"<p:transition><p:sndAc/><p:wipe/></p:transition>"#,
        r#"<p:transition><p:extLst/><p:sndAc/></p:transition>"#,
    ] {
        let xml = transition.replace(
            "<p:transition",
            &format!(r#"<p:transition xmlns:p="{P_NS}""#),
        );
        assert!(
            CT_SlideTransition::from_xml(xml.as_bytes()).is_err(),
            "{xml}"
        );
    }
}

#[test]
fn slide_like_roots_reject_earlier_modelled_children_after_transition() {
    let slide = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:transition/><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"#
    );
    assert!(CT_Slide::from_xml(slide.as_bytes()).is_err());

    let layout = format!(
        r#"<p:sldLayout xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:transition><p:fade/></p:transition><p:hf/></p:sldLayout>"#
    );
    assert!(CT_SlideLayout::from_xml(layout.as_bytes()).is_err());

    for transition in [
        r#"<p:transition spd="fast"/>"#,
        r#"<p:transition></p:transition>"#,
        r#"<p:transition/><x:between/>"#,
    ] {
        let layout = format!(
            r#"<p:sldLayout xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld>{transition}<p:hf/></p:sldLayout>"#
        );
        assert!(CT_SlideLayout::from_xml(layout.as_bytes()).is_err());
    }
}

#[test]
fn empty_layout_transition_immediately_before_hf_is_typed_and_canonicalized() {
    let xml = format!(
        r#"<p:sldLayout xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:transition/><p:hf hdr="0" dt="0"/></p:sldLayout>"#
    );
    let layout = CT_SlideLayout::from_xml(xml.as_bytes()).unwrap();
    assert!(layout.transition.is_some());
    assert!(layout.header_footer.is_some());

    let written = layout.to_xml().unwrap();
    assert_order(&written, &["<p:hf", "<p:transition"]);
    let reparsed = CT_SlideLayout::from_xml(&written).unwrap();
    assert!(reparsed.transition.is_some());
    assert!(reparsed.header_footer.is_some());
    assert_eq!(reparsed.to_xml().unwrap(), written);
}

#[test]
fn transition_parameters_and_extension_duration_use_expanded_names() {
    let xml = format!(
        r#"<p:transition xmlns:p="{P_NS}" xmlns:z="http://schemas.microsoft.com/office/powerpoint/2010/main" z:dur="725" spd="med"><p:wipe dir="d" thruBlk="1"/></p:transition>"#
    );
    let transition = CT_SlideTransition::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(transition.duration_ms, Some(725));
    assert_eq!(transition.effect, Some(TransitionEffect::Wipe));
    assert_eq!(
        transition
            .effect_parameters
            .iter()
            .map(|parameter| (parameter.name.as_str(), parameter.value.as_str()))
            .collect::<Vec<_>>(),
        vec![("dir", "d"), ("thruBlk", "1")]
    );

    let unqualified = format!(r#"<p:transition xmlns:p="{P_NS}" dur="725"/>"#);
    assert_eq!(
        CT_SlideTransition::from_xml(unqualified.as_bytes())
            .unwrap()
            .duration_ms,
        None
    );
}

#[test]
fn transition_effect_compatibility_choice_uses_capabilities_and_fallback() {
    let fallback = format!(
        r#"<p:transition xmlns:p="{P_NS}" xmlns:mc="{MC_NS}" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><p159:morph option="byObject"/></mc:Choice><mc:Fallback><p:fade/></mc:Fallback></mc:AlternateContent></p:transition>"#
    );
    let transition = CT_SlideTransition::from_xml(fallback.as_bytes()).unwrap();
    assert_eq!(transition.effect, Some(TransitionEffect::Fade));
    assert_eq!(transition.morph, None);

    let supported = format!(
        r#"<p:transition xmlns:p="{P_NS}" xmlns:mc="{MC_NS}" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" xmlns:u="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="u"><p159:morph option="foreign"/></mc:Choice><mc:Choice Requires="p159"><p159:morph option="byObject"/></mc:Choice><mc:Fallback><p:fade/></mc:Fallback></mc:AlternateContent></p:transition>"#
    );
    let mut transition = CT_SlideTransition::from_xml(supported.as_bytes()).unwrap();
    assert_eq!(transition.effect, Some(TransitionEffect::Morph));
    assert_eq!(
        transition.morph.as_ref().unwrap().option.as_deref(),
        Some("byObject")
    );
    transition.set_morph_option("byWord").unwrap();
    let written = String::from_utf8(transition.to_xml()).unwrap();
    assert!(written.contains(r#"<p159:morph option="foreign"/>"#));
    assert!(written.contains(r#"<p159:morph option="byWord"/>"#));
}

#[test]
fn direct_morph_mutation_ignores_unsupported_nested_morph() {
    let nested = r#"<x:wrapper marker='A&#x20;B'><p159:morph option='nested' x:keep='C&#x20;D'/></x:wrapper>"#;
    let xml = format!(
        r#"<p:transition xmlns:p="{P_NS}" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" xmlns:x="urn:producer">{nested}<p159:morph option="byObject"/></p:transition>"#
    );
    let mut transition = CT_SlideTransition::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(transition.effect, Some(TransitionEffect::Morph));
    assert_eq!(
        transition.morph.as_ref().unwrap().option.as_deref(),
        Some("byObject")
    );

    transition.set_morph_option("byWord").unwrap();
    let written = String::from_utf8(transition.to_xml()).unwrap();
    assert!(written.contains(nested));
    assert!(written.contains(r#"<p159:morph option="byWord"/>"#));
    let reparsed = CT_SlideTransition::from_xml(written.as_bytes()).unwrap();
    assert_eq!(
        reparsed.morph.as_ref().unwrap().option.as_deref(),
        Some("byWord")
    );
}

#[test]
fn empty_layout_transition_remains_typed_and_enforces_root_sequence() {
    let prefix = format!(
        r#"<p:sldLayout xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:hf/>"#
    );
    let layout =
        CT_SlideLayout::from_xml(format!("{prefix}<p:transition/></p:sldLayout>").as_bytes())
            .unwrap();
    assert!(layout.transition.is_some());

    let duplicate = format!("{prefix}<p:transition/><p:transition/></p:sldLayout>");
    assert!(CT_SlideLayout::from_xml(duplicate.as_bytes()).is_err());

    let backward = format!("{prefix}<p:transition/><p:timing/></p:sldLayout>");
    assert!(CT_SlideLayout::from_xml(backward.as_bytes()).is_err());
}

#[test]
fn the_corpus_timeline_preserves_every_unsupported_sibling() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut coverage = TimelineCoverage::default();
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, xml) in &package.parts {
            if !part.ends_with(".xml") {
                continue;
            }
            if part.starts_with("/ppt/slides/slide") {
                coverage.slides += 1;
                let slide = CT_Slide::from_xml(xml)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                let written = slide.to_xml().unwrap();
                let reparsed = CT_Slide::from_xml(&written).unwrap();
                record_timeline_round_trip(
                    slide.timing.as_ref(),
                    slide.transition.as_ref(),
                    reparsed.timing.as_ref(),
                    reparsed.transition.as_ref(),
                    &mut coverage,
                );
                assert_eq!(slide, reparsed);
            } else if part.starts_with("/ppt/slideLayouts/slideLayout") {
                coverage.layouts += 1;
                let layout = CT_SlideLayout::from_xml(xml)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                let written = layout.to_xml().unwrap();
                let reparsed = CT_SlideLayout::from_xml(&written).unwrap();
                record_timeline_round_trip(
                    layout.timing.as_ref(),
                    layout.transition.as_ref(),
                    reparsed.timing.as_ref(),
                    reparsed.transition.as_ref(),
                    &mut coverage,
                );
                assert_eq!(layout, reparsed);
            } else if part.starts_with("/ppt/slideMasters/slideMaster") {
                coverage.masters += 1;
                let master = CT_SlideMaster::from_xml(xml)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                let written = master.to_xml().unwrap();
                let reparsed = CT_SlideMaster::from_xml(&written).unwrap();
                record_timeline_round_trip(
                    master.timing.as_ref(),
                    master.transition.as_ref(),
                    reparsed.timing.as_ref(),
                    reparsed.transition.as_ref(),
                    &mut coverage,
                );
                assert_eq!(master, reparsed);
            }
        }
    }
    assert!(coverage.slides > 0, "corpus has no slide coverage");
    assert!(coverage.layouts > 0, "corpus has no layout coverage");
    assert!(coverage.masters > 0, "corpus has no master coverage");
    assert!(coverage.timing_roots > 0, "corpus has no timing coverage");
    assert!(
        coverage.transitions > 0,
        "corpus has no transition coverage"
    );
    assert!(coverage.typed_nodes > 0, "corpus has no typed timing nodes");
    assert!(
        coverage.conditions > 0,
        "corpus has no typed timing triggers"
    );
    assert!(coverage.builds > 0, "corpus has no typed timing builds");
    assert!(coverage.set_values > 0, "corpus has no typed set values");
    assert!(
        coverage.transition_effect_parameters > 0,
        "corpus has no typed transition effect parameters"
    );
    assert!(
        coverage.timing_effect_parameters > 0,
        "corpus has no typed timing effect parameters"
    );
    assert!(
        coverage.compatibility_transitions > 0,
        "corpus has no typed compatibility transitions"
    );
    eprintln!("timeline corpus coverage: {coverage:?}");
}

#[derive(Debug, Default)]
struct TimelineCoverage {
    slides: usize,
    layouts: usize,
    masters: usize,
    timing_roots: usize,
    transitions: usize,
    typed_nodes: usize,
    conditions: usize,
    builds: usize,
    set_values: usize,
    transition_effect_parameters: usize,
    timing_effect_parameters: usize,
    compatibility_transitions: usize,
    unsupported_nodes: usize,
}

fn record_timeline_round_trip(
    timing: Option<&CT_Timing>,
    transition: Option<&CT_SlideTransition>,
    reparsed_timing: Option<&CT_Timing>,
    reparsed_transition: Option<&CT_SlideTransition>,
    coverage: &mut TimelineCoverage,
) {
    assert_eq!(
        timing.map(CT_Timing::raw_xml),
        reparsed_timing.map(CT_Timing::raw_xml)
    );
    assert_eq!(
        transition.map(CT_SlideTransition::raw_xml),
        reparsed_transition.map(CT_SlideTransition::raw_xml)
    );
    if let Some(timing) = timing {
        coverage.timing_roots += 1;
        coverage.builds += timing.builds().len();
        let mut unsupported = Vec::new();
        for node in timing.nodes() {
            record_timing_node(node, coverage, &mut unsupported);
        }
        let mut reparsed_unsupported = Vec::new();
        if let Some(reparsed) = reparsed_timing {
            let mut ignored = TimelineCoverage::default();
            for node in reparsed.nodes() {
                record_timing_node(node, &mut ignored, &mut reparsed_unsupported);
            }
        }
        assert_eq!(unsupported, reparsed_unsupported);
    }
    if let Some(transition) = transition {
        coverage.transitions += 1;
        coverage.transition_effect_parameters += transition.effect_parameters.len();
        if transition
            .raw_xml()
            .windows(b"AlternateContent".len())
            .any(|window| window == b"AlternateContent")
        {
            coverage.compatibility_transitions += 1;
        }
    }
}

fn record_timing_node(
    node: &TimingNode,
    coverage: &mut TimelineCoverage,
    unsupported: &mut Vec<Vec<u8>>,
) {
    let common = match node {
        TimingNode::Parallel(container) => Some(&container.common),
        TimingNode::Sequence(sequence) => {
            coverage.conditions +=
                sequence.previous_conditions.len() + sequence.next_conditions.len();
            Some(&sequence.common)
        }
        TimingNode::Set(set) => {
            coverage.set_values += usize::from(set.value.is_some());
            Some(&set.common)
        }
        TimingNode::Animate(animate) => Some(&animate.common),
        TimingNode::Effect(effect) => {
            coverage.timing_effect_parameters +=
                usize::from(effect.filter.is_some()) + usize::from(effect.transition.is_some());
            Some(&effect.common)
        }
        TimingNode::Motion(motion) => Some(&motion.common),
        TimingNode::Audio(media) | TimingNode::Video(media) => Some(&media.common.common),
        TimingNode::MediaCommand(command) => Some(&command.common),
        TimingNode::Unsupported(node) => {
            coverage.unsupported_nodes += 1;
            unsupported.push(node.raw_xml().to_vec());
            None
        }
    };
    if let Some(common) = common {
        coverage.typed_nodes += 1;
        coverage.conditions += common.start_conditions.len() + common.end_conditions.len();
        for child in &common.children {
            record_timing_node(child, coverage, unsupported);
        }
    }
}

#[test]
fn slide_size_mutation_preserves_kind_and_unmodelled_xml() {
    let xml = format!(
        r#"<q:presentation xmlns:q="{P_NS}" xmlns:x="urn:f115"><x:before/><q:sldSz cy="6858000" type="screen16x9" x:keep="yes" cx="12192000"><x:size> A &amp; B </x:size></q:sldSz><x:between/><q:notesSz cx="6858000" cy="9144000"/></q:presentation>"#
    );
    let mut presentation = CT_Presentation::from_xml(xml.as_bytes()).unwrap();

    presentation
        .set_slide_size(Emu(10_000_001), Emu(5_000_003))
        .unwrap();

    let written = presentation.to_xml().unwrap();
    let text = String::from_utf8(written.clone()).unwrap();
    assert!(
        text.contains(r#"<p:sldSz cx="10000001" cy="5000003" type="screen16x9" x:keep="yes">"#)
    );
    assert!(text.contains("<x:size> A &amp; B </x:size>"));
    assert_order(
        &written,
        &["<x:before", "<p:sldSz", "<x:between", "<p:notesSz"],
    );
    let reparsed = CT_Presentation::from_xml(&written).unwrap();
    let size = reparsed.slide_size.unwrap();
    assert_eq!((size.cx, size.cy), (Emu(10_000_001), Emu(5_000_003)));
    assert_eq!(size.kind.as_deref(), Some("screen16x9"));
}

#[test]
fn hidden_flag_uses_inverse_show_semantics() {
    let slide_xml = |show: &str| {
        format!(
            r#"<q:sld xmlns:q="{P_NS}" xmlns:x="urn:f115"{show} x:keep="root"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><x:tail/></q:sld>"#
        )
    };

    let missing = CT_Slide::from_xml(slide_xml("").as_bytes()).unwrap();
    let shown = CT_Slide::from_xml(slide_xml(r#" show="true""#).as_bytes()).unwrap();
    let hidden = CT_Slide::from_xml(slide_xml(r#" show="false""#).as_bytes()).unwrap();
    assert!(!missing.hidden());
    assert!(!shown.hidden());
    assert!(hidden.hidden());

    let mut slide = missing;
    slide.set_hidden(true);
    let hidden_xml = String::from_utf8(slide.to_xml().unwrap()).unwrap();
    assert!(hidden_xml.contains(r#" show="0""#));
    assert!(hidden_xml.contains(r#"x:keep="root""#));
    assert!(hidden_xml.contains("<x:tail/>"));
    slide.set_hidden(false);
    let shown_xml = String::from_utf8(slide.to_xml().unwrap()).unwrap();
    assert!(shown_xml.contains(r#" show="1""#));
}

#[test]
fn background_set_and_clear_preserve_theme_references_and_raw_xml() {
    let reference_xml = format!(
        r#"<q:sld xmlns:q="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:f115"><q:cSld><q:bg x:keep="theme"><q:bgRef idx="1001"><a:schemeClr val="bg1"/></q:bgRef><x:bgTail/></q:bg><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld></q:sld>"#
    );
    let mut reference = CT_Slide::from_xml(reference_xml.as_bytes()).unwrap();
    let reference_before = reference
        .common_slide_data
        .background
        .as_ref()
        .unwrap()
        .raw_xml()
        .to_vec();
    let fill = oxml_drawing::fill::Fill::from_xml(
        br#"<z:solidFill xmlns:z="http://schemas.openxmlformats.org/drawingml/2006/main"><z:srgbClr val="102030"/></z:solidFill>"#,
    )
    .unwrap();
    assert!(reference.set_background(fill.clone()).is_err());
    reference.clear_background();
    assert_eq!(
        reference.common_slide_data.background.unwrap().raw_xml(),
        reference_before
    );

    let direct_xml = format!(
        r#"<q:sld xmlns:q="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:f115"><q:cSld><x:before/><q:bg x:keep="container"><!--before-properties--><q:bgPr shadeToTitle="1" x:keep="properties"><x:before-fill/><!--before-fill--><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><!--after-fill--><a:effectLst><a:outerShdw blurRad="40000"><a:srgbClr val="010203"/></a:outerShdw></a:effectLst><x:after-fill/></q:bgPr><!--after-properties--><x:bg-extension/></q:bg><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree><x:after/></q:cSld></q:sld>"#
    );
    let mut direct = CT_Slide::from_xml(direct_xml.as_bytes()).unwrap();
    direct.set_background(fill).unwrap();
    let expected_background = br#"<q:bg x:keep="container"><!--before-properties--><q:bgPr shadeToTitle="1" x:keep="properties"><x:before-fill/><!--before-fill--><a:solidFill><a:srgbClr val="102030"/></a:solidFill><!--after-fill--><a:effectLst><a:outerShdw blurRad="40000"><a:srgbClr val="010203"/></a:outerShdw></a:effectLst><x:after-fill/></q:bgPr><!--after-properties--><x:bg-extension/></q:bg>"#;
    assert_eq!(
        direct
            .common_slide_data
            .background
            .as_ref()
            .unwrap()
            .raw_xml(),
        expected_background
    );
    let written = direct.to_xml().unwrap();
    assert!(
        written
            .windows(expected_background.len())
            .any(|window| window == expected_background)
    );
    assert_order(
        &written,
        &["<x:before/>", "<q:bg", "<p:spTree", "<x:after/>"],
    );
    direct.clear_background();
    let cleared = String::from_utf8(direct.to_xml().unwrap()).unwrap();
    assert!(!cleared.contains("<p:bg"));
    assert!(cleared.contains("<x:before/>"));
    assert!(cleared.contains("<x:after/>"));
}

#[test]
fn fresh_shape_ids_preserve_other_bytes_and_rewrite_connector_endpoints() {
    let xml = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="{MC_NS}"><p:cSld producer="kept"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name="root"/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="one"><x:raw xmlns:x="urn:f114"> A &amp; B </x:raw></p:cNvPr></p:nvSpPr><p:spPr/></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="link"/><p:cNvCxnSpPr><a:stCxn id="2" idx="0"/><a:endCxn id="9" idx="1"/></p:cNvCxnSpPr></p:nvCxnSpPr><p:spPr/></p:cxnSp><p:sp><p:nvSpPr><p:cNvPr id="9" name="two"/></p:nvSpPr><p:spPr/></p:sp><mc:AlternateContent><mc:Choice Requires="p14"><p:sp><p:nvSpPr><p:cNvPr id="20" name="choice"/></p:nvSpPr></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="21" name="choice link"/><p:cNvCxnSpPr><a:stCxn id="20" idx="0"/><a:endCxn id="20" idx="1"/></p:cNvCxnSpPr></p:nvCxnSpPr></p:cxnSp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="22" name="fallback"/></p:nvSpPr></p:sp><p:cxnSp><p:nvCxnSpPr><p:cNvPr id="23" name="fallback link"/><p:cNvCxnSpPr><a:stCxn id="22" idx="0"/><a:endCxn id="22" idx="1"/></p:cNvCxnSpPr></p:nvCxnSpPr></p:cxnSp></mc:Fallback></mc:AlternateContent></p:spTree></p:cSld></p:sld>"#
    );

    let rewritten = String::from_utf8(rewrite_shape_ids(xml.as_bytes()).unwrap()).unwrap();

    assert!(rewritten.contains(r#"<p:cNvPr id="1" name="root"/>"#));
    assert!(rewritten.contains(
        r#"<p:cNvPr id="3" name="one"><x:raw xmlns:x="urn:f114"> A &amp; B </x:raw></p:cNvPr>"#
    ));
    assert!(rewritten.contains(r#"<p:cNvPr id="4" name="link"/>"#));
    assert!(rewritten.contains(r#"<p:cNvPr id="5" name="two"/>"#));
    assert!(rewritten.contains(r#"<a:stCxn id="3" idx="0"/><a:endCxn id="5" idx="1"/>"#));
    assert!(rewritten.contains(r#"<p:cNvPr id="6" name="choice"/>"#));
    assert!(rewritten.contains(r#"<p:cNvPr id="8" name="choice link"/>"#));
    assert!(rewritten.contains(r#"<a:stCxn id="6" idx="0"/><a:endCxn id="6" idx="1"/>"#));
    assert!(rewritten.contains(r#"<p:cNvPr id="10" name="fallback"/>"#));
    assert!(rewritten.contains(r#"<p:cNvPr id="11" name="fallback link"/>"#));
    assert!(rewritten.contains(r#"<a:stCxn id="10" idx="0"/><a:endCxn id="10" idx="1"/>"#));
    assert!(rewritten.contains(r#"<p:cSld producer="kept">"#));
}

#[test]
fn slide_id_list_raw_children_follow_surviving_ids_after_collection_edits() {
    let xml = format!(
        r#"<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}" xmlns:x="urn:f114"><p:sldIdLst><x:before/><p:sldId id="256" r:id="rId1"/><x:between-one-two/><p:sldId id="257" r:id="rId2"/><x:between-two-three/><p:sldId id="258" r:id="rId3"/><x:after/></p:sldIdLst><p:notesSz cx="1" cy="1"/></p:presentation>"#
    );
    let mut presentation = CT_Presentation::from_xml(xml.as_bytes()).unwrap();
    let third = presentation.slide_ids.remove(2);
    presentation.slide_ids.insert(0, third);
    presentation.slide_ids.remove(1);
    presentation
        .slide_ids
        .insert(1, CT_SlideId::new(300, "rId4").unwrap());

    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();

    let third = written.find(r#"r:id="rId3""#).unwrap();
    let inserted = written.find(r#"r:id="rId4""#).unwrap();
    let between = written.find("<x:between-one-two/>").unwrap();
    let second = written.find(r#"r:id="rId2""#).unwrap();
    assert!(third < inserted && inserted < between && between < second);
    assert!(written.contains("<x:before/>"));
    assert!(written.contains("<x:between-two-three/>"));
    assert!(written.contains("<x:after/>"));
    assert!(!written.contains(r#"r:id="rId1""#));
}

#[test]
fn custom_show_removal_preserves_unrelated_presentation_xml() {
    let xml = format!(
        r#"<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}" xmlns:q="urn:f114" q:producer="kept"><p:sldIdLst><p:sldId id="256" r:id="gone"/><p:sldId id="257" r:id="kept"/></p:sldIdLst><p:custShowLst q:raw=" A &amp; B "><p:custShow name="one" id="1"><p:sldLst><p:sld r:id="gone"/><q:marker/><p:sld r:id="kept"/></p:sldLst></p:custShow><p:custShow name="two" id="2"><p:sldLst><p:sld r:id="gone"></p:sld></p:sldLst></p:custShow></p:custShowLst><p:notesSz cx="1" cy="1"/></p:presentation>"#
    );
    let mut presentation = CT_Presentation::from_xml(xml.as_bytes()).unwrap();

    presentation.remove_custom_show_slide("gone").unwrap();
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();

    assert_eq!(written.matches(r#"r:id="gone""#).count(), 1);
    assert_eq!(written.matches(r#"r:id="kept""#).count(), 2);
    assert!(written.contains(r#"<p:custShowLst q:raw=" A &amp; B ">"#));
    assert!(written.contains("<q:marker/>"));
}

#[derive(Debug)]
struct ManifestEntry<'a> {
    path: &'a str,
    producer: &'a str,
    digest: &'a str,
    url: &'a str,
}

fn manifest_entries() -> Vec<ManifestEntry<'static>> {
    let mut lines = MANIFEST.lines();
    assert_eq!(
        lines.next(),
        Some("path\tproducer\tsha256\turl"),
        "unexpected corpus manifest header"
    );
    lines
        .enumerate()
        .map(|(index, line)| {
            let fields: Vec<_> = line.split('\t').collect();
            assert_eq!(fields.len(), 4, "manifest line {}", index + 2);
            ManifestEntry {
                path: fields[0],
                producer: fields[1],
                digest: fields[2],
                url: fields[3],
            }
        })
        .collect()
}

fn workspace_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../..")
}

fn corpus_dir() -> PathBuf {
    workspace_root().join("corpus/pptx")
}

fn require_or_skip_corpus() -> Option<PathBuf> {
    let corpus = corpus_dir();
    if corpus.is_dir() {
        return Some(corpus);
    }
    assert_ne!(
        std::env::var_os("RDOCX_PPTX_CORPUS_REQUIRED").as_deref(),
        Some(std::ffi::OsStr::new("1")),
        "the pinned corpus is required but {} does not exist",
        corpus.display()
    );
    eprintln!(
        "corpus gate skipped because {} is absent, run scripts/fetch_pptx_corpus.py",
        corpus.display()
    );
    None
}

fn verify_fetched_corpus() {
    let status = Command::new("python3")
        .arg(workspace_root().join("scripts/fetch_pptx_corpus.py"))
        .arg("--check")
        .status()
        .expect("run corpus verifier");
    assert!(status.success(), "pinned corpus verification failed");
}

#[test]
fn rpptx_oxml_is_an_explicit_publication_candidate() {
    let manifest = include_str!("../Cargo.toml");
    assert!(manifest.contains("name = \"rpptx-oxml\""));
    assert!(manifest.contains("version = \"0.11.0\""));
    assert!(manifest.contains("publish = true"));
    assert_eq!(
        P_NS,
        "http://schemas.openxmlformats.org/presentationml/2006/main"
    );
    assert_eq!(P_PREFIX, "p");
    assert_eq!(
        R_NS,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    );
    assert_eq!(R_PREFIX, "r");
    assert_eq!(PRESENTATION_PART, "/ppt/presentation.xml");
}

#[test]
fn corpus_manifest_is_complete_and_verified() {
    let entries = manifest_entries();
    assert_eq!(entries.len(), EXPECTED_DECKS);
    let mut paths = HashSet::new();
    let mut urls = HashSet::new();
    let mut producers = HashSet::new();
    for entry in &entries {
        assert!(paths.insert(entry.path), "duplicate path {}", entry.path);
        assert!(urls.insert(entry.url), "duplicate URL {}", entry.url);
        assert!(
            entry.path.ends_with(".pptx"),
            "non-pptx path {}",
            entry.path
        );
        assert!(
            !entry.path.contains('/'),
            "nested corpus path {}",
            entry.path
        );
        assert!(
            !entry.producer.is_empty(),
            "missing producer for {}",
            entry.path
        );
        assert!(
            entry.digest.len() == 64 && entry.digest.bytes().all(|byte| byte.is_ascii_hexdigit()),
            "invalid SHA-256 for {}",
            entry.path
        );
        assert!(
            entry.url.starts_with("https://"),
            "non-HTTPS URL for {}",
            entry.path
        );
        producers.insert(entry.producer);
    }
    for expected in [
        "microsoft-powerpoint",
        "google-slides",
        "keynote-export",
        "libreoffice-impress",
    ] {
        assert!(producers.contains(expected), "missing producer {expected}");
    }
    if corpus_dir().is_dir() {
        verify_fetched_corpus();
    }
}

#[test]
fn code_built_presentation_package_round_trips_without_external_corpus() {
    let mut package = OpcPackage::with_main_part(
        "ppt/presentation.xml",
        oxml_opc::content_types::PRESENTATION,
    );
    package.set_part(PRESENTATION_PART, br#"<p:presentation/>"#.to_vec());
    let mut bytes = Cursor::new(Vec::new());
    package.write_to(&mut bytes).unwrap();
    bytes.set_position(0);
    let reopened = OpcPackage::from_reader(bytes).unwrap();
    assert_eq!(
        reopened.main_document_part().as_deref(),
        Some(PRESENTATION_PART)
    );
    assert_eq!(
        reopened.get_part(PRESENTATION_PART),
        Some(br#"<p:presentation/>"#.as_slice())
    );
}

#[test]
fn every_corpus_drawingml_table_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut coverage = TableCoverage::default();
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, xml) in &package.parts {
            if !part.ends_with(".xml") {
                continue;
            }
            for extracted in drawingml_tables(xml)
                .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
            {
                let table = CT_Table::from_xml_with_inherited_namespaces(
                    &extracted.xml,
                    &extracted.inherited_namespaces,
                )
                .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                let written = table.to_xml().unwrap();
                assert_eq!(
                    table,
                    CT_Table::from_xml(&written)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
                );
                coverage.add(&table);
            }
        }
    }
    assert_eq!(coverage.tables, 26);
    assert_eq!(coverage.rows, 152);
    assert_eq!(coverage.cells, 724);
    assert_eq!(coverage.merge_origins, 21);
    assert_eq!(coverage.horizontal_continuations, 126);
    assert_eq!(coverage.vertical_continuations, 9);
    eprintln!("DrawingML table corpus gate checked {coverage:?}");
}

#[derive(Debug, Default)]
struct TableCoverage {
    tables: usize,
    rows: usize,
    cells: usize,
    merge_origins: usize,
    horizontal_continuations: usize,
    vertical_continuations: usize,
}

impl TableCoverage {
    fn add(&mut self, table: &CT_Table) {
        self.tables += 1;
        self.rows += table.rows.len();
        for row in &table.rows {
            self.cells += row.cells.len();
            for cell in &row.cells {
                self.merge_origins += usize::from(cell.row_span > 1 || cell.grid_span > 1);
                self.horizontal_continuations += usize::from(cell.horizontal_merge);
                self.vertical_continuations += usize::from(cell.vertical_merge);
            }
        }
    }
}

struct ExtractedTable {
    xml: Vec<u8>,
    inherited_namespaces: Vec<(String, String)>,
}

fn drawingml_tables(xml: &[u8]) -> Result<Vec<ExtractedTable>, String> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut tables = Vec::new();
    let mut namespace_frames = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                let local_declarations = namespace_declarations(&element)?;
                if local_name(element.name().as_ref()) == b"tbl"
                    && element_namespace(
                        element.name().as_ref(),
                        &namespace_frames,
                        &local_declarations,
                    ) == Some(oxml_drawing::namespace::A_NS)
                {
                    tables.push(ExtractedTable {
                        xml: capture_element(&mut reader, &element)
                            .map_err(|error| format!("table scan failed: {error}"))?,
                        inherited_namespaces: effective_namespaces(&namespace_frames),
                    });
                } else {
                    namespace_frames.push(local_declarations);
                }
            }
            Ok(Event::Empty(element)) => {
                let local_declarations = namespace_declarations(&element)?;
                if local_name(element.name().as_ref()) == b"tbl"
                    && element_namespace(
                        element.name().as_ref(),
                        &namespace_frames,
                        &local_declarations,
                    ) == Some(oxml_drawing::namespace::A_NS)
                {
                    tables.push(ExtractedTable {
                        xml: capture_empty_element(&element)
                            .map_err(|error| format!("table scan failed: {error}"))?,
                        inherited_namespaces: effective_namespaces(&namespace_frames),
                    });
                }
            }
            Ok(Event::End(_)) => {
                namespace_frames
                    .pop()
                    .ok_or_else(|| "namespace stack underflow".to_owned())?;
            }
            Ok(Event::Eof) => return Ok(tables),
            Ok(_) => {}
            Err(error) => return Err(format!("XML scan failed: {error}")),
        }
        buffer.clear();
    }
}

fn namespace_declarations(element: &BytesStart<'_>) -> Result<Vec<(String, String)>, String> {
    let mut declarations = Vec::new();
    for attribute in element.attributes() {
        let attribute = attribute.map_err(|error| format!("namespace scan failed: {error}"))?;
        let key = attribute.key.as_ref();
        let prefix = if key == b"xmlns" {
            Some("")
        } else {
            key.strip_prefix(b"xmlns:")
                .map(|value| std::str::from_utf8(value))
                .transpose()
                .map_err(|error| format!("namespace prefix is not UTF-8: {error}"))?
        };
        if let Some(prefix) = prefix {
            let uri = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, element.decoder())
                .map_err(|error| format!("namespace URI decode failed: {error}"))?
                .into_owned();
            declarations.push((prefix.to_owned(), uri));
        }
    }
    Ok(declarations)
}

fn element_namespace<'a>(
    name: &[u8],
    frames: &'a [Vec<(String, String)>],
    local_declarations: &'a [(String, String)],
) -> Option<&'a str> {
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&[][..], |position| &name[..position]);
    let prefix = std::str::from_utf8(prefix).ok()?;
    local_declarations
        .iter()
        .rev()
        .chain(frames.iter().rev().flat_map(|frame| frame.iter().rev()))
        .find(|(candidate, _)| candidate == prefix)
        .map(|(_, uri)| uri.as_str())
}

fn effective_namespaces(frames: &[Vec<(String, String)>]) -> Vec<(String, String)> {
    let mut effective = Vec::new();
    for (prefix, uri) in frames.iter().flatten() {
        if let Some((_, existing_uri)) = effective
            .iter_mut()
            .find(|(existing_prefix, _)| existing_prefix == prefix)
        {
            *existing_uri = uri.clone();
        } else {
            effective.push((prefix.clone(), uri.clone()));
        }
    }
    effective
}

#[test]
fn drawingml_table_extraction_carries_ancestor_namespaces() {
    let xml = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:d="{}" xmlns:x="urn:producer" xmlns:a="urn:ancestor-producer"><p:cSld><d:tbl xmlns:a="{}"><d:tblGrid><d:gridCol w="100"/></d:tblGrid><d:tr h="200"><d:tc><x:extension/></d:tc></d:tr></d:tbl><x:tbl/></p:cSld></p:sld>"#,
        oxml_drawing::namespace::A_NS,
        oxml_drawing::namespace::A_NS
    );
    let extracted = drawingml_tables(xml.as_bytes()).unwrap();
    assert_eq!(extracted.len(), 1);
    let table = CT_Table::from_xml_with_inherited_namespaces(
        &extracted[0].xml,
        &extracted[0].inherited_namespaces,
    )
    .unwrap();
    let written = table.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.contains(r#"xmlns:x="urn:producer""#));
    assert!(text.contains("<x:extension/>"));
}

#[test]
fn all_corpus_decks_round_trip_opaquely() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    for entry in manifest_entries() {
        let path = corpus.join(entry.path);
        let original = OpcPackage::open(&path)
            .unwrap_or_else(|error| panic!("{}: open failed: {error}", entry.path));
        assert_eq!(
            original.main_document_part().as_deref(),
            Some(PRESENTATION_PART),
            "{}: wrong presentation main part",
            entry.path
        );
        let mut canonical = Cursor::new(Vec::new());
        original
            .write_to(&mut canonical)
            .unwrap_or_else(|error| panic!("{}: canonical save failed: {error}", entry.path));
        canonical.set_position(0);
        let saved = OpcPackage::from_reader(canonical)
            .unwrap_or_else(|error| panic!("{}: reopen failed: {error}", entry.path));
        assert_package_parts_equal(entry.path, &original, &saved);
    }
}

fn assert_package_parts_equal(deck: &str, original: &OpcPackage, saved: &OpcPackage) {
    assert_eq!(
        original.content_types.defaults, saved.content_types.defaults,
        "{deck}: content-type defaults changed"
    );
    assert_eq!(
        original.content_types.overrides, saved.content_types.overrides,
        "{deck}: content-type overrides changed"
    );
    assert_eq!(
        original.package_rels.items, saved.package_rels.items,
        "{deck}: package relationships changed"
    );
    assert_eq!(
        original.parts.len(),
        saved.parts.len(),
        "{deck}: part count changed"
    );
    for (part, bytes) in &original.parts {
        assert_eq!(
            saved.parts.get(part),
            Some(bytes),
            "{deck}: part {part} changed"
        );
    }
    assert_eq!(
        original.part_rels.len(),
        saved.part_rels.len(),
        "{deck}: relationship-part count changed"
    );
    for (part, relationships) in &original.part_rels {
        let saved_relationships = saved
            .part_rels
            .get(part)
            .unwrap_or_else(|| panic!("{deck}: relationship part {part} is missing"));
        assert_eq!(
            relationships.items, saved_relationships.items,
            "{deck}: relationships for {part} changed"
        );
    }
}

#[test]
fn carried_m7_drawingml_gate_passes_for_the_corpus() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut text_body_count = 0usize;
    let mut shape_properties_count = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path))
            .unwrap_or_else(|error| panic!("{}: open failed: {error}", entry.path));
        for (part, bytes) in &package.parts {
            if !part.ends_with(".xml") {
                continue;
            }
            let (text_bodies, shape_properties) = drawing_elements(bytes)
                .unwrap_or_else(|error| panic!("{} {part}: XML scan failed: {error}", entry.path));
            for xml in text_bodies {
                let parsed = CT_TextBody::from_xml(&xml).unwrap_or_else(|error| {
                    panic!("{} {part}: txBody parse failed: {error}", entry.path)
                });
                let written = parsed.to_xml().unwrap_or_else(|error| {
                    panic!("{} {part}: txBody write failed: {error}", entry.path)
                });
                let reparsed = CT_TextBody::from_xml(&written).unwrap_or_else(|error| {
                    panic!(
                        "{} {part}: written txBody parse failed: {error}",
                        entry.path
                    )
                });
                assert_eq!(parsed, reparsed, "{} {part}: txBody changed", entry.path);
                text_body_count += 1;
            }
            for xml in shape_properties {
                let parsed = CT_ShapeProperties::from_xml(&xml).unwrap_or_else(|error| {
                    panic!("{} {part}: spPr parse failed: {error}", entry.path)
                });
                let written = parsed.to_xml().unwrap_or_else(|error| {
                    panic!("{} {part}: spPr write failed: {error}", entry.path)
                });
                let reparsed = CT_ShapeProperties::from_xml(&written).unwrap_or_else(|error| {
                    panic!("{} {part}: written spPr parse failed: {error}", entry.path)
                });
                assert_eq!(parsed, reparsed, "{} {part}: spPr changed", entry.path);
                shape_properties_count += 1;
            }
        }
    }
    assert!(
        text_body_count > 0,
        "the corpus contained no txBody elements"
    );
    assert!(
        shape_properties_count > 0,
        "the corpus contained no spPr elements"
    );
    eprintln!(
        "DrawingML corpus gate checked {text_body_count} txBody and {shape_properties_count} spPr elements"
    );
}

fn drawing_elements(xml: &[u8]) -> Result<DrawingElements, String> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut text_bodies = Vec::new();
    let mut shape_properties = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if local_name(element.name().as_ref()) == b"txBody" => {
                text_bodies.push(
                    capture_element(&mut reader, &element).map_err(|error| error.to_string())?,
                );
            }
            Ok(Event::Empty(element)) if local_name(element.name().as_ref()) == b"txBody" => {
                text_bodies
                    .push(capture_empty_element(&element).map_err(|error| error.to_string())?);
            }
            Ok(Event::Start(element)) if local_name(element.name().as_ref()) == b"spPr" => {
                shape_properties.push(
                    capture_element(&mut reader, &element).map_err(|error| error.to_string())?,
                );
            }
            Ok(Event::Empty(element)) if local_name(element.name().as_ref()) == b"spPr" => {
                shape_properties
                    .push(capture_empty_element(&element).map_err(|error| error.to_string())?);
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(error.to_string()),
        }
        buffer.clear();
    }
    Ok((text_bodies, shape_properties))
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

#[test]
fn presentation_reads_any_prefix_and_writes_fixed_prefixes_in_schema_order() {
    let xml = format!(
        r#"<q:presentation xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:rel="{R_NS}" xmlns:v="urn:producer" marker="yes"><q:sldMasterIdLst><q:sldMasterId id="2147483648" rel:id="rId1"/></q:sldMasterIdLst><q:notesMasterIdLst><q:notesMasterId rel:id="rIdNotes"/></q:notesMasterIdLst><v:sldIdLst><v:sldId id="not-a-slide"/></v:sldIdLst><q:sldIdLst keep="list"><q:before/><q:sldId id="256" rel:id="rId2" v:id="producer-id"/><q:between/><q:sldId id="300" rel:id="rId3"/><q:after/></q:sldIdLst><q:sldSz cx="12192000" cy="6858000" type="screen16x9"/><q:notesSz cx="6858000" cy="9144000"/><q:custom value="a &amp; b"><q:n/></q:custom><q:defaultTextStyle/><q:extLst><q:ext uri="raw"/></q:extLst></q:presentation>"#
    );
    let parsed = CT_Presentation::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(parsed.slide_master_ids.len(), 1);
    assert_eq!(parsed.slide_master_ids[0].id, Some(2_147_483_648));
    assert_eq!(parsed.slide_master_ids[0].relationship_id, "rId1");
    assert_eq!(
        parsed
            .slide_ids
            .iter()
            .map(|slide| (slide.id, slide.relationship_id.as_str()))
            .collect::<Vec<_>>(),
        vec![(256, "rId2"), (300, "rId3")]
    );
    assert_eq!(parsed.slide_size.as_ref().unwrap().cx.0, 12_192_000);
    assert_eq!(parsed.notes_size.cy.0, 9_144_000);
    assert!(parsed.default_text_style.is_some());

    let written = parsed.to_xml().unwrap();
    let written_text = std::str::from_utf8(&written).unwrap();
    assert!(written_text.starts_with(&format!(
        r#"<p:presentation xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R_NS}""#
    )));
    assert!(written_text.contains(r#"<p:sldMasterId id="2147483648" r:id="rId1"/>"#));
    assert!(written_text.contains(r#"<p:sldId id="256" r:id="rId2" v:id="producer-id"/>"#));
    assert!(written_text.contains(r#"<p:defaultTextStyle/>"#));
    assert!(written_text.contains(r#"<q:custom value="a &amp; b"><q:n/></q:custom>"#));
    assert!(written_text.contains(
        r#"<q:notesMasterIdLst><q:notesMasterId rel:id="rIdNotes"/></q:notesMasterIdLst>"#
    ));
    assert!(written_text.contains(r#"<v:sldIdLst><v:sldId id="not-a-slide"/></v:sldIdLst>"#));
    assert!(written_text.contains(
        r#"<q:before/><p:sldId id="256" r:id="rId2" v:id="producer-id"/><q:between/><p:sldId id="300" r:id="rId3"/><q:after/>"#
    ));
    let master = written_text.find("<p:sldMasterIdLst").unwrap();
    let notes_master = written_text.find("<q:notesMasterIdLst").unwrap();
    let slides = written_text.find("<p:sldIdLst").unwrap();
    let slide_size = written_text.find("<p:sldSz").unwrap();
    let notes_size = written_text.find("<p:notesSz").unwrap();
    let defaults = written_text.find("<p:defaultTextStyle").unwrap();
    let extensions = written_text.find("<q:extLst").unwrap();
    assert!(
        master < notes_master
            && notes_master < slides
            && slides < slide_size
            && slide_size < notes_size
            && notes_size < defaults
            && defaults < extensions
    );
    assert_eq!(parsed, CT_Presentation::from_xml(&written).unwrap());
}

#[test]
fn slide_ids_preserve_order_and_defer_semantic_validation() {
    let valid = format!(
        r#"<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst><p:sldId id="2147483647" r:id="last"/><p:sldId id="256" r:id="first"/></p:sldIdLst><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    );
    let parsed = CT_Presentation::from_xml(valid.as_bytes()).unwrap();
    assert_eq!(
        parsed
            .slide_ids
            .iter()
            .map(|slide| slide.id)
            .collect::<Vec<_>>(),
        vec![2_147_483_647, 256]
    );

    for accepted_for_facade_validation in [
        r#"<p:sldId id="255" r:id="rId1"/>"#,
        r#"<p:sldId id="2147483648" r:id="rId1"/>"#,
    ] {
        let xml = format!(
            r#"<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst>{accepted_for_facade_validation}</p:sldIdLst><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
        );
        assert!(CT_Presentation::from_xml(xml.as_bytes()).is_ok());
    }

    let invalid = format!(
        r#"<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst><p:sldId id="not-an-integer" r:id="rId1"/></p:sldIdLst><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    );
    assert!(CT_Presentation::from_xml(invalid.as_bytes()).is_err());

    let duplicate = format!(
        r#"<p:presentation xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    );
    assert!(CT_Presentation::from_xml(duplicate.as_bytes()).is_ok());

    for invalid_size in [
        r#"<p:notesSz cy="9144000"/>"#,
        r#"<p:notesSz cx="broken" cy="9144000"/>"#,
        r#"<p:notesSz cx="0" cy="9144000"/>"#,
    ] {
        let xml = format!(r#"<p:presentation xmlns:p="{P_NS}">{invalid_size}</p:presentation>"#);
        assert!(
            CT_Presentation::from_xml(xml.as_bytes()).is_err(),
            "{invalid_size}"
        );
    }

    let wrong_root_namespace =
        br#"<x:presentation xmlns:x="urn:not-presentationml"><x:notesSz cx="1" cy="1"/></x:presentation>"#;
    assert!(CT_Presentation::from_xml(wrong_root_namespace).is_err());

    let conflicting_fixed_prefix = format!(
        r#"<q:presentation xmlns:q="{P_NS}" xmlns:p="urn:producer"><q:notesSz cx="6858000" cy="9144000"/></q:presentation>"#
    );
    assert!(CT_Presentation::from_xml(conflicting_fixed_prefix.as_bytes()).is_err());

    let nested_conflicting_prefix = format!(
        r#"<q:presentation xmlns:q="{P_NS}" xmlns:rel="{R_NS}"><q:sldIdLst xmlns:r="urn:producer"><q:sldId id="256" rel:id="rId1"/></q:sldIdLst><q:notesSz cx="6858000" cy="9144000"/></q:presentation>"#
    );
    assert!(CT_Presentation::from_xml(nested_conflicting_prefix.as_bytes()).is_err());
}

#[test]
fn zero_slide_template_round_trips() {
    let xml = format!(
        r#"<p:presentation xmlns:p="{P_NS}"><p:sldIdLst/><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    );
    let parsed = CT_Presentation::from_xml(xml.as_bytes()).unwrap();
    assert!(parsed.slide_ids.is_empty());
    let written = parsed.to_xml().unwrap();
    assert!(
        std::str::from_utf8(&written)
            .unwrap()
            .contains("<p:sldIdLst/>")
    );
    assert_eq!(parsed, CT_Presentation::from_xml(&written).unwrap());
}

#[test]
fn every_corpus_presentation_part_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut presentation_count = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path))
            .unwrap_or_else(|error| panic!("{}: open failed: {error}", entry.path));
        let xml = package
            .get_part(PRESENTATION_PART)
            .unwrap_or_else(|| panic!("{}: presentation part missing", entry.path));
        let parsed = CT_Presentation::from_xml(xml)
            .unwrap_or_else(|error| panic!("{}: presentation parse failed: {error}", entry.path));
        let written = parsed
            .to_xml()
            .unwrap_or_else(|error| panic!("{}: presentation write failed: {error}", entry.path));
        let reparsed = CT_Presentation::from_xml(&written).unwrap_or_else(|error| {
            panic!("{}: written presentation parse failed: {error}", entry.path)
        });
        assert_eq!(
            parsed, reparsed,
            "{}: presentation changed structurally",
            entry.path
        );
        presentation_count += 1;
    }
    assert_eq!(presentation_count, EXPECTED_DECKS);
    eprintln!("PresentationML corpus gate checked {presentation_count} presentation parts");
}

#[test]
fn slide_layout_and_master_write_their_own_schema_order() {
    let slide = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><x:transition/><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><p:transition/><p:timing/><p:extLst/></p:sld>"#
    );
    let layout = format!(
        r#"<p:sldLayout xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><p:hf/><p:timing/><p:transition/><p:extLst/></p:sldLayout>"#
    );
    let master = format!(
        r#"<p:sldMaster xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst/><p:transition/><p:timing/><p:hf/><x:beforeTextStyles><x:data/></x:beforeTextStyles><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles><p:extLst/></p:sldMaster>"#
    );

    let written_slide = CT_Slide::from_xml(slide.as_bytes())
        .unwrap()
        .to_xml()
        .unwrap();
    let written_layout = CT_SlideLayout::from_xml(layout.as_bytes())
        .unwrap()
        .to_xml()
        .unwrap();
    let written_master = CT_SlideMaster::from_xml(master.as_bytes())
        .unwrap()
        .to_xml()
        .unwrap();
    assert_order(
        &written_slide,
        &[
            "<x:transition",
            "<p:cSld",
            "<p:clrMapOvr",
            "<p:transition",
            "<p:timing",
            "<p:extLst",
        ],
    );
    assert_order(
        &written_layout,
        &[
            "<p:cSld",
            "<p:clrMapOvr",
            "<p:hf",
            "<p:timing",
            "<p:transition",
            "<p:extLst",
        ],
    );
    assert_order(
        &written_master,
        &[
            "<p:cSld",
            "<p:clrMap",
            "<p:sldLayoutIdLst",
            "<p:transition",
            "<p:timing",
            "<p:hf",
            "<x:beforeTextStyles",
            "<p:txStyles",
            "<p:extLst",
        ],
    );
}

#[test]
fn background_projection_preserves_the_source_subtree_verbatim() {
    let properties = format!(
        r#"<q:sld xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><q:cSld><q:bg x:keep="background"><q:bgPr shadeToTitle="1"><x:before/><d:gradFill rotWithShape="1"><d:gsLst><d:gs pos="0"><d:srgbClr val="102030"/></d:gs><d:gs pos="100000"><d:srgbClr val="D0E0F0"/></d:gs></d:gsLst><d:lin ang="0"/></d:gradFill><d:effectLst><x:effect/></d:effectLst><x:after/></q:bgPr></q:bg><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld></q:sld>"#
    );
    let expected = r#"<q:bg x:keep="background"><q:bgPr shadeToTitle="1"><x:before/><d:gradFill rotWithShape="1"><d:gsLst><d:gs pos="0"><d:srgbClr val="102030"/></d:gs><d:gs pos="100000"><d:srgbClr val="D0E0F0"/></d:gs></d:gsLst><d:lin ang="0"/></d:gradFill><d:effectLst><x:effect/></d:effectLst><x:after/></q:bgPr></q:bg>"#;
    let parsed = CT_Slide::from_xml(properties.as_bytes()).unwrap();
    let background = parsed.common_slide_data.background.as_ref().unwrap();
    assert_eq!(background.raw_xml(), expected.as_bytes());
    assert!(matches!(
        background.rendering(),
        BackgroundRendering::Properties(Some(fill))
            if matches!(fill.as_ref(), oxml_drawing::fill::Fill::Gradient(_))
    ));
    let written = parsed.to_xml().unwrap();
    assert!(
        written
            .windows(expected.len())
            .any(|window| window == expected.as_bytes())
    );
    assert_order(&written, &["<q:bg ", "<p:spTree"]);
    assert_eq!(parsed, CT_Slide::from_xml(&written).unwrap());

    let reference = format!(
        r#"<q:sld xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><q:cSld><q:bg><q:bgRef idx="1001"><x:before/><d:schemeClr val="accent2"><d:tint val="25000"/></d:schemeClr><x:after/></q:bgRef></q:bg><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld></q:sld>"#
    );
    let expected = r#"<q:bg><q:bgRef idx="1001"><x:before/><d:schemeClr val="accent2"><d:tint val="25000"/></d:schemeClr><x:after/></q:bgRef></q:bg>"#;
    let parsed = CT_Slide::from_xml(reference.as_bytes()).unwrap();
    let background = parsed.common_slide_data.background.as_ref().unwrap();
    assert_eq!(background.raw_xml(), expected.as_bytes());
    assert!(matches!(
        background.rendering(),
        BackgroundRendering::Reference {
            index: 1001,
            color: Some(_)
        }
    ));
    let written = parsed.to_xml().unwrap();
    assert!(
        written
            .windows(expected.len())
            .any(|window| window == expected.as_bytes())
    );
    assert_eq!(parsed, CT_Slide::from_xml(&written).unwrap());

    let wrong_namespace = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><p:cSld><p:bg><p:bgPr><x:gradFill/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"#
    );
    let parsed = CT_Slide::from_xml(wrong_namespace.as_bytes()).unwrap();
    assert!(matches!(
        parsed
            .common_slide_data
            .background
            .as_ref()
            .unwrap()
            .rendering(),
        BackgroundRendering::Properties(None)
    ));
}

#[test]
fn colour_maps_read_any_prefix_and_preserve_extensions() {
    let xml = format!(
        r#"<q:sld xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMapOvr x:keep="yes"><x:before/><d:overrideClrMapping bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent6" accent2="accent5" accent3="accent4" accent4="accent3" accent5="accent2" accent6="accent1" hlink="folHlink" folHlink="hlink" x:map="kept"><x:ext/></d:overrideClrMapping><x:after/></q:clrMapOvr></q:sld>"#
    );
    let parsed = CT_Slide::from_xml(xml.as_bytes()).unwrap();
    let override_map = parsed.color_map_override.as_ref().unwrap();
    let ColorMapOverrideKind::Override(map) = &override_map.kind else {
        panic!("expected an override colour map");
    };
    assert_eq!(
        map.theme_slot(ColorMapSlot::Background1),
        ThemeColorSlot::Dark1
    );
    assert_eq!(
        map.theme_slot(ColorMapSlot::Accent1),
        ThemeColorSlot::Accent6
    );
    let written = parsed.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.contains("<p:clrMapOvr x:keep=\"yes\"><x:before/><a:overrideClrMapping"));
    assert!(text.contains("x:map=\"kept\""));
    assert!(text.contains("<x:ext/></a:overrideClrMapping><x:after/>"));
    assert_eq!(parsed, CT_Slide::from_xml(&written).unwrap());
}

#[test]
fn modelled_nested_elements_reject_fixed_prefix_rebinding() {
    let colour_choice = format!(
        r#"<q:sld xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMapOvr><d:overrideClrMapping xmlns:a="urn:producer" bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"><a:raw/></d:overrideClrMapping></q:clrMapOvr></q:sld>"#
    );
    assert!(CT_Slide::from_xml(colour_choice.as_bytes()).is_err());

    let text_style = format!(
        r#"<q:sldMaster xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><q:txStyles><q:titleStyle xmlns:p="urn:producer"><p:raw/></q:titleStyle><q:bodyStyle/><q:otherStyle/></q:txStyles></q:sldMaster>"#
    );
    assert!(CT_SlideMaster::from_xml(text_style.as_bytes()).is_err());

    let default_text_level = format!(
        r#"<q:presentation xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:notesSz cx="6858000" cy="9144000"/><q:defaultTextStyle><d:lvl1pPr xmlns:a="urn:producer"><a:raw/></d:lvl1pPr></q:defaultTextStyle></q:presentation>"#
    );
    assert!(CT_Presentation::from_xml(default_text_level.as_bytes()).is_err());

    let default_text_descendant = format!(
        r#"<q:presentation xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:notesSz cx="6858000" cy="9144000"/><q:defaultTextStyle><d:lvl1pPr><d:defRPr xmlns:a="urn:producer"><a:raw/></d:defRPr></d:lvl1pPr></q:defaultTextStyle></q:presentation>"#
    );
    assert!(CT_Presentation::from_xml(default_text_descendant.as_bytes()).is_err());

    let master_text_level = format!(
        r#"<q:sldMaster xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><q:txStyles><q:titleStyle><d:lvl1pPr xmlns:a="urn:producer"><a:raw/></d:lvl1pPr></q:titleStyle><q:bodyStyle/><q:otherStyle/></q:txStyles></q:sldMaster>"#
    );
    assert!(CT_SlideMaster::from_xml(master_text_level.as_bytes()).is_err());

    let master_text_descendant = format!(
        r#"<q:sldMaster xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><q:txStyles><q:titleStyle><d:lvl1pPr><d:defRPr xmlns:a="urn:producer"><a:raw/></d:defRPr></d:lvl1pPr></q:titleStyle><q:bodyStyle/><q:otherStyle/></q:txStyles></q:sldMaster>"#
    );
    assert!(CT_SlideMaster::from_xml(master_text_descendant.as_bytes()).is_err());

    let required_shell = format!(
        r#"<q:spTree xmlns:q="{P_NS}"><q:nvGrpSpPr xmlns:p="urn:producer"><p:raw/></q:nvGrpSpPr><q:grpSpPr/></q:spTree>"#
    );
    assert!(CT_ShapeTree::from_xml(required_shell.as_bytes()).is_err());

    let group_transform = format!(
        r#"<q:spTree xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:nvGrpSpPr/><q:grpSpPr><d:xfrm xmlns:a="urn:producer"><a:raw/></d:xfrm></q:grpSpPr></q:spTree>"#
    );
    assert!(CT_ShapeTree::from_xml(group_transform.as_bytes()).is_err());

    let nested_group = format!(
        r#"<q:spTree xmlns:q="{P_NS}"><q:nvGrpSpPr/><q:grpSpPr/><q:grpSp xmlns:p="urn:producer"><q:nvGrpSpPr/><q:grpSpPr/><p:raw/></q:grpSp></q:spTree>"#
    );
    assert!(CT_ShapeTree::from_xml(nested_group.as_bytes()).is_err());
}

#[test]
fn common_slide_data_uses_the_typed_shape_tree() {
    let producer = r#"<x:producer xmlns:x="urn:producer"><x:data/></x:producer>"#;
    let xml = format!(
        r#"<q:sld xmlns:q="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cSld name="Named"><q:bg><q:bgPr/></q:bg><q:spTree marker="a &amp; b"><q:nvGrpSpPr/><q:grpSpPr/>{producer}</q:spTree><q:extLst/></q:cSld></q:sld>"#
    );
    let parsed = CT_Slide::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(parsed.common_slide_data.name.as_deref(), Some("Named"));
    assert!(parsed.common_slide_data.shape_tree.children.is_empty());
    let written = parsed.to_xml().unwrap();
    assert!(
        written
            .windows(producer.len())
            .any(|window| window == producer.as_bytes())
    );
    assert_eq!(parsed, CT_Slide::from_xml(&written).unwrap());
}

#[test]
fn every_corpus_slide_layout_and_master_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut counts = [0usize; 3];
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            if !package.parts.contains_key(part) {
                continue;
            }
            let xml = package
                .get_part(part)
                .unwrap_or_else(|| panic!("{} {part}: missing part", entry.path));
            match content_type.as_str() {
                content_types::SLIDE => {
                    let parsed = CT_Slide::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    let written = parsed.to_xml().unwrap();
                    assert_eq!(parsed, CT_Slide::from_xml(&written).unwrap());
                    counts[0] += 1;
                }
                content_types::SLIDE_LAYOUT => {
                    let parsed = CT_SlideLayout::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    let written = parsed.to_xml().unwrap();
                    assert_eq!(parsed, CT_SlideLayout::from_xml(&written).unwrap());
                    counts[1] += 1;
                }
                content_types::SLIDE_MASTER => {
                    let parsed = CT_SlideMaster::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    let written = parsed.to_xml().unwrap();
                    assert_eq!(parsed, CT_SlideMaster::from_xml(&written).unwrap());
                    counts[2] += 1;
                }
                _ => {}
            }
        }
    }
    assert!(
        counts.iter().all(|count| *count > 0),
        "part counts: {counts:?}"
    );
    eprintln!("PresentationML corpus gate checked {counts:?} slide, layout, and master parts");
}

#[test]
fn corpus_part_relationship_counts_are_valid() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            if !package.parts.contains_key(part) {
                continue;
            }
            let required_type = match content_type.as_str() {
                content_types::SLIDE => Some(rel_types::SLIDE_LAYOUT),
                content_types::SLIDE_LAYOUT => Some(rel_types::SLIDE_MASTER),
                content_types::SLIDE_MASTER => Some(rel_types::THEME),
                _ => None,
            };
            let Some(required_type) = required_type else {
                continue;
            };
            let count = package
                .get_part_rels(part)
                .map(|rels| {
                    rels.items
                        .iter()
                        .filter(|rel| rel.rel_type == required_type)
                        .count()
                })
                .unwrap_or(0);
            assert_eq!(count, 1, "{} {part}: {required_type}", entry.path);
        }
    }
}

fn assert_order(xml: &[u8], tags: &[&str]) {
    let text = std::str::from_utf8(xml).unwrap();
    let positions: Vec<_> = tags
        .iter()
        .map(|tag| {
            text.find(tag)
                .unwrap_or_else(|| panic!("missing {tag}: {text}"))
        })
        .collect();
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]), "{text}");
}

#[test]
fn shape_tree_requires_non_visual_and_group_properties_in_order() {
    let valid = format!(
        r#"<q:spTree xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:nvGrpSpPr marker="kept"/><q:grpSpPr><d:xfrm><d:off x="10" y="20"/><d:ext cx="30" cy="40"/><d:chOff x="1" y="2"/><d:chExt cx="3" cy="4"/></d:xfrm><d:solidFill><d:srgbClr val="112233"/></d:solidFill></q:grpSpPr><q:sp><q:nvSpPr><q:cNvPr/><q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr/></q:sp></q:spTree>"#
    );
    let parsed = CT_ShapeTree::from_xml(valid.as_bytes()).unwrap();
    let transform = parsed.group_transform().unwrap();
    assert_eq!(transform.offset.unwrap().x.0, 10);
    assert_eq!(transform.child_extent.unwrap().cy.0, 4);
    let written = parsed.to_xml().unwrap();
    assert_order(&written, &["<p:nvGrpSpPr", "<p:grpSpPr", "<p:sp xmlns"]);
    assert!(std::str::from_utf8(&written).unwrap().contains("<a:xfrm>"));
    assert_eq!(parsed, CT_ShapeTree::from_xml(&written).unwrap());

    for invalid in [
        format!(r#"<p:spTree xmlns:p="{P_NS}"><p:grpSpPr/></p:spTree>"#),
        format!(r#"<p:spTree xmlns:p="{P_NS}"><p:nvGrpSpPr/></p:spTree>"#),
        format!(r#"<p:spTree xmlns:p="{P_NS}"><p:grpSpPr/><p:nvGrpSpPr/></p:spTree>"#),
        format!(
            r#"<p:spTree xmlns:p="{P_NS}"><p:nvGrpSpPr/><p:grpSpPr/><p:nvGrpSpPr/></p:spTree>"#
        ),
        format!(r#"<p:spTree xmlns:p="{P_NS}"><p:nvGrpSpPr/><p:grpSpPr/><p:grpSpPr/></p:spTree>"#),
    ] {
        assert!(
            CT_ShapeTree::from_xml(invalid.as_bytes()).is_err(),
            "{invalid}"
        );
    }

    let tree_xml = format!(
        r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr/><p:xfrm/><a:graphic><a:graphicData uri="urn:producer"><x:data name="mc"/></a:graphicData></a:graphic></p:graphicFrame></p:spTree>"#
    );
    let tree = CT_ShapeTree::from_xml(tree_xml.as_bytes()).unwrap();
    assert_eq!(
        tree,
        CT_ShapeTree::from_xml(&tree.to_xml().unwrap()).unwrap()
    );
}

#[test]
fn all_six_child_variants_keep_document_order() {
    let payloads = [
        r#"<p:sp marker="shape"><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>"#,
        r#"<p:pic marker="picture"><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>"#,
        r#"<p:graphicFrame marker="frame"><p:nvGraphicFramePr/><p:xfrm/><a:graphic><a:graphicData uri="urn:frame"><x:data/></a:graphicData></a:graphic></p:graphicFrame>"#,
        r#"<p:cxnSp marker="connector"><p:nvCxnSpPr><p:cNvPr id="2" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#,
        r#"<mc:AlternateContent marker="alternate"><mc:Choice Requires="p14"><p:sp/></mc:Choice></mc:AlternateContent>"#,
    ];
    let xml = format!(
        r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:producer"><p:nvGrpSpPr/><p:grpSpPr/>{}<x:between/>{}<p:grpSp><p:nvGrpSpPr/><p:grpSpPr/><p:sp marker="nested"><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></p:grpSp>{}{}{}</p:spTree>"#,
        payloads[0], payloads[1], payloads[2], payloads[3], payloads[4]
    );
    let parsed = CT_ShapeTree::from_xml(xml.as_bytes()).unwrap();
    assert!(matches!(parsed.children[0], ShapeTreeChild::Shape(_)));
    assert!(matches!(parsed.children[1], ShapeTreeChild::Picture(_)));
    assert!(matches!(parsed.children[2], ShapeTreeChild::GroupShape(_)));
    assert!(matches!(
        parsed.children[3],
        ShapeTreeChild::GraphicFrame(_)
    ));
    assert!(matches!(parsed.children[4], ShapeTreeChild::Connector(_)));
    assert!(matches!(
        parsed.children[5],
        ShapeTreeChild::AlternateContent(_)
    ));
    let written = parsed.to_xml().unwrap();
    assert_order(
        &written,
        &[
            "marker=\"shape\"",
            "<x:between",
            "marker=\"picture\"",
            "marker=\"nested\"",
            "marker=\"frame\"",
            "marker=\"connector\"",
            "marker=\"alternate\"",
        ],
    );
    assert!(std::str::from_utf8(&written).unwrap().contains("<x:data/>"));
    let alternate_content = payloads[4].as_bytes();
    assert!(
        written
            .windows(alternate_content.len())
            .any(|window| window == alternate_content)
    );
    assert_eq!(parsed, CT_ShapeTree::from_xml(&written).unwrap());
}

#[test]
fn alternate_content_selects_fallback_without_parsing_choices() {
    let raw = r#"<mc:AlternateContent marker="selected"><mc:Choice Requires="p14"><p:sp/></mc:Choice><mc:Fallback><p:sp marker="fallback-shape"><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><x:ignored/><p:pic marker="fallback-picture"><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic><p:cxnSp marker="fallback-connector"><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp></mc:Fallback></mc:AlternateContent>"#;
    let tree = alternate_content_tree(raw.as_bytes(), &[]).unwrap();
    let ShapeTreeChild::AlternateContent(alternate) = &tree.children[0] else {
        panic!("expected typed AlternateContent");
    };
    let selected = alternate.selected_fallback().unwrap();
    assert_eq!(selected.len(), 3);
    assert!(matches!(selected[0], ShapeTreeChild::Shape(_)));
    assert!(matches!(selected[1], ShapeTreeChild::Picture(_)));
    assert!(matches!(selected[2], ShapeTreeChild::Connector(_)));
}

#[test]
fn alternate_content_without_fallback_preserves_choices_and_selects_none() {
    let raw = r#"<mc:AlternateContent marker="choice-only"><mc:Choice Requires="p14"><p:sp/></mc:Choice></mc:AlternateContent>"#;
    let tree = alternate_content_tree(raw.as_bytes(), &[]).unwrap();
    let ShapeTreeChild::AlternateContent(alternate) = &tree.children[0] else {
        panic!("expected typed AlternateContent");
    };
    assert_eq!(alternate.raw_xml(), raw.as_bytes());
    assert_eq!(alternate.selected_fallback(), None);
    assert!(
        tree.to_xml()
            .unwrap()
            .windows(raw.len())
            .any(|part| part == raw.as_bytes())
    );
}

#[test]
fn alternate_content_chart_choice_exposes_non_visual_identity() {
    let raw = r#"<mc:AlternateContent><mc:Choice Requires="p14"><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="42" name="!!Chart &amp; Data"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic></p:graphicFrame></mc:Choice><mc:Fallback/></mc:AlternateContent>"#;
    let tree = alternate_content_tree(raw.as_bytes(), &[]).unwrap();
    let wrapped = &tree.children[0];

    assert_eq!(wrapped.non_visual_id(), Some(42));
    assert_eq!(wrapped.non_visual_name().as_deref(), Some("!!Chart & Data"));
    assert_eq!(
        tree.to_xml().unwrap(),
        CT_ShapeTree::from_xml(&tree.to_xml().unwrap())
            .unwrap()
            .to_xml()
            .unwrap()
    );
}

#[test]
fn alternate_content_resolves_branch_namespaces_and_rejects_duplicate_fallbacks() {
    let raw = r#"<m:AlternateContent><m:Choice Requires="q"><m:Fallback><q:sp/></m:Fallback></m:Choice><x:Fallback><q:sp/></x:Fallback><m:Fallback><q:sp marker="alias"><q:nvSpPr><q:cNvPr/><q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr/></q:sp></m:Fallback></m:AlternateContent>"#;
    let inherited = [
        (
            "m",
            "http://schemas.openxmlformats.org/markup-compatibility/2006",
        ),
        ("q", P_NS),
        ("x", "urn:not-mc"),
    ];
    let tree = alternate_content_tree(raw.as_bytes(), &inherited).unwrap();
    let ShapeTreeChild::AlternateContent(alternate) = &tree.children[0] else {
        panic!("expected typed AlternateContent");
    };
    let selected = alternate.selected_fallback().unwrap();
    assert_eq!(selected.len(), 1);
    assert!(matches!(selected[0], ShapeTreeChild::Shape(_)));

    let empty = r#"<m:AlternateContent><m:Fallback/></m:AlternateContent>"#;
    let empty_tree = alternate_content_tree(empty.as_bytes(), &inherited).unwrap();
    let ShapeTreeChild::AlternateContent(empty_alternate) = &empty_tree.children[0] else {
        panic!("expected empty typed AlternateContent fallback");
    };
    assert_eq!(empty_alternate.selected_fallback(), Some(&[][..]));

    let duplicate = r#"<m:AlternateContent><m:Fallback/><m:Fallback/></m:AlternateContent>"#;
    assert!(alternate_content_tree(duplicate.as_bytes(), &inherited).is_err());
}

#[test]
fn alternate_content_fallback_selection_does_not_change_stored_alternatives() {
    let raw = r#"<mc:AlternateContent x:marker="a &amp; b">
  <?producer keep?>
  <!-- choice stays opaque -->
  <mc:Choice Requires="p14"><p:sp x:data="&quot;raw&quot;"/></mc:Choice>
  <mc:Fallback><x:before/><p:sp marker="selected"><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><x:after/></mc:Fallback>
</mc:AlternateContent>"#;
    let tree = alternate_content_tree(raw.as_bytes(), &[]).unwrap();
    let ShapeTreeChild::AlternateContent(alternate) = &tree.children[0] else {
        panic!("expected typed AlternateContent");
    };
    assert_eq!(alternate.raw_xml(), raw.as_bytes());
    assert_eq!(alternate.selected_fallback().unwrap().len(), 1);
    let written = tree.to_xml().unwrap();
    assert!(
        written
            .windows(raw.len())
            .any(|part| part == raw.as_bytes())
    );
    let reparsed = CT_ShapeTree::from_xml(&written).unwrap();
    let ShapeTreeChild::AlternateContent(reparsed_alternate) = &reparsed.children[0] else {
        panic!("expected typed AlternateContent after round-trip");
    };
    assert_eq!(reparsed_alternate.raw_xml(), raw.as_bytes());
}

#[test]
fn alternate_content_fallback_keeps_shape_tree_order_inside_nested_groups() {
    let raw = r#"<mc:AlternateContent marker="alternate"><mc:Fallback><p:sp marker="first"><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:grpSp marker="second"><p:nvGrpSpPr/><p:grpSpPr/><p:pic marker="nested-picture"><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></p:grpSp><p:cxnSp marker="third"><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp></mc:Fallback></mc:AlternateContent>"#;
    let outer = format!(
        r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><p:nvGrpSpPr/><p:grpSpPr/><p:grpSp marker="outer"><p:nvGrpSpPr/><p:grpSpPr/>{raw}</p:grpSp><p:sp marker="after"><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp></p:spTree>"#
    );
    let tree = CT_ShapeTree::from_xml(outer.as_bytes()).unwrap();
    let ShapeTreeChild::GroupShape(group) = &tree.children[0] else {
        panic!("expected outer group");
    };
    let ShapeTreeChild::AlternateContent(alternate) = &group.children[0] else {
        panic!("expected AlternateContent inside group");
    };
    let selected = alternate.selected_fallback().unwrap();
    assert!(matches!(selected[0], ShapeTreeChild::Shape(_)));
    let ShapeTreeChild::GroupShape(selected_group) = &selected[1] else {
        panic!("expected selected nested group");
    };
    assert!(matches!(
        selected_group.children[0],
        ShapeTreeChild::Picture(_)
    ));
    assert!(matches!(selected[2], ShapeTreeChild::Connector(_)));
    assert!(matches!(tree.children[1], ShapeTreeChild::Shape(_)));
    let written = tree.to_xml().unwrap();
    assert_order(
        &written,
        &[
            "marker=\"outer\"",
            "marker=\"alternate\"",
            "marker=\"after\"",
        ],
    );
    assert_eq!(tree, CT_ShapeTree::from_xml(&written).unwrap());
}

#[test]
fn every_corpus_alternate_content_subtree_round_trips_byte_identically() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut coverage = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, xml) in &package.parts {
            if !part.ends_with(".xml") {
                continue;
            }
            for alternate in alternate_content_subtrees(xml)
                .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
            {
                let inherited: Vec<_> = alternate
                    .inherited_namespaces
                    .iter()
                    .map(|(prefix, uri)| (prefix.as_str(), uri.as_str()))
                    .collect();
                let tree = alternate_content_tree(&alternate.xml, &inherited)
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                let ShapeTreeChild::AlternateContent(model) = &tree.children[0] else {
                    panic!("{} {part}: expected typed AlternateContent", entry.path);
                };
                assert_eq!(model.raw_xml(), alternate.xml);
                let written = tree.to_xml().unwrap();
                assert!(
                    written
                        .windows(alternate.xml.len())
                        .any(|part| part == alternate.xml)
                );
                assert_eq!(tree, CT_ShapeTree::from_xml(&written).unwrap());
                coverage += 1;
            }
        }
    }
    assert!(
        coverage > 0,
        "the pinned corpus contained no MC AlternateContent"
    );
    eprintln!("AlternateContent corpus gate checked {coverage} subtrees");
}

fn alternate_content_tree(
    raw: &[u8],
    inherited: &[(&str, &str)],
) -> Result<CT_ShapeTree, oxml_core::OxmlError> {
    let mut attributes = vec![
        ("xmlns:p".to_owned(), P_NS.to_owned()),
        (
            "xmlns:a".to_owned(),
            oxml_drawing::namespace::A_NS.to_owned(),
        ),
        ("xmlns:r".to_owned(), R_NS.to_owned()),
        ("xmlns:mc".to_owned(), MC_NS.to_owned()),
    ];
    for (prefix, uri) in inherited {
        if matches!(*prefix, "p" | "a" | "r" | "mc") {
            continue;
        }
        let name = if prefix.is_empty() {
            "xmlns".to_owned()
        } else {
            format!("xmlns:{prefix}")
        };
        attributes.push((name, (*uri).to_owned()));
    }

    let mut writer = Writer::new(Vec::new());
    let mut root = BytesStart::new("p:spTree");
    for (name, value) in &attributes {
        root.push_attribute((name.as_str(), value.as_str()));
    }
    writer.write_event(Event::Start(root))?;
    writer.write_event(Event::Empty(BytesStart::new("p:nvGrpSpPr")))?;
    writer.write_event(Event::Empty(BytesStart::new("p:grpSpPr")))?;
    writer.get_mut().extend_from_slice(raw);
    writer.write_event(Event::End(BytesEnd::new("p:spTree")))?;
    CT_ShapeTree::from_xml(&writer.into_inner())
}

struct ExtractedAlternateContent {
    xml: Vec<u8>,
    inherited_namespaces: Vec<(String, String)>,
}

fn alternate_content_subtrees(xml: &[u8]) -> Result<Vec<ExtractedAlternateContent>, String> {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut alternatives = Vec::new();
    let mut namespace_frames = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                let local_declarations = namespace_declarations(&element)?;
                if local_name(element.name().as_ref()) == b"AlternateContent"
                    && element_namespace(
                        element.name().as_ref(),
                        &namespace_frames,
                        &local_declarations,
                    ) == Some(MC_NS)
                {
                    alternatives.push(ExtractedAlternateContent {
                        xml: capture_element(&mut reader, &element)
                            .map_err(|error| format!("AlternateContent scan failed: {error}"))?,
                        inherited_namespaces: effective_namespaces(&namespace_frames),
                    });
                } else {
                    namespace_frames.push(local_declarations);
                }
            }
            Ok(Event::Empty(element)) => {
                let local_declarations = namespace_declarations(&element)?;
                if local_name(element.name().as_ref()) == b"AlternateContent"
                    && element_namespace(
                        element.name().as_ref(),
                        &namespace_frames,
                        &local_declarations,
                    ) == Some(MC_NS)
                {
                    alternatives.push(ExtractedAlternateContent {
                        xml: capture_empty_element(&element)
                            .map_err(|error| format!("AlternateContent scan failed: {error}"))?,
                        inherited_namespaces: effective_namespaces(&namespace_frames),
                    });
                }
            }
            Ok(Event::End(_)) => {
                namespace_frames
                    .pop()
                    .ok_or_else(|| "namespace stack underflow".to_owned())?;
            }
            Ok(Event::Eof) => return Ok(alternatives),
            Ok(_) => {}
            Err(error) => return Err(format!("XML scan failed: {error}")),
        }
        buffer.clear();
    }
}

#[test]
fn connector_with_start_and_end_connections_round_trips() {
    let xml = format!(
        r#"<p:cxnSp xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvCxnSpPr><p:cNvPr id="2" name="Connector 1"/><p:cNvCxnSpPr><a:stCxn id="3" idx="0"/><a:endCxn id="4" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#
    );
    let parsed = CT_ConnectionShape::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(parsed.start_connection.as_ref().unwrap().id, 3);
    assert_eq!(parsed.start_connection.as_ref().unwrap().idx, 0);
    assert_eq!(parsed.end_connection.as_ref().unwrap().id, 4);
    assert_eq!(parsed.end_connection.as_ref().unwrap().idx, 1);
    assert_eq!(
        parsed,
        CT_ConnectionShape::from_xml(&parsed.to_xml().unwrap()).unwrap()
    );
}

#[test]
fn connector_reads_any_prefix_and_writes_fixed_prefixes_in_schema_order() {
    let xml = format!(
        r#"<q:cxnSp xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:nvCxnSpPr><q:cNvPr id="2" name="Connector"/><q:cNvCxnSpPr><d:stCxn id="3" idx="2"/><d:endCxn id="4" idx="1"/></q:cNvCxnSpPr><q:nvPr/></q:nvCxnSpPr><q:spPr/></q:cxnSp>"#
    );
    let written = CT_ConnectionShape::from_xml(xml.as_bytes())
        .unwrap()
        .to_xml()
        .unwrap();
    assert_order(
        &written,
        &[
            "<p:nvCxnSpPr",
            "<p:cNvPr",
            "<p:cNvCxnSpPr",
            "<a:stCxn",
            "<a:endCxn",
            "<p:nvPr",
            "<p:spPr",
        ],
    );
}

#[test]
fn connector_requires_non_visual_and_shape_properties_in_schema_order() {
    let documents = [
        format!(r#"<p:cxnSp xmlns:p="{P_NS}"><p:spPr/></p:cxnSp>"#),
        format!(
            r#"<p:cxnSp xmlns:p="{P_NS}"><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr></p:cxnSp>"#
        ),
        format!(
            r#"<p:cxnSp xmlns:p="{P_NS}"><p:spPr/><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr></p:cxnSp>"#
        ),
        format!(
            r#"<p:cxnSp xmlns:p="{P_NS}"><p:nvCxnSpPr><p:cNvCxnSpPr/><p:cNvPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#
        ),
        format!(
            r#"<p:cxnSp xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr><a:endCxn id="4" idx="1"/><a:stCxn id="3" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#
        ),
    ];
    for document in documents {
        assert!(
            CT_ConnectionShape::from_xml(document.as_bytes()).is_err(),
            "{document}"
        );
    }

    for endpoints in [
        "".to_owned(),
        r#"<a:stCxn id="3" idx="0"/>"#.to_owned(),
        r#"<a:endCxn id="4" idx="1"/>"#.to_owned(),
    ] {
        let document = format!(
            r#"<p:cxnSp xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr>{endpoints}</p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#
        );
        CT_ConnectionShape::from_xml(document.as_bytes()).unwrap();
    }
}

#[test]
fn connector_connection_attributes_cannot_be_shadowed_by_qualified_names() {
    let valid = format!(
        r#"<p:cxnSp xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:lookalike"><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr><a:stCxn x:id="99" x:idx="98" id="3" idx="2"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#
    );
    let parsed = CT_ConnectionShape::from_xml(valid.as_bytes()).unwrap();
    assert_eq!(parsed.start_connection.as_ref().unwrap().id, 3);
    assert_eq!(parsed.start_connection.as_ref().unwrap().idx, 2);
    let written = String::from_utf8(parsed.to_xml().unwrap()).unwrap();
    assert!(written.contains(r#"x:id="99" x:idx="98""#));

    let missing_unqualified = format!(
        r#"<p:cxnSp xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:lookalike"><p:nvCxnSpPr><p:cNvPr/><p:cNvCxnSpPr><a:endCxn x:id="97" x:idx="96"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp>"#
    );
    assert!(CT_ConnectionShape::from_xml(missing_unqualified.as_bytes()).is_err());
}

#[test]
fn connector_preserves_unknown_children_in_their_schema_slots() {
    let xml = format!(
        r#"<p:cxnSp xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer" x:root="kept"><x:before/><p:nvCxnSpPr x:nv="kept"><x:nv-before/><p:cNvPr id="2" name="Connector"><x:drawing/></p:cNvPr><x:nv-middle/><p:cNvCxnSpPr x:connector="kept"><x:locks/><a:stCxn id="3" idx="0" x:endpoint="kept"><x:payload/></a:stCxn><x:between/><a:endCxn id="4" idx="1"/><x:extension/></p:cNvCxnSpPr><x:nv-after/><p:nvPr><x:application/></p:nvPr></p:nvCxnSpPr><x:middle/><p:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr><x:after/><p:style><x:style/></p:style><p:extLst><x:extension/></p:extLst></p:cxnSp>"#
    );
    let mut parsed = CT_ConnectionShape::from_xml(xml.as_bytes()).unwrap();
    parsed.start_connection.as_mut().unwrap().id = 30;
    parsed.end_connection.as_mut().unwrap().idx = 10;
    let written = parsed.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    for preserved in [
        r#"<x:before/>"#,
        r#"<x:drawing/>"#,
        r#"<x:locks/>"#,
        r#"<x:payload/>"#,
        r#"<x:between/>"#,
        r#"<x:application/>"#,
        r#"<x:after/>"#,
        r#"<p:style><x:style/></p:style>"#,
        r#"<p:extLst><x:extension/></p:extLst>"#,
    ] {
        assert!(text.contains(preserved), "missing {preserved}: {text}");
    }
    assert!(text.contains(r#"<a:stCxn id="30" idx="0""#));
    assert!(text.contains(r#"<a:endCxn id="4" idx="10""#));
    assert_eq!(parsed, CT_ConnectionShape::from_xml(&written).unwrap());
}

#[derive(Clone, Copy, Default)]
struct ConnectorCoverage {
    total: usize,
    starts: usize,
    ends: usize,
    nested: usize,
}

impl ConnectorCoverage {
    fn add(&mut self, other: Self) {
        self.total += other.total;
        self.starts += other.starts;
        self.ends += other.ends;
        self.nested += other.nested;
    }
}

fn verify_connectors(
    children: &[ShapeTreeChild],
    nested: bool,
    deck: &str,
    part: &str,
) -> ConnectorCoverage {
    let mut coverage = ConnectorCoverage::default();
    for child in children {
        match child {
            ShapeTreeChild::Connector(connector) => {
                let written = connector.to_xml().unwrap();
                assert_eq!(
                    connector,
                    &CT_ConnectionShape::from_xml(&written)
                        .unwrap_or_else(|error| panic!("{deck} {part}: {error}"))
                );
                coverage.total += 1;
                coverage.starts += usize::from(connector.start_connection.is_some());
                coverage.ends += usize::from(connector.end_connection.is_some());
                coverage.nested += usize::from(nested);
            }
            ShapeTreeChild::GroupShape(group) => {
                coverage.add(verify_connectors(&group.children, true, deck, part));
            }
            _ => {}
        }
    }
    coverage
}

#[test]
fn every_corpus_connector_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut coverage = ConnectorCoverage::default();
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            let Some(xml) = package.get_part(part) else {
                continue;
            };
            let children = match content_type.as_str() {
                content_types::SLIDE => {
                    let parsed = CT_Slide::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    parsed.common_slide_data.shape_tree.children
                }
                content_types::SLIDE_LAYOUT => {
                    let parsed = CT_SlideLayout::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    parsed.common_slide_data.shape_tree.children
                }
                content_types::SLIDE_MASTER => {
                    let parsed = CT_SlideMaster::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    parsed.common_slide_data.shape_tree.children
                }
                _ => continue,
            };
            coverage.add(verify_connectors(&children, false, entry.path, part));
        }
    }
    assert!(coverage.total > 0);
    assert!(coverage.starts > 0);
    assert!(coverage.ends > 0);
    assert!(coverage.nested > 0);
    eprintln!(
        "Connector corpus gate checked {} connectors, {} starts, {} ends, and {} nested connectors",
        coverage.total, coverage.starts, coverage.ends, coverage.nested
    );
}

#[test]
fn nested_group_shape_tree_round_trips_with_tree_shape_preserved() {
    let xml = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:grpSp><p:nvGrpSpPr/><p:grpSpPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/><a:chOff x="10" y="20"/><a:chExt cx="30" cy="40"/></a:xfrm></p:grpSpPr><p:sp marker="outer"><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp><p:grpSp><p:nvGrpSpPr/><p:grpSpPr/><p:pic marker="inner"><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic></p:grpSp></p:grpSp></p:spTree></p:cSld></p:sld>"#
    );
    let parsed = CT_Slide::from_xml(xml.as_bytes()).unwrap();
    let ShapeTreeChild::GroupShape(outer) = &parsed.common_slide_data.shape_tree.children[0] else {
        panic!("expected outer group");
    };
    assert_eq!(
        outer.group_transform().unwrap().child_offset.unwrap().x.0,
        10
    );
    assert!(matches!(outer.children[0], ShapeTreeChild::Shape(_)));
    let ShapeTreeChild::GroupShape(inner) = &outer.children[1] else {
        panic!("expected inner group");
    };
    assert!(matches!(inner.children[0], ShapeTreeChild::Picture(_)));
    let written = parsed.to_xml().unwrap();
    let reparsed = CT_Slide::from_xml(&written).unwrap();
    assert_eq!(parsed, reparsed);
    assert_eq!(
        count_groups(&reparsed.common_slide_data.shape_tree.children),
        2
    );
}

#[test]
fn every_corpus_shape_tree_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut trees = 0usize;
    let mut groups = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            let Some(xml) = package.get_part(part) else {
                continue;
            };
            match content_type.as_str() {
                content_types::SLIDE => {
                    let parsed = CT_Slide::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    groups += count_groups(&parsed.common_slide_data.shape_tree.children);
                    assert_eq!(
                        parsed,
                        CT_Slide::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    trees += 1;
                }
                content_types::SLIDE_LAYOUT => {
                    let parsed = CT_SlideLayout::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    groups += count_groups(&parsed.common_slide_data.shape_tree.children);
                    assert_eq!(
                        parsed,
                        CT_SlideLayout::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    trees += 1;
                }
                content_types::SLIDE_MASTER => {
                    let parsed = CT_SlideMaster::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    groups += count_groups(&parsed.common_slide_data.shape_tree.children);
                    assert_eq!(
                        parsed,
                        CT_SlideMaster::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    trees += 1;
                }
                _ => {}
            }
        }
    }
    assert!(trees > 0);
    assert!(groups > 0);
    eprintln!("Shape-tree corpus gate checked {trees} trees and {groups} recursive groups");
}

fn count_groups(children: &[ShapeTreeChild]) -> usize {
    children
        .iter()
        .map(|child| match child {
            ShapeTreeChild::GroupShape(group) => 1 + count_groups(&group.children),
            _ => 0,
        })
        .sum()
}

#[test]
fn cropped_picture_round_trips_with_relationship_and_source_rectangle() {
    let xml = format!(
        r#"<p:pic xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R_NS}"><p:nvPicPr><p:cNvPr id="7" name="Photo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill dpi="144"><a:blip r:embed="rId7"/><a:srcRect l="1000" t="2000" r="3000" b="4000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr></p:pic>"#
    );
    let picture = CT_Picture::from_xml(xml.as_bytes()).unwrap();
    let blip_fill = picture.blip_fill.as_ref().unwrap();
    assert_eq!(
        blip_fill.blip.as_ref().unwrap().embed.as_deref(),
        Some("rId7")
    );
    let crop = blip_fill.source_rect.as_ref().unwrap();
    assert_eq!(crop.left.unwrap().0, 1000);
    assert_eq!(crop.top.unwrap().0, 2000);
    assert_eq!(crop.right.unwrap().0, 3000);
    assert_eq!(crop.bottom.unwrap().0, 4000);
    let written = picture.to_xml().unwrap();
    assert_eq!(picture, CT_Picture::from_xml(&written).unwrap());
}

#[test]
fn picture_reads_any_prefix_and_writes_fixed_prefixes_in_schema_order() {
    let xml = format!(
        r#"<q:pic xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:rel="{R_NS}" xmlns:x="urn:producer"><q:nvPicPr><q:cNvPr id="1" name="Picture"/><q:cNvPicPr/><q:nvPr><q:ph type="pic" idx="3"/></q:nvPr></q:nvPicPr><q:blipFill x:fill-marker="kept"><d:blip rel:embed="rId1"/></q:blipFill><q:spPr><d:solidFill><d:srgbClr val="112233"/></d:solidFill></q:spPr></q:pic>"#
    );
    let picture = CT_Picture::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(picture.placeholder.as_ref().unwrap().idx, Some(3));
    let written = picture.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.starts_with(&format!(r#"<p:pic xmlns:p="{P_NS}""#)));
    assert!(text.contains(&format!(r#"<a:blip xmlns:r="{R_NS}" r:embed="rId1"/>"#)));
    assert!(text.contains(r#"<p:blipFill x:fill-marker="kept">"#));
    assert_order(&written, &["<p:nvPicPr", "<p:blipFill", "<p:spPr"]);
    assert_eq!(picture, CT_Picture::from_xml(&written).unwrap());
}

#[test]
fn picture_requires_non_visual_and_shape_properties_in_schema_order() {
    let invalid = [
        format!(r#"<p:pic xmlns:p="{P_NS}"><p:spPr/></p:pic>"#),
        format!(
            r#"<p:pic xmlns:p="{P_NS}"><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr></p:pic>"#
        ),
        format!(
            r#"<p:pic xmlns:p="{P_NS}"><p:spPr/><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr></p:pic>"#
        ),
        format!(
            r#"<p:pic xmlns:p="{P_NS}"><p:nvPicPr><p:cNvPicPr/><p:cNvPr/><p:nvPr/></p:nvPicPr><p:spPr/></p:pic>"#
        ),
        format!(
            r#"<p:pic xmlns:p="{P_NS}"><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:spPr/><p:blipFill/></p:pic>"#
        ),
        format!(
            r#"<p:pic xmlns:p="{P_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:producer"><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><mc:AlternateContent><mc:Choice Requires="x"><x:not-a-fill/></mc:Choice></mc:AlternateContent><p:spPr/></p:pic>"#
        ),
    ];
    for xml in invalid {
        assert!(CT_Picture::from_xml(xml.as_bytes()).is_err(), "{xml}");
    }
}

#[test]
fn picture_preserves_unknown_children_and_alternate_blip_choice_verbatim() {
    let xml = format!(
        r#"<q:pic xmlns:q="{P_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="urn:p14" xmlns:x="urn:producer" x:marker="kept"><q:nvPicPr><q:cNvPr/><q:cNvPicPr/><q:nvPr><x:before/><q:ph type="pic"/><x:after/></q:nvPr></q:nvPicPr><mc:AlternateContent><mc:Choice Requires="p14"><q:blipFill x:opaque="yes"/></mc:Choice></mc:AlternateContent><q:spPr/><x:after-properties/><q:style><x:style-data/></q:style><q:extLst><x:extension/></q:extLst></q:pic>"#
    );
    let picture = CT_Picture::from_xml(xml.as_bytes()).unwrap();
    assert!(picture.blip_fill.is_none());
    assert!(picture.placeholder.is_some());
    let written = picture.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.contains(r#"x:marker="kept""#));
    assert!(text.contains(r#"<mc:AlternateContent><mc:Choice Requires="p14"><q:blipFill x:opaque="yes"/></mc:Choice></mc:AlternateContent>"#));
    assert_order(
        &written,
        &[
            "<p:nvPicPr",
            "<mc:AlternateContent",
            "<p:spPr",
            "<x:after-properties",
            "<q:style",
            "<q:extLst",
        ],
    );
    assert_eq!(picture, CT_Picture::from_xml(&written).unwrap());
}

#[test]
fn qualified_picture_relationship_attributes_are_not_shadowed() {
    let xml = format!(
        r#"<p:pic xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R_NS}" xmlns:x="urn:not-relationships"><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1" x:embed="shadow"/></p:blipFill><p:spPr/></p:pic>"#
    );
    assert!(CT_Picture::from_xml(xml.as_bytes()).is_err());
}

#[test]
fn every_corpus_picture_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut coverage = PictureCoverage::default();
    let mut raw_pictures = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            let Some(xml) = package.get_part(part) else {
                continue;
            };
            let part_coverage = match content_type.as_str() {
                content_types::SLIDE => {
                    let parsed = CT_Slide::from_xml(xml).unwrap();
                    let result = verify_pictures(
                        &parsed.common_slide_data.shape_tree.children,
                        false,
                        entry.path,
                        part,
                    );
                    assert_eq!(
                        parsed,
                        CT_Slide::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    result
                }
                content_types::SLIDE_LAYOUT => {
                    let parsed = CT_SlideLayout::from_xml(xml).unwrap();
                    let result = verify_pictures(
                        &parsed.common_slide_data.shape_tree.children,
                        false,
                        entry.path,
                        part,
                    );
                    assert_eq!(
                        parsed,
                        CT_SlideLayout::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    result
                }
                content_types::SLIDE_MASTER => {
                    let parsed = CT_SlideMaster::from_xml(xml).unwrap();
                    let result = verify_pictures(
                        &parsed.common_slide_data.shape_tree.children,
                        false,
                        entry.path,
                        part,
                    );
                    assert_eq!(
                        parsed,
                        CT_SlideMaster::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    result
                }
                _ => continue,
            };
            let part_raw_pictures = count_picture_elements(xml);
            assert_eq!(
                part_coverage.typed + part_coverage.opaque,
                part_raw_pictures,
                "{} {part}: every picture must be typed or inside a preserved opaque child",
                entry.path
            );
            raw_pictures += part_raw_pictures;
            coverage.add(part_coverage);
        }
    }
    assert_eq!(raw_pictures, 240);
    assert_eq!(coverage.typed + coverage.opaque, raw_pictures);
    assert_eq!(coverage.cropped + coverage.opaque_cropped, 198);
    assert_eq!(coverage.nested, 98);
    assert_eq!(coverage.placeholders, 18);
    assert_eq!(coverage.alternate_blip, 1);
    assert!(coverage.cropped > 0);
    assert!(coverage.embedded > 0);
    assert!(coverage.linked > 0);
    assert!(coverage.nested > 0);
    assert!(coverage.placeholders > 0);
    assert!(coverage.alternate_blip > 0);
    eprintln!("Picture corpus gate checked {raw_pictures} p:pic elements: {coverage:?}");
}

#[derive(Debug, Default)]
struct PictureCoverage {
    typed: usize,
    opaque: usize,
    opaque_cropped: usize,
    cropped: usize,
    embedded: usize,
    linked: usize,
    nested: usize,
    placeholders: usize,
    alternate_blip: usize,
}

impl PictureCoverage {
    fn add(&mut self, other: Self) {
        self.typed += other.typed;
        self.opaque += other.opaque;
        self.opaque_cropped += other.opaque_cropped;
        self.cropped += other.cropped;
        self.embedded += other.embedded;
        self.linked += other.linked;
        self.nested += other.nested;
        self.placeholders += other.placeholders;
        self.alternate_blip += other.alternate_blip;
    }
}

fn verify_pictures(
    children: &[ShapeTreeChild],
    nested: bool,
    deck: &str,
    part: &str,
) -> PictureCoverage {
    let mut coverage = PictureCoverage::default();
    for child in children {
        match child {
            ShapeTreeChild::Picture(picture) => {
                let written = picture.to_xml().unwrap();
                assert_eq!(
                    picture,
                    &CT_Picture::from_xml(&written)
                        .unwrap_or_else(|error| panic!("{deck} {part}: {error}"))
                );
                coverage.typed += 1;
                coverage.nested += usize::from(nested);
                coverage.placeholders += usize::from(picture.placeholder.is_some());
                coverage.alternate_blip += usize::from(picture.blip_fill.is_none());
                if let Some(fill) = &picture.blip_fill {
                    coverage.cropped += usize::from(fill.source_rect.is_some());
                    if let Some(blip) = &fill.blip {
                        coverage.embedded += usize::from(blip.embed.is_some());
                        coverage.linked += usize::from(blip.link.is_some());
                    }
                }
            }
            ShapeTreeChild::GroupShape(group) => {
                coverage.add(verify_pictures(&group.children, true, deck, part));
            }
            ShapeTreeChild::GraphicFrame(frame) => {
                let xml = frame.to_xml().unwrap();
                coverage.opaque += count_picture_elements(&xml);
                coverage.opaque_cropped += count_pictures_with_descendant(&xml, b"srcRect");
            }
            ShapeTreeChild::Connector(connector) => {
                let xml = connector.to_xml().unwrap();
                coverage.opaque += count_picture_elements(&xml);
                coverage.opaque_cropped += count_pictures_with_descendant(&xml, b"srcRect");
            }
            ShapeTreeChild::AlternateContent(alternate) => {
                coverage.opaque += count_picture_elements(alternate.raw_xml());
                coverage.opaque_cropped +=
                    count_pictures_with_descendant(alternate.raw_xml(), b"srcRect");
            }
            _ => {}
        }
    }
    coverage
}

fn count_picture_elements(xml: &[u8]) -> usize {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut count = 0usize;
    loop {
        match reader.read_event_into(&mut buffer).unwrap() {
            Event::Start(start) | Event::Empty(start)
                if local_name(start.name().as_ref()) == b"pic" =>
            {
                count += 1;
            }
            Event::Eof => return count,
            _ => {}
        }
        buffer.clear();
    }
}

fn count_pictures_with_descendant(xml: &[u8], descendant: &[u8]) -> usize {
    let mut reader = Reader::from_reader(xml);
    let mut buffer = Vec::new();
    let mut picture_matches = Vec::new();
    let mut count = 0usize;
    loop {
        match reader.read_event_into(&mut buffer).unwrap() {
            Event::Start(start) if local_name(start.name().as_ref()) == b"pic" => {
                picture_matches.push(false);
            }
            Event::Start(start) | Event::Empty(start)
                if local_name(start.name().as_ref()) == descendant =>
            {
                if let Some(matches) = picture_matches.last_mut() {
                    *matches = true;
                }
            }
            Event::End(end) if local_name(end.name().as_ref()) == b"pic" => {
                count += usize::from(picture_matches.pop().unwrap());
            }
            Event::Eof => return count,
            _ => {}
        }
        buffer.clear();
    }
}

#[test]
fn placeholders_match_by_idx_before_type() {
    let title = PlaceholderKey {
        ph_type: PhType::Title,
        idx: Some(7),
    };
    let body = PlaceholderKey {
        ph_type: PhType::Body,
        idx: Some(7),
    };
    let same_type_other_idx = PlaceholderKey {
        ph_type: PhType::Title,
        idx: Some(8),
    };
    assert!(title.matches(&body));
    assert!(!title.matches(&same_type_other_idx));
}

#[test]
fn placeholders_match_by_type_when_either_idx_is_absent() {
    let indexed = PlaceholderKey {
        ph_type: PhType::Chart,
        idx: Some(3),
    };
    let unindexed = PlaceholderKey {
        ph_type: PhType::Chart,
        idx: None,
    };
    let other_type = PlaceholderKey {
        ph_type: PhType::Table,
        idx: None,
    };
    assert!(indexed.matches(&unindexed));
    assert!(!indexed.matches(&other_type));
}

#[test]
fn absent_placeholder_type_defaults_to_body() {
    let xml = format!(r#"<q:ph xmlns:q="{P_NS}" idx="4"/>"#);
    let placeholder = CT_Placeholder::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(placeholder.ph_type, None);
    assert_eq!(placeholder.effective_type(), PhType::Body);
    assert_eq!(placeholder.key().idx, Some(4));
}

#[test]
fn title_and_centered_title_are_equivalent_placeholders() {
    let title = PlaceholderKey {
        ph_type: PhType::Title,
        idx: None,
    };
    let centered = PlaceholderKey {
        ph_type: PhType::CenteredTitle,
        idx: None,
    };
    assert!(title.matches(&centered));
    assert!(centered.matches(&title));
}

#[test]
fn body_subtitle_and_object_are_equivalent_placeholders() {
    let keys = [PhType::Body, PhType::Subtitle, PhType::Object]
        .map(|ph_type| PlaceholderKey { ph_type, idx: None });
    for left in &keys {
        for right in &keys {
            assert!(left.matches(right));
        }
    }
}

#[test]
fn placeholder_attributes_and_unknown_children_round_trip_in_place() {
    let xml = format!(
        r#"<q:ph xmlns:q="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer" type="vertBody" orient="vert" sz="half" idx="4294967295" hasCustomPrompt="1" x:marker="kept"><x:before/><a:opaque/><q:extLst><q:ext uri="raw"/></q:extLst></q:ph>"#
    );
    let placeholder = CT_Placeholder::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(placeholder.ph_type, Some(PhType::VerticalBody));
    assert_eq!(placeholder.idx, Some(u32::MAX));
    let written = placeholder.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.starts_with(&format!(r#"<p:ph xmlns:p="{P_NS}""#)));
    assert!(text.contains(r#"orient="vert""#));
    assert!(text.contains(r#"sz="half""#));
    assert!(text.contains(r#"hasCustomPrompt="1""#));
    assert!(text.contains(r#"x:marker="kept""#));
    assert!(text.contains(r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main""#));
    assert!(text.contains("<x:before/><a:opaque/><q:extLst><q:ext uri=\"raw\"/></q:extLst>"));
    assert_eq!(placeholder, CT_Placeholder::from_xml(&written).unwrap());
}

#[test]
fn typed_shape_placeholder_round_trips_inside_nested_groups() {
    let xml = format!(
        r#"<q:spTree xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><q:nvGrpSpPr/><q:grpSpPr/><q:grpSp><q:nvGrpSpPr/><q:grpSpPr/><q:sp marker="shape"><q:nvSpPr x:keep="yes"><q:cNvPr id="2" name="Title"/><x:between/><q:cNvSpPr/><q:nvPr><x:before/><q:ph type="ctrTitle" idx="5"/><x:after/></q:nvPr></q:nvSpPr><q:spPr/><q:txBody><d:bodyPr/><x:raw/><d:p/></q:txBody></q:sp></q:grpSp></q:spTree>"#
    );
    let tree = CT_ShapeTree::from_xml(xml.as_bytes()).unwrap();
    let ShapeTreeChild::GroupShape(group) = &tree.children[0] else {
        panic!("expected group");
    };
    let ShapeTreeChild::Shape(shape) = &group.children[0] else {
        panic!("expected typed shape");
    };
    let placeholder = shape.placeholder.as_ref().unwrap();
    assert_eq!(placeholder.ph_type, Some(PhType::CenteredTitle));
    assert_eq!(placeholder.idx, Some(5));
    let standalone_shape = shape.to_xml().unwrap();
    let standalone_text = std::str::from_utf8(&standalone_shape).unwrap();
    assert!(standalone_text.contains(&format!(r#"xmlns:q="{P_NS}""#)));
    assert!(standalone_text.contains(r#"xmlns:x="urn:producer""#));
    assert_eq!(shape, &CT_Shape::from_xml(&standalone_shape).unwrap());
    let standalone_placeholder = placeholder.to_xml().unwrap();
    assert_eq!(
        placeholder,
        &CT_Placeholder::from_xml(&standalone_placeholder).unwrap()
    );
    let written = tree.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.contains("<p:sp marker=\"shape\""));
    assert!(text.contains("<p:nvSpPr x:keep=\"yes\">"));
    assert_order(&written, &["<x:before/>", "<p:ph", "<x:after/>"]);
    assert_eq!(tree, CT_ShapeTree::from_xml(&written).unwrap());
}

#[test]
fn ordinary_shape_properties_round_trip_in_schema_order() {
    let xml = format!(
        r#"<q:sp xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><q:nvSpPr><q:cNvPr id="7" name="Shape 7"/><q:cNvSpPr/><q:nvPr/></q:nvSpPr><q:spPr x:marker="kept"><x:before/><d:xfrm><d:off x="10" y="20"/><d:ext cx="30" cy="40"/></d:xfrm><x:between/><d:solidFill><d:srgbClr val="112233"/></d:solidFill><x:after/></q:spPr><x:between-shape/><q:style><x:style-payload/><d:lnRef idx="0"/><d:fillRef idx="0"/><d:effectRef idx="0"/><d:fontRef idx="none"/></q:style><q:txBody><d:bodyPr/><d:p/></q:txBody></q:sp>"#
    );
    let shape = CT_Shape::from_xml(xml.as_bytes()).unwrap();
    let properties: &CT_ShapeProperties = &shape.shape_properties;
    let transform = properties.transform.as_ref().unwrap();
    assert_eq!(transform.offset.unwrap().x.0, 10);
    assert!(properties.fill.is_some());

    let written = shape.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.contains(r#"<p:spPr x:marker="kept">"#));
    assert!(text.contains("<x:before/>"));
    assert!(text.contains("<x:between/>"));
    assert!(text.contains("<x:after/>"));
    assert_order(
        &written,
        &[
            "<p:spPr",
            "<x:before/>",
            "<a:xfrm",
            "<x:between/>",
            "<a:solidFill",
            "<x:after/>",
            "<x:between-shape/>",
            "<p:style",
            "<p:txBody",
        ],
    );
    assert_eq!(shape, CT_Shape::from_xml(&written).unwrap());
}

#[test]
fn every_corpus_shape_placeholder_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut placeholders = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            let Some(xml) = package.get_part(part) else {
                continue;
            };
            let children = match content_type.as_str() {
                content_types::SLIDE => {
                    &CT_Slide::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
                        .common_slide_data
                        .shape_tree
                        .children
                }
                content_types::SLIDE_LAYOUT => {
                    &CT_SlideLayout::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
                        .common_slide_data
                        .shape_tree
                        .children
                }
                content_types::SLIDE_MASTER => {
                    &CT_SlideMaster::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
                        .common_slide_data
                        .shape_tree
                        .children
                }
                _ => continue,
            };
            placeholders += verify_shape_placeholders(children, entry.path, part);
        }
    }
    assert!(placeholders > 0);
    eprintln!("Placeholder corpus gate checked {placeholders} typed p:ph elements");
}

fn verify_shape_placeholders(children: &[ShapeTreeChild], deck: &str, part: &str) -> usize {
    children
        .iter()
        .map(|child| match child {
            ShapeTreeChild::Shape(shape) => {
                let written = shape.to_xml().unwrap();
                assert_eq!(
                    shape,
                    &CT_Shape::from_xml(&written)
                        .unwrap_or_else(|error| panic!("{deck} {part}: {error}"))
                );
                usize::from(shape.placeholder.is_some())
            }
            ShapeTreeChild::GroupShape(group) => {
                verify_shape_placeholders(&group.children, deck, part)
            }
            _ => 0,
        })
        .sum()
}

#[test]
fn graphic_data_uri_dispatch_recognises_table_chart_smartart_and_ole() {
    let table =
        r#"<a:tbl><a:tblGrid><a:gridCol w="100"/></a:tblGrid><a:tr h="200"><a:tc/></a:tr></a:tbl>"#;
    let chart = r#"<c:chart xmlns:c="urn:chart" r:id="rId1"/>"#;
    let smartart = r#"<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId2" r:lo="rId3" r:qs="rId4" r:cs="rId5"/>"#;
    let ole = r#"<mc:AlternateContent><mc:Choice Requires="p14"><p:oleObj name="object"/></mc:Choice></mc:AlternateContent>"#;

    let table_frame = CT_GraphicFrame::from_xml(
        graphic_frame_fixture(
            "http://schemas.openxmlformats.org/drawingml/2006/table",
            table,
        )
        .as_bytes(),
    )
    .unwrap();
    let GraphicDataPayload::Table(table) = table_frame.graphic_data.payload() else {
        panic!("table URI did not select the typed table branch");
    };
    assert_eq!(table.grid.columns.len(), 1);
    assert_eq!(table.rows.len(), 1);

    for (uri, payload, expected) in [
        (
            "http://schemas.openxmlformats.org/drawingml/2006/chart",
            chart,
            "chart",
        ),
        (
            "http://schemas.openxmlformats.org/drawingml/2006/diagram",
            smartart,
            "smartart",
        ),
        (
            "http://schemas.openxmlformats.org/presentationml/2006/ole",
            ole,
            "ole",
        ),
        ("urn:producer:graphic-data", "<x:payload/>", "other"),
    ] {
        let frame =
            CT_GraphicFrame::from_xml(graphic_frame_fixture(uri, payload).as_bytes()).unwrap();
        let raw = match (frame.graphic_data.payload(), expected) {
            (GraphicDataPayload::Chart(raw), "chart")
            | (GraphicDataPayload::Other(raw), "other") => raw.clone(),
            (GraphicDataPayload::SmartArt(relationships), "smartart") => relationships.to_xml(),
            (GraphicDataPayload::Ole { raw, .. }, "ole") => raw.clone(),
            (actual, _) => panic!("{uri} selected the wrong branch: {actual:?}"),
        };
        assert_eq!(raw, payload.as_bytes());
        assert_eq!(
            frame,
            CT_GraphicFrame::from_xml(&frame.to_xml().unwrap()).unwrap()
        );
    }
}

#[test]
fn shape_ids_ignore_foreign_non_visual_elements() {
    let xml = br#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:extension"><p:nvGrpSpPr><p:cNvPr id="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:grpSp><p:nvGrpSpPr><x:cNvPr id="99"/><p:cNvPr id="4"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><x:cNvPr id="98"/><p:cNvPr id="6"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></p:xfrm><a:graphic><a:graphicData uri="urn:producer"><x:payload/></a:graphicData></a:graphic></p:graphicFrame></p:grpSp></p:spTree>"#;
    let tree = CT_ShapeTree::from_xml(xml).unwrap();
    let ShapeTreeChild::GroupShape(group) = &tree.children[0] else {
        panic!("expected group shape");
    };

    assert_eq!(tree.children[0].non_visual_id(), Some(4));
    assert_eq!(group.children[0].non_visual_id(), Some(6));
    let mut allocator = ShapeIdAllocator::scan(&tree);
    assert_eq!(allocator.allocate(), 2);
    assert_eq!(allocator.allocate(), 3);
    assert_eq!(allocator.allocate(), 5);
}

#[test]
fn ole_payload_projects_optional_fallback_picture_without_rewriting_raw_bytes() {
    let payload = r#"<mc:AlternateContent><mc:Choice Requires="p"><p:oleObj name="object"><p:embed/></p:oleObj></mc:Choice><mc:Fallback><p:oleObj name="object"><p:embed/><p:pic><p:nvPicPr><p:cNvPr/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="11" y="22"/><a:ext cx="33" cy="44"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></mc:Fallback></mc:AlternateContent>"#;
    let frame = CT_GraphicFrame::from_xml(
        graphic_frame_fixture(
            "http://schemas.openxmlformats.org/presentationml/2006/ole",
            payload,
        )
        .as_bytes(),
    )
    .unwrap();
    let GraphicDataPayload::Ole { raw, preview } = frame.graphic_data.payload() else {
        panic!("OLE URI did not select the OLE branch");
    };

    assert_eq!(raw, payload.as_bytes());
    let preview = preview
        .as_ref()
        .expect("fallback picture was not projected");
    assert_eq!(
        preview
            .blip_fill
            .as_ref()
            .and_then(|fill| fill.blip.as_ref())
            .and_then(|blip| blip.embed.as_deref()),
        Some("rId7")
    );
    let transform = preview
        .shape_properties
        .transform
        .as_ref()
        .expect("fallback picture transform was not projected");
    assert_eq!(transform.offset.as_ref().map(|offset| offset.x.0), Some(11));
    assert_eq!(
        transform.extent.as_ref().map(|extents| extents.cx.0),
        Some(33)
    );
    assert_eq!(
        frame,
        CT_GraphicFrame::from_xml(&frame.to_xml().unwrap()).unwrap()
    );

    let absent = CT_GraphicFrame::from_xml(
        graphic_frame_fixture(
            "http://schemas.openxmlformats.org/presentationml/2006/ole",
            r#"<p:oleObj name="object"><p:embed/></p:oleObj>"#,
        )
        .as_bytes(),
    )
    .unwrap();
    assert!(matches!(
        absent.graphic_data.payload(),
        GraphicDataPayload::Ole { preview: None, .. }
    ));
}

#[test]
fn graphic_frame_requires_children_in_schema_order_and_writes_fixed_prefixes() {
    let xml = format!(
        r#"<q:graphicFrame xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="urn:chart"><q:nvGraphicFramePr><q:cNvPr id="2" name="Chart"/><q:cNvGraphicFramePr/><q:nvPr/></q:nvGraphicFramePr><q:xfrm rot="60000"><d:off x="10" y="20"/><d:ext cx="30" cy="40"/></q:xfrm><d:graphic><d:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></d:graphicData></d:graphic></q:graphicFrame>"#
    );
    let frame = CT_GraphicFrame::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(frame.transform.offset.unwrap().x.0, 10);
    let written = frame.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    assert!(text.starts_with("<p:graphicFrame xmlns:p="));
    assert_order(
        &written,
        &[
            "<p:nvGraphicFramePr",
            "<p:xfrm",
            "<a:off",
            "<a:graphic",
            "<a:graphicData",
        ],
    );
    assert_eq!(frame, CT_GraphicFrame::from_xml(&written).unwrap());

    for invalid in [
        format!(
            r#"<p:graphicFrame xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:xfrm/><a:graphic><a:graphicData uri="urn:x"><a:x/></a:graphicData></a:graphic></p:graphicFrame>"#
        ),
        format!(
            r#"<p:graphicFrame xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvGraphicFramePr/><a:graphic><a:graphicData uri="urn:x"><a:x/></a:graphicData></a:graphic><p:xfrm/></p:graphicFrame>"#
        ),
        format!(
            r#"<p:graphicFrame xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvGraphicFramePr/><p:xfrm/></p:graphicFrame>"#
        ),
        format!(
            r#"<p:graphicFrame xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvGraphicFramePr/><p:xfrm/><p:xfrm/><a:graphic><a:graphicData uri="urn:x"><a:x/></a:graphicData></a:graphic></p:graphicFrame>"#
        ),
    ] {
        assert!(
            CT_GraphicFrame::from_xml(invalid.as_bytes()).is_err(),
            "{invalid}"
        );
    }
}

#[test]
fn table_graphic_frame_constructor_writes_the_canonical_shell() {
    let table = CT_Table::new(2, 3, Emu(302), Emu(201)).unwrap();
    let frame =
        CT_GraphicFrame::new_table(7, "Table & 7", CT_Transform2D::default(), table).unwrap();
    let written = frame.to_xml().unwrap();
    let xml = String::from_utf8(written.clone()).unwrap();
    assert!(xml.contains(r#"<p:cNvPr id="7" name="Table &amp; 7"/>"#));
    assert!(xml.contains("<p:cNvGraphicFramePr/>"));
    assert!(xml.contains("<p:nvPr/>"));
    assert!(xml.contains(
        r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">"#
    ));
    assert!(xml.find("<p:nvGraphicFramePr>").unwrap() < xml.find("<p:xfrm").unwrap());
    assert!(xml.find("<p:xfrm").unwrap() < xml.find("<a:graphic>").unwrap());
    let reparsed = CT_GraphicFrame::from_xml(&written).unwrap();
    assert_eq!(reparsed, frame);
    assert!(matches!(
        reparsed.graphic_data.payload(),
        GraphicDataPayload::Table(_)
    ));

    let mut invalid_table = CT_Table::new(1, 1, Emu(100), Emu(100)).unwrap();
    invalid_table.rows.clear();
    let error =
        CT_GraphicFrame::new_table(8, "Invalid table", CT_Transform2D::default(), invalid_table)
            .unwrap_err();
    assert!(error.to_string().contains("at least one a:tr"));
}

#[test]
fn chart_relationship_projection_ignores_unqualified_and_foreign_ids() {
    for payload in [
        r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" id="plain"/>"#,
        r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:x="urn:foreign" x:id="foreign"/>"#,
        r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id=""/>"#,
    ] {
        let frame = CT_GraphicFrame::from_xml(
            graphic_frame_fixture(
                "http://schemas.openxmlformats.org/drawingml/2006/chart",
                payload,
            )
            .as_bytes(),
        )
        .unwrap();
        assert_eq!(frame.chart_relationship_id(), None);
    }
}

#[test]
fn authored_chart_graphic_frame_round_trips() {
    let mut frame =
        CT_GraphicFrame::new_chart(9, "Chart & 9", CT_Transform2D::default(), "rId7").unwrap();
    assert!(frame.graphic_data.table_mut().is_none());
    assert_eq!(frame.chart_relationship_id(), Some("rId7"));
    let written = frame.to_xml().unwrap();
    let xml = String::from_utf8(written.clone()).unwrap();
    assert!(xml.contains(r#"<p:cNvPr id="9" name="Chart &amp; 9"/>"#));
    assert!(xml.contains(
        r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">"#
    ));
    assert!(xml.contains(
        r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId7"/>"#
    ));
    assert_order(
        &written,
        &[
            "<p:nvGraphicFramePr",
            "<p:xfrm",
            "<a:graphic",
            "<a:graphicData",
            "<c:chart",
        ],
    );
    let reparsed = CT_GraphicFrame::from_xml(&written).unwrap();
    assert_eq!(reparsed.chart_relationship_id(), Some("rId7"));
    assert_eq!(reparsed, frame);
    assert!(matches!(
        reparsed.graphic_data.payload(),
        GraphicDataPayload::Chart(_)
    ));
}

#[test]
fn chart_choice_and_picture_fallback_remain_byte_preserved() {
    let raw = format!(
        r#"<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:z="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:rel="{R_NS}" marker="kept"><mc:Choice Requires="z"><q:graphicFrame><q:nvGraphicFramePr/><q:xfrm><d:off x="1" y="2"/><d:ext cx="127000" cy="254000"/></q:xfrm><d:graphic><d:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><z:chart rel:id="chartAlias"/></d:graphicData></d:graphic></q:graphicFrame></mc:Choice><mc:Fallback><q:pic><q:nvPicPr><q:cNvPr/><q:cNvPicPr/><q:nvPr/></q:nvPicPr><q:blipFill><d:blip rel:embed="previewAlias"/><d:stretch><d:fillRect/></d:stretch></q:blipFill><q:spPr/></q:pic></mc:Fallback></mc:AlternateContent>"#
    );
    let tree_xml = format!(
        r#"<q:spTree xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="{MC_NS}" xmlns:z="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:rel="{R_NS}"><q:nvGrpSpPr/><q:grpSpPr/>{raw}</q:spTree>"#
    );
    let tree = rpptx_oxml::shape_tree::CT_ShapeTree::from_xml(tree_xml.as_bytes()).unwrap();
    let ShapeTreeChild::AlternateContent(alternate) = &tree.children[0] else {
        panic!("expected chart alternate content");
    };

    assert_eq!(
        alternate
            .chart_choice()
            .and_then(CT_GraphicFrame::chart_relationship_id),
        Some("chartAlias")
    );
    assert!(alternate.picture_fallback().is_some());
    assert_eq!(alternate.raw_xml(), raw.as_bytes());

    let written = tree.to_xml().unwrap();
    assert!(
        written
            .windows(raw.len())
            .any(|window| window == raw.as_bytes()),
        "the preserved alternate-content subtree must remain byte-identical"
    );
}

#[test]
fn malformed_non_chart_choice_remains_opaque() {
    let raw = format!(
        r#"<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:p="{P_NS}"><mc:Choice Requires="x"><p:graphicFrame><p:xfrm/></p:graphicFrame></mc:Choice><mc:Fallback/></mc:AlternateContent>"#
    );
    let tree_xml = format!(
        r#"<p:spTree xmlns:p="{P_NS}" xmlns:mc="{MC_NS}"><p:nvGrpSpPr/><p:grpSpPr/>{raw}</p:spTree>"#
    );

    let tree = rpptx_oxml::shape_tree::CT_ShapeTree::from_xml(tree_xml.as_bytes()).unwrap();
    let ShapeTreeChild::AlternateContent(alternate) = &tree.children[0] else {
        panic!("expected preserved alternate content");
    };
    assert!(alternate.chart_choice().is_none());
    assert_eq!(alternate.raw_xml(), raw.as_bytes());
}

#[test]
fn non_chart_choice_with_descendant_chart_uri_remains_opaque() {
    let raw = format!(
        r#"<mc:AlternateContent xmlns:mc="{MC_NS}" xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><mc:Choice Requires="x"><p:graphicFrame><p:nvGraphicFramePr/><p:xfrm/><a:graphic><a:graphicData uri="urn:producer"><x:extension><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></x:extension></a:graphicData></a:graphic></p:graphicFrame></mc:Choice><mc:Fallback/></mc:AlternateContent>"#
    );
    let tree_xml = format!(
        r#"<p:spTree xmlns:p="{P_NS}" xmlns:mc="{MC_NS}"><p:nvGrpSpPr/><p:grpSpPr/>{raw}</p:spTree>"#
    );

    let tree = rpptx_oxml::shape_tree::CT_ShapeTree::from_xml(tree_xml.as_bytes()).unwrap();
    let ShapeTreeChild::AlternateContent(alternate) = &tree.children[0] else {
        panic!("expected preserved alternate content");
    };
    assert!(alternate.chart_choice().is_none());
    assert_eq!(alternate.raw_xml(), raw.as_bytes());
}

#[test]
fn graphic_frame_preserves_unknown_payload_and_extension_xml_byte_for_byte() {
    let payload = r#"<u:payload xmlns:u="urn:unknown" marker="a &amp; b"><u:data><!--kept--></u:data></u:payload>"#;
    let xml = format!(
        r#"<p:graphicFrame xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer" x:frame="kept"><x:before/><!--frame-kept--><p:nvGraphicFramePr x:nv="kept"><!--nv-kept--><x:nv-child/><?nv kept?></p:nvGraphicFramePr><x:between-nv-and-xfrm/><p:xfrm/><x:between-xfrm-and-graphic/><a:graphic x:graphic="kept"><x:graphic-before/><!--graphic-kept--><a:graphicData uri="urn:producer:data" x:data="kept"><!--data-before-->{payload}<?data kept?><x:payload-after/></a:graphicData><x:graphic-after/></a:graphic><x:before-ext/><p:extLst><p:ext uri="kept"><x:extension/></p:ext></p:extLst><x:after/></p:graphicFrame>"#
    );
    let frame = CT_GraphicFrame::from_xml(xml.as_bytes()).unwrap();
    let GraphicDataPayload::Other(actual) = frame.graphic_data.payload() else {
        panic!("unknown URI did not select the opaque branch");
    };
    assert_eq!(actual, payload.as_bytes());
    let written = frame.to_xml().unwrap();
    let text = std::str::from_utf8(&written).unwrap();
    for raw in [
        "<x:before/>",
        "<!--frame-kept-->",
        "<!--nv-kept-->",
        "<x:nv-child/>",
        "<?nv kept?>",
        "<x:between-nv-and-xfrm/>",
        "<x:between-xfrm-and-graphic/>",
        "<x:graphic-before/>",
        "<!--graphic-kept-->",
        "<!--data-before-->",
        payload,
        "<?data kept?>",
        "<x:payload-after/>",
        "<x:graphic-after/>",
        "<x:before-ext/>",
        "<p:extLst><p:ext uri=\"kept\"><x:extension/></p:ext></p:extLst>",
        "<x:after/>",
    ] {
        assert!(text.contains(raw), "missing {raw}: {text}");
    }
    assert_order(
        &written,
        &[
            "<x:before/>",
            "<p:nvGraphicFramePr",
            "<x:between-nv-and-xfrm/>",
            "<p:xfrm",
            "<x:between-xfrm-and-graphic/>",
            "<a:graphic",
            "<x:before-ext/>",
            "<p:extLst",
            "<x:after/>",
        ],
    );
    assert_eq!(frame, CT_GraphicFrame::from_xml(&written).unwrap());
}

#[test]
fn empty_namespace_shadow_does_not_hide_later_inherited_binding() {
    let xml = format!(
        r#"<p:spTree xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:outer"><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr/><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr><x:shadow xmlns:x="urn:inner"/></a:tblPr><a:tblGrid><a:gridCol w="100"/></a:tblGrid><a:tr h="200"><a:tc><x:needed/><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree>"#
    );
    let tree = CT_ShapeTree::from_xml(xml.as_bytes()).unwrap();
    let ShapeTreeChild::GraphicFrame(frame) = &tree.children[0] else {
        panic!("shape tree did not retain the graphic frame");
    };
    let frame_xml = frame.to_xml().unwrap();
    let frame_text = std::str::from_utf8(&frame_xml).unwrap();
    assert!(frame_text.contains(r#"xmlns:x="urn:outer""#));
    assert!(frame_text.contains(r#"<x:shadow xmlns:x="urn:inner"/>"#));
    assert!(frame_text.contains("<x:needed/>"));

    let GraphicDataPayload::Table(table) = frame.graphic_data.payload() else {
        panic!("table URI did not select the typed table branch");
    };
    let table_xml = table.to_xml().unwrap();
    let table_text = std::str::from_utf8(&table_xml).unwrap();
    assert!(table_text.contains(r#"xmlns:x="urn:outer""#));
    assert!(table_text.contains("<x:needed/>"));
}

#[test]
fn every_corpus_graphic_frame_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut coverage = GraphicFrameCoverage::default();
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            let Some(xml) = package.get_part(part) else {
                continue;
            };
            let children = match content_type.as_str() {
                content_types::SLIDE => {
                    &CT_Slide::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
                        .common_slide_data
                        .shape_tree
                        .children
                }
                content_types::SLIDE_LAYOUT => {
                    &CT_SlideLayout::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
                        .common_slide_data
                        .shape_tree
                        .children
                }
                content_types::SLIDE_MASTER => {
                    &CT_SlideMaster::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path))
                        .common_slide_data
                        .shape_tree
                        .children
                }
                _ => continue,
            };
            verify_graphic_frames(children, entry.path, part, &mut coverage);
        }
    }
    assert_eq!(coverage.tables, 26);
    assert_eq!(coverage.charts, 26);
    assert_eq!(coverage.smartart, 18);
    assert_eq!(coverage.ole, 16);
    assert_eq!(coverage.other, 0);
    assert_eq!(coverage.total(), 86);
    eprintln!("Graphic-frame corpus gate checked {coverage:?}");
}

#[test]
fn text_body_plain_text_preserves_run_field_break_and_paragraph_order() {
    let body = CT_TextBody::from_xml(
        br#"<a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>first</a:t></a:r><a:br/><a:fld id="{14E3C6B4-82AF-4BA3-9737-6C8DF2E2A225}" type="datetime"><a:t>field</a:t></a:fld><a:r><a:t> last</a:t></a:r></a:p><a:p><a:r><a:t>second</a:t></a:r></a:p></a:txBody>"#,
    )
    .unwrap();
    assert_eq!(body.plain_text(), "first\nfield last\nsecond");
}

#[test]
fn notes_slide_extracts_only_body_placeholder_text() {
    let xml = format!(
        r#"<p:notes xmlns:p="{P_NS}" xmlns:a="{}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>slide image</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{{14E3C6B4-82AF-4BA3-9737-6C8DF2E2A225}}"><a:t>12</a:t></a:fld></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr><p:ph type="ftr"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>footer</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr/><p:cNvSpPr/><p:nvPr><p:ph/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>speaker</a:t></a:r><a:br/><a:r><a:t>note</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:notes>"#,
        oxml_drawing::namespace::A_NS
    );
    let notes = CT_NotesSlide::from_xml(xml.as_bytes()).unwrap();
    assert_eq!(notes.notes_text(), "speaker\nnote");
    assert_eq!(
        notes,
        CT_NotesSlide::from_xml(&notes.to_xml().unwrap()).unwrap()
    );
}

#[test]
fn notes_parts_read_any_prefix_write_fixed_prefixes_and_schema_order() {
    let color_map = r#"bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink""#;
    let slide_xml = format!(
        r#"<q:notes xmlns:q="{P_NS}" xmlns:d="{}"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMapOvr><d:masterClrMapping/></q:clrMapOvr><q:extLst><q:ext uri="slide"/></q:extLst></q:notes>"#,
        oxml_drawing::namespace::A_NS
    );
    let slide = CT_NotesSlide::from_xml(slide_xml.as_bytes()).unwrap();
    let slide_written = slide.to_xml().unwrap();
    assert_order(&slide_written, &["<p:cSld", "<p:clrMapOvr", "<q:extLst"]);
    assert_eq!(slide, CT_NotesSlide::from_xml(&slide_written).unwrap());

    let master_xml = format!(
        r#"<q:notesMaster xmlns:q="{P_NS}" xmlns:d="{}"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap {color_map}/><q:hf hdr="1"/><q:notesStyle><d:lvl1pPr/></q:notesStyle><q:extLst><q:ext uri="master"/></q:extLst></q:notesMaster>"#,
        oxml_drawing::namespace::A_NS
    );
    let master = CT_NotesMaster::from_xml(master_xml.as_bytes()).unwrap();
    let master_written = master.to_xml().unwrap();
    assert_order(
        &master_written,
        &[
            "<p:cSld",
            "<p:clrMap",
            "<p:hf",
            "<p:notesStyle",
            "<q:extLst",
        ],
    );
    assert_eq!(master, CT_NotesMaster::from_xml(&master_written).unwrap());

    let out_of_order = format!(
        r#"<p:notesMaster xmlns:p="{P_NS}"><p:clrMap {color_map}/><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:notesStyle/></p:notesMaster>"#
    );
    assert!(CT_NotesMaster::from_xml(out_of_order.as_bytes()).is_err());
}

#[test]
fn notes_master_without_notes_style_round_trips_with_extension_in_schema_order() {
    let color_map = r#"bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink""#;
    for header_footer in ["", r#"<q:hf hdr="1"/>"#] {
        let xml = format!(
            r#"<q:notesMaster xmlns:q="{P_NS}" xmlns:x="urn:producer"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap {color_map}/>{header_footer}<x:between/><q:extLst><x:payload/></q:extLst></q:notesMaster>"#
        );
        let parsed = CT_NotesMaster::from_xml(xml.as_bytes()).unwrap();
        assert!(parsed.notes_style.is_none());
        let written = parsed.to_xml().unwrap();
        let mut expected = vec!["<p:cSld", "<p:clrMap"];
        if !header_footer.is_empty() {
            expected.push("<p:hf");
        }
        expected.extend(["<x:between", "<q:extLst"]);
        assert_order(&written, &expected);
        assert_eq!(
            parsed,
            CT_NotesMaster::from_xml(&written).unwrap(),
            "{header_footer}"
        );
    }
}

#[test]
fn notes_parts_preserve_unmodelled_children_in_their_schema_slots() {
    let producer_before =
        r#"<x:before marker="a &amp; b"><x:data/><!--kept--><?producer value?></x:before>"#;
    let producer_between = r#"<x:between value="2"/>"#;
    let producer_after = r#"<x:after value="3"/>"#;
    let xml = format!(
        r#"<q:notes xmlns:q="{P_NS}" xmlns:d="{}" xmlns:x="urn:producer" x:root="kept">{producer_before}<q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld>{producer_between}<q:clrMapOvr><d:masterClrMapping/></q:clrMapOvr>{producer_after}<q:extLst><x:payload/></q:extLst></q:notes>"#,
        oxml_drawing::namespace::A_NS
    );
    let parsed = CT_NotesSlide::from_xml(xml.as_bytes()).unwrap();
    let written = parsed.to_xml().unwrap();
    let written_text = std::str::from_utf8(&written).unwrap();
    assert!(written_text.contains(r#"xmlns:x="urn:producer""#));
    assert!(written_text.contains(r#"x:root="kept""#));
    for raw in [producer_before, producer_between, producer_after] {
        assert!(
            written
                .windows(raw.len())
                .any(|window| window == raw.as_bytes())
        );
    }
    assert_order(
        &written,
        &[
            producer_before,
            "<p:cSld",
            producer_between,
            "<p:clrMapOvr",
            producer_after,
        ],
    );
}

#[test]
fn notes_handout_and_header_footer_preserve_direct_raw_events_at_schema_boundaries() {
    use rpptx_oxml::notes_parts::CT_HandoutMaster;

    let color_map = r#"bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink""#;
    let notes_before =
        "notes-text<![CDATA[notes-cdata]]><!--notes-comment--><?notes keep?><!DOCTYPE notes-event>";
    let notes_between = "between-text<![CDATA[between-cdata]]><!--between-comment--><?between keep?><!DOCTYPE between-event>";
    let header_footer_events =
        "hf-text<![CDATA[hf-cdata]]><!--hf-comment--><?hf keep?><!DOCTYPE hf-event>";
    let notes_after =
        "after-text<![CDATA[after-cdata]]><!--after-comment--><?after keep?><!DOCTYPE after-event>";
    let notes_xml = format!(
        r#"<q:notesMaster xmlns:q="{P_NS}" xmlns:x="urn:f217">{notes_before}<q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap {color_map}/>{notes_between}<q:hf hdr="1">{header_footer_events}<x:hf-child/></q:hf>{notes_after}</q:notesMaster>"#
    );
    let notes = CT_NotesMaster::from_xml(notes_xml.as_bytes()).unwrap();
    let notes_written = notes.to_xml().unwrap();
    for raw in [
        notes_before,
        notes_between,
        header_footer_events,
        notes_after,
    ] {
        assert!(
            notes_written
                .windows(raw.len())
                .any(|window| window == raw.as_bytes()),
            "missing {raw}: {}",
            String::from_utf8_lossy(&notes_written)
        );
    }
    assert_order(
        &notes_written,
        &[
            notes_before,
            "<p:cSld",
            "<p:clrMap",
            notes_between,
            "<p:hf",
            header_footer_events,
            "<x:hf-child",
            notes_after,
        ],
    );
    assert_eq!(notes, CT_NotesMaster::from_xml(&notes_written).unwrap());

    let handout_events = "handout-text<![CDATA[handout-cdata]]><!--handout-comment--><?handout keep?><!DOCTYPE handout-event>";
    let handout_xml = format!(
        r#"<q:handoutMaster xmlns:q="{P_NS}"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld>{handout_events}<q:clrMap {color_map}/><q:hf ftr="0"/></q:handoutMaster>"#
    );
    let handout = CT_HandoutMaster::from_xml(handout_xml.as_bytes()).unwrap();
    let handout_written = handout.to_xml().unwrap();
    assert!(
        handout_written
            .windows(handout_events.len())
            .any(|window| window == handout_events.as_bytes()),
        "{}",
        String::from_utf8_lossy(&handout_written)
    );
    assert_order(
        &handout_written,
        &["<p:cSld", handout_events, "<p:clrMap", "<p:hf"],
    );
    assert_eq!(
        handout,
        CT_HandoutMaster::from_xml(&handout_written).unwrap()
    );
}

#[test]
fn every_corpus_notes_slide_and_master_round_trips_structurally() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut notes_slides = 0usize;
    let mut notes_masters = 0usize;
    let mut nonempty_notes = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            let Some(xml) = package.get_part(part) else {
                continue;
            };
            match content_type.as_str() {
                content_types::NOTES_SLIDE => {
                    let parsed = CT_NotesSlide::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    nonempty_notes += usize::from(!parsed.notes_text().is_empty());
                    assert_eq!(
                        parsed,
                        CT_NotesSlide::from_xml(&parsed.to_xml().unwrap()).unwrap(),
                        "{} {part}",
                        entry.path
                    );
                    notes_slides += 1;
                }
                content_types::NOTES_MASTER => {
                    let parsed = CT_NotesMaster::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    assert_eq!(
                        parsed,
                        CT_NotesMaster::from_xml(&parsed.to_xml().unwrap()).unwrap(),
                        "{} {part}",
                        entry.path
                    );
                    notes_masters += 1;
                }
                _ => {}
            }
        }
    }
    assert!(notes_slides > 0);
    assert!(notes_masters > 0);
    assert!(nonempty_notes > 0);
    eprintln!(
        "Notes corpus gate checked {notes_slides} notes slides, {notes_masters} notes masters, and {nonempty_notes} nonempty note bodies"
    );
}

#[test]
fn corpus_notes_relationships_are_complete() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut notes_slides = 0usize;
    let mut notes_masters = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            if !package.parts.contains_key(part) {
                continue;
            }
            let relationships = package.get_part_rels(part);
            let count = |relationship_type: &str| {
                relationships.map_or(0, |items| {
                    items
                        .items
                        .iter()
                        .filter(|relationship| relationship.rel_type == relationship_type)
                        .count()
                })
            };
            match content_type.as_str() {
                content_types::NOTES_SLIDE => {
                    assert_eq!(count(rel_types::NOTES_MASTER), 1, "{} {part}", entry.path);
                    assert_eq!(count(rel_types::SLIDE), 1, "{} {part}", entry.path);
                    notes_slides += 1;
                }
                content_types::NOTES_MASTER => {
                    assert_eq!(count(rel_types::THEME), 1, "{} {part}", entry.path);
                    notes_masters += 1;
                }
                content_types::SLIDE => {
                    assert!(count(rel_types::NOTES_SLIDE) <= 1, "{} {part}", entry.path);
                }
                _ => {}
            }
        }
    }
    assert!(notes_slides > 0);
    assert!(notes_masters > 0);
}

const TABLE_GRAPHIC_DATA_URI: &str = "http://schemas.openxmlformats.org/drawingml/2006/table";

fn graphic_frame_fixture(uri: &str, payload: &str) -> String {
    format!(
        r#"<p:graphicFrame xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:producer"><p:nvGraphicFramePr/><p:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></p:xfrm><a:graphic><a:graphicData uri="{uri}">{payload}</a:graphicData></a:graphic></p:graphicFrame>"#
    )
}

#[derive(Debug, Default)]
struct GraphicFrameCoverage {
    tables: usize,
    charts: usize,
    smartart: usize,
    ole: usize,
    other: usize,
}

impl GraphicFrameCoverage {
    const fn total(&self) -> usize {
        self.tables + self.charts + self.smartart + self.ole + self.other
    }
}

fn verify_graphic_frames(
    children: &[ShapeTreeChild],
    deck: &str,
    part: &str,
    coverage: &mut GraphicFrameCoverage,
) {
    for child in children {
        match child {
            ShapeTreeChild::GraphicFrame(frame) => {
                let written = frame.to_xml().unwrap();
                assert_eq!(
                    frame.as_ref(),
                    &CT_GraphicFrame::from_xml(&written)
                        .unwrap_or_else(|error| panic!("{deck} {part}: {error}")),
                    "{deck} {part}: graphic frame changed"
                );
                match frame.graphic_data.payload() {
                    GraphicDataPayload::Table(table) => {
                        assert_eq!(frame.graphic_data.uri, TABLE_GRAPHIC_DATA_URI);
                        assert_eq!(
                            table.as_ref(),
                            &CT_Table::from_xml(&table.to_xml().unwrap()).unwrap()
                        );
                        coverage.tables += 1;
                    }
                    GraphicDataPayload::Chart(_) => coverage.charts += 1,
                    GraphicDataPayload::SmartArt(_) => coverage.smartart += 1,
                    GraphicDataPayload::Ole { .. } => coverage.ole += 1,
                    GraphicDataPayload::Other(_) => coverage.other += 1,
                }
            }
            ShapeTreeChild::GroupShape(group) => {
                verify_graphic_frames(&group.children, deck, part, coverage);
            }
            _ => {}
        }
    }
}

#[test]
fn rewrites_relationship_attributes_and_no_other_bytes() {
    let raw = format!(
        r#"<x:payload xmlns:x="urn:producer" xmlns:r="{R_NS}" r:embed="rId1" x:keep="rId1"><x:item r:link='rId2'/><x:item r:dm="rId3"/></x:payload>"#
    );
    let map = HashMap::from([
        ("rId1".to_owned(), "rId11".to_owned()),
        ("rId2".to_owned(), "rId12".to_owned()),
        ("rId3".to_owned(), "rId13".to_owned()),
    ]);

    let rewritten = rewrite_rel_ids(raw.as_bytes(), &map).unwrap();

    assert_eq!(
        rewritten,
        format!(
            r#"<x:payload xmlns:x="urn:producer" xmlns:r="{R_NS}" r:embed="rId11" x:keep="rId1"><x:item r:link='rId12'/><x:item r:dm="rId13"/></x:payload>"#
        )
        .into_bytes()
    );
}

#[test]
fn relationship_namespace_aliases_and_shadowing_are_respected() {
    let raw = format!(
        r#"<p:root xmlns:p="urn:producer" xmlns:rel="{R_NS}" xmlns:r="{R_NS}" rel:embed="rId1"><p:item xmlns:r="urn:shadow" r:link="rId2" rel:dm="rId3"/><p:item r:id="rId4" r:lo="rId&#52;"/></p:root>"#
    );
    let map = HashMap::from([
        ("rId1".to_owned(), "rId11".to_owned()),
        ("rId2".to_owned(), "rId12".to_owned()),
        ("rId3".to_owned(), "rId13".to_owned()),
        ("rId4".to_owned(), "rId14".to_owned()),
    ]);

    let rewritten = rewrite_rel_ids(raw.as_bytes(), &map).unwrap();

    assert_eq!(
        rewritten,
        format!(
            r#"<p:root xmlns:p="urn:producer" xmlns:rel="{R_NS}" xmlns:r="{R_NS}" rel:embed="rId11"><p:item xmlns:r="urn:shadow" r:link="rId2" rel:dm="rId13"/><p:item r:id="rId14" r:lo="rId14"/></p:root>"#
        )
        .into_bytes()
    );
}

#[test]
fn unmapped_and_non_numeric_relationship_values_are_unchanged() {
    let raw = format!(
        r#"<x:payload xmlns:x="urn:producer" xmlns:r="{R_NS}" r:id="rId1" r:embed="rIdABC" r:link="rId" r:dm="rId12x" id="rId2" x:id="rId3"/>"#
    );
    let map = HashMap::from([
        ("rId2".to_owned(), "rId20".to_owned()),
        ("rId3".to_owned(), "rId30".to_owned()),
        ("rIdABC".to_owned(), "rId40".to_owned()),
    ]);

    assert_eq!(
        rewrite_rel_ids(raw.as_bytes(), &map).unwrap(),
        raw.as_bytes()
    );
}

#[test]
fn exact_relationship_rewrite_changes_only_named_nonnumeric_ids() {
    let raw = format!(
        r#"<x:payload xmlns:x="urn:producer" xmlns:r="{R_NS}" r:id="slide-link" syntax=" A &amp; B "><x:other r:id='other-link'/><x:foreign x:id="slide-link"/></x:payload>"#
    );
    let map = HashMap::from([("slide-link".to_owned(), "rId9".to_owned())]);

    let rewritten = rewrite_exact_rel_ids(raw.as_bytes(), &map).unwrap();

    assert_eq!(
        rewritten,
        format!(
            r#"<x:payload xmlns:x="urn:producer" xmlns:r="{R_NS}" r:id="rId9" syntax=" A &amp; B "><x:other r:id='other-link'/><x:foreign x:id="slide-link"/></x:payload>"#
        )
        .into_bytes()
    );
}

#[test]
fn comments_processing_instructions_and_attribute_syntax_are_preserved() {
    let raw = format!(
        "<?producer keep?><x:payload  xmlns:x='urn:producer' xmlns:r=\"{R_NS}\"\n r:embed = 'rId1' x:keep=\"a &amp; b\"><!-- keep --><x:item r:link=\"rId2\" /></x:payload>"
    );
    let map = HashMap::from([
        ("rId1".to_owned(), "rId&amp;'\"".to_owned()),
        ("rId2".to_owned(), "rId22".to_owned()),
    ]);

    let rewritten = rewrite_rel_ids(raw.as_bytes(), &map).unwrap();

    assert_eq!(
        rewritten,
        format!(
            "<?producer keep?><x:payload  xmlns:x='urn:producer' xmlns:r=\"{R_NS}\"\n r:embed = 'rId&amp;amp;&apos;&quot;' x:keep=\"a &amp; b\"><!-- keep --><x:item r:link=\"rId22\" /></x:payload>"
        )
        .into_bytes()
    );
}

#[test]
fn self_closing_reply_list_preserves_fixed_prefix_shadow_attributes() {
    use rpptx_oxml::comments::CommentList;

    let xml = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z"><q:unknownAnchor/><q:replyLst xmlns:p188="urn:f217:producer" p188:flag='a&#x20;b'/></q:cm></q:cmLst>"#;
    let comments = CommentList::from_xml(xml).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();

    assert!(written.contains(&format!(
        r#"<p188m:replyLst xmlns:p188m="{}""#,
        "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
    )));
    assert!(
        written.contains(r#"xmlns:p188="urn:f217:producer""#),
        "{written}"
    );
    assert!(written.contains(r#"p188:flag='a&#x20;b'"#), "{written}");
    assert_eq!(comments, CommentList::from_xml(written.as_bytes()).unwrap());
}

#[test]
fn unknown_anchor_source_subtrees_remain_byte_exact() {
    use rpptx_oxml::comments::CommentList;

    let empty = r#"<q:unknownAnchor xmlns:x='urn:f217:empty' x:flag='a&#x20;b'/>"#;
    let nonempty = r#"<q:unknownAnchor xmlns:y="urn:f217:full" y:flag='c&#x20;d'><!--anchor-comment--><?anchor keep?><y:child value='e&#x20;f'/></q:unknownAnchor>"#;
    let xml = format!(
        r#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><q:cm id="{{22222222-2222-2222-2222-222222222222}}" authorId="{{11111111-1111-1111-1111-111111111111}}" created="2026-08-29T10:11:12Z">{empty}</q:cm><q:cm id="{{33333333-3333-3333-3333-333333333333}}" authorId="{{11111111-1111-1111-1111-111111111111}}" created="2026-08-29T10:12:13Z">{nonempty}</q:cm></q:cmLst>"#
    );
    let comments = CommentList::from_xml(xml.as_bytes()).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();

    assert!(written.contains(empty), "{written}");
    assert!(written.contains(nonempty), "{written}");
    assert_eq!(comments, CommentList::from_xml(written.as_bytes()).unwrap());
}

#[test]
fn materialized_comment_text_precedes_byte_preserved_extension_list() {
    use rpptx_oxml::comments::CommentList;

    let extension = r#"<p:extLst x:flag='a&#x20;b'><p:ext uri='comment-ext'><!--comment-extension--><x:payload value='c&#x20;d'/></p:ext></p:extLst>"#;
    let xml = format!(
        r#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:p="{P_NS}" xmlns:x="urn:f217"><q:cm id="{{22222222-2222-2222-2222-222222222222}}" authorId="{{11111111-1111-1111-1111-111111111111}}" created="2026-08-29T10:11:12Z"><q:unknownAnchor/>{extension}</q:cm></q:cmLst>"#
    );
    let mut comments = CommentList::from_xml(xml.as_bytes()).unwrap();
    comments.comments[0].set_text("materialized comment");

    let written = comments.to_xml().unwrap();
    assert!(
        written
            .windows(extension.len())
            .any(|window| window == extension.as_bytes()),
        "{}",
        String::from_utf8_lossy(&written)
    );
    assert_order(&written, &["<q:unknownAnchor", "<p188:txBody", extension]);
    let reopened = CommentList::from_xml(&written).unwrap();
    assert_eq!(reopened.comments[0].text(), "materialized comment");
    assert!(
        reopened
            .to_xml()
            .unwrap()
            .windows(extension.len())
            .any(|window| window == extension.as_bytes())
    );
}

#[test]
fn materialized_reply_text_precedes_byte_preserved_extension_list() {
    use rpptx_oxml::comments::{Comment, CommentList};

    let extension = r#"<p:extLst x:flag='e&#x20;f'><p:ext uri='reply-ext'><!--reply-extension--><x:payload value='g&#x20;h'/></p:ext></p:extLst>"#;
    let xml = format!(
        r#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:p="{P_NS}" xmlns:x="urn:f217"><q:cm id="{{22222222-2222-2222-2222-222222222222}}" authorId="{{11111111-1111-1111-1111-111111111111}}" created="2026-08-29T10:11:12Z"><q:unknownAnchor/><q:replyLst><q:reply id="{{33333333-3333-3333-3333-333333333333}}" authorId="{{11111111-1111-1111-1111-111111111111}}" created="2026-08-29T10:12:13Z">{extension}</q:reply></q:replyLst></q:cm></q:cmLst>"#
    );
    let mut comments = CommentList::from_xml(xml.as_bytes()).unwrap();
    let mut reply = comments.comments[0].replies()[0].clone();
    reply.set_text("materialized reply");
    let mut comment = Comment::new(
        "{22222222-2222-2222-2222-222222222222}",
        "{11111111-1111-1111-1111-111111111111}",
        "2026-08-29T10:11:12Z",
        "host comment",
    )
    .unwrap();
    comment.add_reply(reply).unwrap();
    comments.comments[0] = comment;

    let written = comments.to_xml().unwrap();
    assert!(
        written
            .windows(extension.len())
            .any(|window| window == extension.as_bytes()),
        "{}",
        String::from_utf8_lossy(&written)
    );
    assert_order(
        &written,
        &["<p188:reply", "<p188:txBody", extension, "</p188:reply>"],
    );
    let reopened = CommentList::from_xml(&written).unwrap();
    assert_eq!(
        reopened.comments[0].replies()[0].text(),
        "materialized reply"
    );
    assert!(
        reopened
            .to_xml()
            .unwrap()
            .windows(extension.len())
            .any(|window| window == extension.as_bytes())
    );
}

#[test]
fn malformed_preserved_xml_returns_an_error() {
    for raw in [
        br#"<x:payload xmlns:x="urn:producer">"#.as_slice(),
        br#"<x:payload xmlns:x="urn:producer"><x:item></x:payload>"#.as_slice(),
        br#"<x:payload xmlns:x="urn:producer" bad="unterminated/>"#.as_slice(),
    ] {
        assert!(rewrite_rel_ids(raw, &HashMap::new()).is_err());
    }
}

#[test]
fn every_corpus_preserved_payload_is_identity_with_an_empty_map() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut xml_parts = 0usize;
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, raw) in &package.parts {
            if !part.ends_with(".xml") {
                continue;
            }
            assert_eq!(
                rewrite_rel_ids(raw, &HashMap::new())
                    .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path)),
                *raw,
                "{} {part}",
                entry.path
            );
            xml_parts += 1;
        }
    }
    assert!(xml_parts > 0);
}

#[test]
fn visibility_and_header_footer_inputs_round_trip_in_schema_order() {
    let slide_xml = format!(
        r#"<q:sld xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer" showMasterSp="false" x:keep="slide"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMapOvr><d:masterClrMapping/></q:clrMapOvr><q:extLst/></q:sld>"#
    );
    let layout_xml = format!(
        r#"<q:sldLayout xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer" showMasterSp="0" x:keep="layout"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMapOvr><d:masterClrMapping/></q:clrMapOvr><x:beforeHf/><q:hf sldNum="true" hdr="false" ftr="0" dt="1" x:keep="hf"><x:child/></q:hf><q:timing/><q:transition/><q:extLst/></q:sldLayout>"#
    );
    let master_xml = format!(
        r#"<q:sldMaster xmlns:q="{P_NS}" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:producer"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><q:sldLayoutIdLst/><q:timing/><q:hf sldNum="0" ftr="1" dt="false" x:keep="master"><x:child/></q:hf><x:afterHf/><q:txStyles><q:titleStyle/><q:bodyStyle/><q:otherStyle/></q:txStyles><q:extLst/></q:sldMaster>"#
    );

    let slide = CT_Slide::from_xml(slide_xml.as_bytes()).unwrap();
    let layout = CT_SlideLayout::from_xml(layout_xml.as_bytes()).unwrap();
    let master = CT_SlideMaster::from_xml(master_xml.as_bytes()).unwrap();

    assert_eq!(slide.show_master_shapes, Some(false));
    assert_eq!(layout.show_master_shapes, Some(false));
    let layout_hf = layout.header_footer.as_ref().unwrap();
    assert_eq!(layout_hf.slide_number, Some(true));
    assert_eq!(layout_hf.header, Some(false));
    assert_eq!(layout_hf.footer, Some(false));
    assert_eq!(layout_hf.date_time, Some(true));
    let master_hf = master.header_footer.as_ref().unwrap();
    assert_eq!(master_hf.slide_number, Some(false));
    assert_eq!(master_hf.footer, Some(true));
    assert_eq!(master_hf.date_time, Some(false));

    let written_slide = slide.to_xml().unwrap();
    let written_layout = layout.to_xml().unwrap();
    let written_master = master.to_xml().unwrap();
    let slide_text = std::str::from_utf8(&written_slide).unwrap();
    let layout_text = std::str::from_utf8(&written_layout).unwrap();
    let master_text = std::str::from_utf8(&written_master).unwrap();

    assert!(slide_text.starts_with("<p:sld "));
    assert!(slide_text.contains("showMasterSp=\"0\""));
    assert!(slide_text.contains("x:keep=\"slide\""));
    assert!(layout_text.contains(
        "<p:hf sldNum=\"1\" hdr=\"0\" ftr=\"0\" dt=\"1\" x:keep=\"hf\"><x:child/></p:hf>"
    ));
    assert!(
        master_text
            .contains("<p:hf sldNum=\"0\" ftr=\"1\" dt=\"0\" x:keep=\"master\"><x:child/></p:hf>")
    );
    assert!(!layout_text.contains("<q:hf"));
    assert!(!master_text.contains("<q:hf"));
    assert_order(
        &written_layout,
        &[
            "<p:clrMapOvr",
            "<x:beforeHf",
            "<p:hf",
            "<q:timing",
            "<q:transition",
            "<q:extLst",
        ],
    );
    assert_order(
        &written_master,
        &[
            "<p:clrMap",
            "<q:sldLayoutIdLst",
            "<q:timing",
            "<p:hf",
            "<x:afterHf",
            "<p:txStyles",
            "<q:extLst",
        ],
    );
    assert_eq!(slide, CT_Slide::from_xml(&written_slide).unwrap());
    assert_eq!(layout, CT_SlideLayout::from_xml(&written_layout).unwrap());
    assert_eq!(master, CT_SlideMaster::from_xml(&written_master).unwrap());
}

#[test]
fn all_corpus_slide_layout_and_master_parts_reparse_after_visibility_typing() {
    let Some(corpus) = require_or_skip_corpus() else {
        return;
    };
    verify_fetched_corpus();
    let mut decks = 0usize;
    let mut counts = [0usize; 3];
    for entry in manifest_entries() {
        let package = OpcPackage::open(corpus.join(entry.path)).unwrap();
        for (part, content_type) in &package.content_types.overrides {
            let Some(xml) = package.get_part(part) else {
                continue;
            };
            match content_type.as_str() {
                content_types::SLIDE => {
                    let parsed = CT_Slide::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    let _ = parsed.show_master_shapes;
                    assert_eq!(
                        parsed,
                        CT_Slide::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    counts[0] += 1;
                }
                content_types::SLIDE_LAYOUT => {
                    let parsed = CT_SlideLayout::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    let _ = (parsed.show_master_shapes, parsed.header_footer.as_ref());
                    assert_eq!(
                        parsed,
                        CT_SlideLayout::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    counts[1] += 1;
                }
                content_types::SLIDE_MASTER => {
                    let parsed = CT_SlideMaster::from_xml(xml)
                        .unwrap_or_else(|error| panic!("{} {part}: {error}", entry.path));
                    let _ = parsed.header_footer.as_ref();
                    assert_eq!(
                        parsed,
                        CT_SlideMaster::from_xml(&parsed.to_xml().unwrap()).unwrap()
                    );
                    counts[2] += 1;
                }
                _ => {}
            }
        }
        decks += 1;
    }
    assert_eq!(decks, EXPECTED_DECKS);
    assert!(
        counts.iter().all(|count| *count > 0),
        "part counts: {counts:?}"
    );
}

#[test]
fn modern_comments_and_replies_preserve_order_and_unmodelled_xml() {
    use rpptx_oxml::comments::{Comment, CommentAuthorList, CommentList, CommentReply};

    assert!(
        Comment::new(
            "{22222222-2222-2222-2222-222222222222}",
            "{11111111-1111-1111-1111-111111111111}",
            "short",
            "comment",
        )
        .is_err()
    );
    assert!(
        CommentReply::new(
            "{33333333-3333-3333-3333-333333333333}",
            "{11111111-1111-1111-1111-111111111111}",
            "2026-02-30T10:12:13+24:00",
            "reply",
        )
        .is_err()
    );

    let authors = CommentAuthorList::from_xml(
        br#"<c:authorLst xmlns:c="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:x="urn:f217"><x:before/><c:author id="{11111111-1111-1111-1111-111111111111}" name="Ada &amp; Co" initials="AC" userId="ada@example.test" providerId="test"><x:author-child marker="kept"/></c:author><x:after/></c:authorLst>"#,
    )
    .expect("parse aliased modern authors");
    let author_xml = authors.to_xml().expect("write modern authors");
    assert!(author_xml.starts_with(b"<p188:authorLst"));
    assert!(
        author_xml
            .windows(b"<x:author-child marker=\"kept\"/>".len())
            .any(|window| window == b"<x:author-child marker=\"kept\"/>")
    );
    assert_eq!(authors, CommentAuthorList::from_xml(&author_xml).unwrap());

    let comments = CommentList::from_xml(
        br#"<c:cmLst xmlns:c="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:x="urn:f217"><x:before/><c:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z" x:flag="kept"><x:sldMkLst/><c:unknownAnchor/><x:after-anchor/><c:replyLst><x:before-reply/><c:reply id="{33333333-3333-3333-3333-333333333333}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:12:13+00:00"><c:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>reply</a:t></a:r></a:p></c:txBody></c:reply><x:after-reply/></c:replyLst><c:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>comment</a:t></a:r></a:p></c:txBody><x:after-text/></c:cm><x:after/></c:cmLst>"#,
    )
    .expect("parse aliased modern comments");
    let comment_xml = comments.to_xml().expect("write modern comments");
    let written = String::from_utf8(comment_xml.clone()).unwrap();
    assert!(written.starts_with("<p188:cmLst"));
    assert_order(
        written.as_bytes(),
        &[
            "<x:sldMkLst",
            "<c:unknownAnchor",
            "<x:after-anchor",
            "<p188:replyLst",
            "<c:txBody",
            "<x:after-text",
        ],
    );
    assert!(written.contains("<x:before-reply/>"));
    assert!(written.contains("<x:after-reply/>"));
    assert_eq!(comments, CommentList::from_xml(&comment_xml).unwrap());
}

#[test]
fn sections_notes_and_handout_settings_write_in_schema_order() {
    use rpptx_oxml::notes_parts::{CT_HandoutMaster, CT_NotesMaster};

    let presentation = CT_Presentation::from_xml(
        br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst><p14:section name="One" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst><p:extLst><p:ext uri="raw"><x:kept xmlns:x="urn:f217"/></p:ext></p:extLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>"#,
    )
    .expect("parse sections");
    assert_eq!(presentation.sections().len(), 1);
    assert_eq!(presentation.sections()[0].slide_ids, vec![256]);
    let presentation_xml = presentation.to_xml().unwrap();
    let written = String::from_utf8(presentation_xml.clone()).unwrap();
    assert_order(
        written.as_bytes(),
        &["<p:sldIdLst", "<p:notesSz", "<p:extLst", "<p14:sectionLst"],
    );
    assert!(written.contains("<x:kept"));
    assert_eq!(
        presentation,
        CT_Presentation::from_xml(&presentation_xml).unwrap()
    );
    let mut without_sections = presentation.clone();
    without_sections.set_sections(Vec::new()).unwrap();
    let without_sections = String::from_utf8(without_sections.to_xml().unwrap()).unwrap();
    assert!(!without_sections.contains("p14:sectionLst"));
    assert!(without_sections.contains("<p:extLst"));

    let notes = CT_NotesMaster::from_xml(
        br#"<q:notesMaster xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><x:before xmlns:x="urn:f217"/><q:hf sldNum="0" hdr="1" ftr="0" dt="1"><x:hf-child/></q:hf><x:after/></q:notesMaster>"#,
    )
    .expect("parse typed notes header footer");
    assert_eq!(
        notes.header_footer.as_ref().unwrap().slide_number,
        Some(false)
    );
    let notes_xml = String::from_utf8(notes.to_xml().unwrap()).unwrap();
    assert_order(
        notes_xml.as_bytes(),
        &["<p:clrMap", "<x:before", "<p:hf", "<x:after"],
    );

    let handout = CT_HandoutMaster::from_xml(
        br#"<q:handoutMaster xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cSld><q:spTree><q:nvGrpSpPr/><q:grpSpPr/></q:spTree></q:cSld><q:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><q:hf hdr="0"/></q:handoutMaster>"#,
    )
    .expect("parse handout master");
    assert_eq!(handout.header_footer.as_ref().unwrap().header, Some(false));
    assert_eq!(
        handout,
        CT_HandoutMaster::from_xml(&handout.to_xml().unwrap()).unwrap()
    );
}

#[test]
fn sections_require_the_direct_p14_extension_contract() {
    use rpptx_oxml::presentation::Section;

    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:x="urn:f217"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="wrong"><p14:sectionLst><p14:section name="Wrong" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><x:wrapper><p14:sectionLst><p14:section name="Nested" id="{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section></p14:sectionLst></x:wrapper></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    assert!(presentation.sections().is_empty());
    presentation
        .set_sections(vec![
            Section::new("{CCCCCCCC-CCCC-CCCC-CCCC-CCCCCCCCCCCC}", "Typed", vec![256]).unwrap(),
        ])
        .unwrap();
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    assert!(written.contains("name=\"Wrong\""));
    assert!(written.contains("<x:wrapper><p14:sectionLst>"));
    assert_eq!(written.matches("name=\"Typed\"").count(), 1);
}

#[test]
fn sections_expand_empty_extension_lists_and_preserve_section_list_sidecars() {
    use rpptx_oxml::presentation::Section;

    let empty = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><q:extLst xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:f217" x:keep='a&#x20;b'/></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(empty).unwrap();
    presentation
        .set_sections(vec![
            Section::new("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}", "One", vec![256]).unwrap(),
        ])
        .unwrap();
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    assert!(written.contains(r#"<q:extLst xmlns:q="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:x="urn:f217" x:keep='a&#x20;b'>"#));
    assert_eq!(written.matches("<q:extLst").count(), 1);
    CT_Presentation::from_xml(written.as_bytes()).unwrap();

    let raw = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:x="urn:f217"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst x:flag='a&#x20;b'><!--before--><?sections keep?>direct<x:before/><p14:section name="Old" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section><x:after/><!--after--></p14:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(raw).unwrap();
    let clean = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    assert!(clean.contains(
        "<p14:sectionLst x:flag='a&#x20;b'><!--before--><?sections keep?>direct<x:before/>"
    ));
    presentation
        .set_sections(vec![
            Section::new("{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}", "New", vec![256]).unwrap(),
        ])
        .unwrap();
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    assert!(written.contains("<p14:sectionLst x:flag='a&#x20;b'>"));
    for raw in [
        "<!--before-->",
        "<?sections keep?>",
        "direct",
        "<x:before/>",
        "<x:after/>",
        "<!--after-->",
    ] {
        assert!(written.contains(raw), "missing section sidecar {raw}");
    }
    assert!(written.contains("name=\"New\""));
    assert!(!written.contains("name=\"Old\""), "{written}");
    CT_Presentation::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn self_closing_slide_extension_list_is_reused_for_modern_comments() {
    let xml = format!(
        r#"<p:sld xmlns:p="{P_NS}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="{R_NS}"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr><q:extLst xmlns:q="{P_NS}" xmlns:x="urn:f217" x:keep='a&#x20;b'/></p:sld>"#
    );
    let mut slide = CT_Slide::from_xml(xml.as_bytes()).unwrap();
    slide.ensure_modern_comment_relationship("rId9").unwrap();
    let written = String::from_utf8(slide.to_xml().unwrap()).unwrap();
    assert_eq!(written.matches("<q:extLst").count(), 1);
    assert!(!written.contains("<p:extLst"));
    assert!(written.contains(&format!(
        r#"<q:extLst xmlns:q="{P_NS}" xmlns:x="urn:f217" x:keep='a&#x20;b'>"#
    )));
    assert!(written.contains("r:id=\"rId9\""));
    CT_Slide::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn modern_comment_status_values_are_strict() {
    use rpptx_oxml::comments::{
        Comment, CommentAuthor, CommentAuthorList, CommentList, CommentReply,
    };
    const AUTHOR: &str = "{11111111-1111-1111-1111-111111111111}";
    let invalid = format!(
        r#"<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:cm id="{{22222222-2222-2222-2222-222222222222}}" authorId="{AUTHOR}" status="bogus" created="2026-08-29T10:11:12Z"><p188:unknownAnchor/></p188:cm></p188:cmLst>"#
    );
    assert!(CommentList::from_xml(invalid.as_bytes()).is_err());
    let mut comment = Comment::new(
        "{22222222-2222-2222-2222-222222222222}",
        AUTHOR,
        "2026-08-29T10:11:12Z",
        "comment",
    )
    .unwrap();
    comment.status = Some("bogus".to_owned());
    let mut list = CommentList::new();
    list.comments.push(comment);
    assert!(list.to_xml().is_err());
    let mut comment = Comment::new(
        "{22222222-2222-2222-2222-222222222222}",
        AUTHOR,
        "2026-08-29T10:11:12Z",
        "comment",
    )
    .unwrap();
    let mut reply = CommentReply::new(
        "{33333333-3333-3333-3333-333333333333}",
        AUTHOR,
        "2026-08-29T10:12:13Z",
        "reply",
    )
    .unwrap();
    reply.status = Some("bogus".to_owned());
    comment.add_reply(reply).unwrap();
    let mut list = CommentList::new();
    list.comments.push(comment);
    assert!(list.to_xml().is_err());

    let mut authors = CommentAuthorList::new();
    let mut author = CommentAuthor::new(AUTHOR, "Ada", None, "ada@example.test", "test").unwrap();
    author.id = "bad".to_owned();
    authors.authors.push(author);
    assert!(authors.to_xml().is_err());
    let mut comment = Comment::new(
        "{22222222-2222-2222-2222-222222222222}",
        AUTHOR,
        "2026-08-29T10:11:12Z",
        "comment",
    )
    .unwrap();
    comment.id = "bad".to_owned();
    let mut list = CommentList::new();
    list.comments.push(comment);
    assert!(list.to_xml().is_err());
    let mut comment = Comment::new(
        "{22222222-2222-2222-2222-222222222222}",
        AUTHOR,
        "2026-08-29T10:11:12Z",
        "comment",
    )
    .unwrap();
    comment.created = "bad".to_owned();
    let mut list = CommentList::new();
    list.comments.push(comment);
    assert!(list.to_xml().is_err());
    let mut comment = Comment::new(
        "{22222222-2222-2222-2222-222222222222}",
        AUTHOR,
        "2026-08-29T10:11:12Z",
        "comment",
    )
    .unwrap();
    let mut reply = CommentReply::new(
        "{33333333-3333-3333-3333-333333333333}",
        AUTHOR,
        "2026-08-29T10:12:13Z",
        "reply",
    )
    .unwrap();
    reply.created = "bad".to_owned();
    comment.add_reply(reply).unwrap();
    let mut list = CommentList::new();
    list.comments.push(comment);
    assert!(list.to_xml().is_err());
}

#[test]
fn comment_sidecars_preserve_shadows_raw_nodes_and_attribute_lexemes() {
    use rpptx_oxml::comments::{CommentAuthorList, CommentList};

    let authors = br#"<q:authorLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><!--before--><?authors keep?>raw<q:author id="{11111111-1111-1111-1111-111111111111}" name="Ada" userId="ada@example.test" providerId="test" xmlns:p188="urn:producer" x:flag='a&#x20;b' xmlns:x="urn:f217"><!--inside--><?author keep?>shell<p188:producer xmlns:p188="urn:producer"/><a:producer xmlns:a="urn:drawing-producer"/></q:author><!--after--></q:authorLst>"#;
    let authors = CommentAuthorList::from_xml(authors).unwrap();
    let written = String::from_utf8(authors.to_xml().unwrap()).unwrap();
    assert!(written.contains("<!--before--><?authors keep?>raw"));
    assert!(written.contains("<!--inside--><?author keep?>shell"));
    assert!(written.contains("x:flag='a&#x20;b'"));
    assert!(written.contains("<p188:producer xmlns:p188=\"urn:producer\"/>"));
    assert!(written.contains("<a:producer xmlns:a=\"urn:drawing-producer\"/>"));
    assert_eq!(written.matches("xmlns:p188=\"urn:producer\"").count(), 1);
    CommentAuthorList::from_xml(written.as_bytes()).unwrap();

    let comments = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><!--before--><?comments keep?>raw<q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z" xmlns:p188="urn:producer" x:flag='a&#x20;b' xmlns:x="urn:f217"><q:unknownAnchor/><p188:producer xmlns:p188="urn:producer"/><q:replyLst><!--reply-list--><?reply-list keep?>text<q:reply id="{33333333-3333-3333-3333-333333333333}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:12:13Z" xmlns:p188="urn:producer"><!--reply--><?reply keep?>body<a:producer xmlns:a="urn:drawing-producer"/></q:reply></q:replyLst></q:cm><!--after--></q:cmLst>"#;
    let comments = CommentList::from_xml(comments).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();
    assert!(written.contains("<!--before--><?comments keep?>raw"));
    assert!(written.contains("<!--reply-list--><?reply-list keep?>text"));
    assert!(written.contains("<!--reply--><?reply keep?>body"));
    assert!(written.contains("x:flag='a&#x20;b'"));
    assert!(written.contains("<p188:producer xmlns:p188=\"urn:producer\"/>"));
    assert!(written.contains("<a:producer xmlns:a=\"urn:drawing-producer\"/>"));
    assert_eq!(written.matches("xmlns:p188=\"urn:producer\"").count(), 1);
    CommentList::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn modelled_comment_shell_preserves_parent_owned_fixed_prefix_shadows() {
    use rpptx_oxml::comments::CommentList;

    let xml = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z" xmlns:p188="urn:producer" xmlns:a="urn:drawing-producer"><q:unknownAnchor/><p188:producer/><a:producer/></q:cm></q:cmLst>"#;
    let comments = CommentList::from_xml(xml).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();
    assert!(written.contains(
        r#"<p188m:cm xmlns:p188m="http://schemas.microsoft.com/office/powerpoint/2018/8/main""#
    ));
    assert!(written.contains(r#"xmlns:p188="urn:producer" xmlns:a="urn:drawing-producer""#));
    assert!(written.contains("<p188:producer/><a:producer/>"));
    assert_eq!(comments, CommentList::from_xml(written.as_bytes()).unwrap());
}

#[test]
fn typed_section_and_slide_id_sidecars_remain_lexically_exact_when_dirty() {
    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:x="urn:f217"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="900" r:id="rId2"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst><p14:section name="One" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" x:flag='a&#x20;b'><!--section--><?section keep?>section-text<![CDATA[section-cdata]]><p14:sldIdLst x:list='c&#x20;d'><!--list--><?list keep?>list-text<![CDATA[list-cdata]]><p14:sldId id="256" x:item='e&#x20;f'><!--item--><?item keep?>item-text<![CDATA[item-cdata]]><x:item/></p14:sldId><p14:sldId id="900"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    presentation.remove_slide_from_sections(900);
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    for raw in [
        "x:flag='a&#x20;b'",
        "<!--section-->",
        "<?section keep?>",
        "section-text",
        "<![CDATA[section-cdata]]>",
        "x:list='c&#x20;d'",
        "<!--list-->",
        "<?list keep?>",
        "list-text",
        "<![CDATA[list-cdata]]>",
        "x:item='e&#x20;f'",
        "<!--item-->",
        "<?item keep?>",
        "item-text",
        "<![CDATA[item-cdata]]>",
        "<x:item/>",
    ] {
        assert!(
            written.contains(raw),
            "missing nested section sidecar {raw}"
        );
    }
    assert!(!written.contains("<p14:sldId id=\"900\""));
    CT_Presentation::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn dirty_alias_only_section_list_binds_generated_fixed_prefix_children() {
    use rpptx_oxml::presentation::Section;

    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q:sectionLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><q:section name="Old" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><q:sldIdLst><q:sldId id="256"/></q:sldIdLst></q:section></q:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    presentation
        .set_sections(vec![
            Section::new("{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}", "New", vec![256]).unwrap(),
        ])
        .unwrap();
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    assert!(written.contains(r#"<q:sectionLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:section xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main""#));
    let reparsed = CT_Presentation::from_xml(written.as_bytes()).unwrap();
    assert_eq!(reparsed.sections()[0].name.as_deref(), Some("New"));
}

#[test]
fn list_owned_comment_shadows_and_typed_drawing_children_keep_their_namespaces() {
    use rpptx_oxml::comments::{CommentAuthorList, CommentList};

    let list = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:p188="urn:list-producer"><p188:producer/><q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z" xmlns:a="urn:drawing-producer" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:unknownAnchor/><a:producer/><q:txBody><d:bodyPr/><d:lstStyle/><d:p/></q:txBody></q:cm></q:cmLst>"#;
    let comments = CommentList::from_xml(list).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();
    assert!(written.contains(r#"xmlns:p188="urn:list-producer"><p188:producer/>"#));
    assert!(written.contains("xmlns:a=\"urn:drawing-producer\""));
    assert!(written.contains(r#"<q:txBody><d:bodyPr/><d:lstStyle/><d:p/></q:txBody>"#));
    CommentList::from_xml(written.as_bytes()).unwrap();

    let authors = br#"<q:authorLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="urn:list-drawing"><a:producer/><q:author id="{11111111-1111-1111-1111-111111111111}" name="Ada" userId="ada@example.test" providerId="test"/></q:authorLst>"#;
    let authors = CommentAuthorList::from_xml(authors).unwrap();
    let written = String::from_utf8(authors.to_xml().unwrap()).unwrap();
    assert!(written.contains(r#"xmlns:a="urn:list-drawing"><a:producer/>"#));
    CommentAuthorList::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn section_insertion_ignores_self_closing_descendants_and_alias_sections_clear_cleanly() {
    use rpptx_oxml::presentation::Section;

    let nonempty = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="producer"/></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(nonempty).unwrap();
    presentation
        .set_sections(vec![
            Section::new("{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}", "One", vec![256]).unwrap(),
        ])
        .unwrap();
    let written = presentation.to_xml().unwrap();
    assert!(String::from_utf8_lossy(&written).contains("<p:ext uri=\"producer\"/>"));
    CT_Presentation::from_xml(&written).unwrap();

    let alias = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q:sectionLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><q:section name="Old" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><q:sldIdLst><q:sldId id="256"/></q:sldIdLst></q:section></q:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(alias).unwrap();
    presentation.set_sections(Vec::new()).unwrap();
    let written = presentation.to_xml().unwrap();
    assert!(!String::from_utf8_lossy(&written).contains("sectionLst"));
    CT_Presentation::from_xml(&written).unwrap();
}

#[test]
fn appended_author_stays_before_the_original_trailing_raw_sidecar() {
    use rpptx_oxml::comments::{CommentAuthor, CommentAuthorList};

    let xml = br#"<q:authorLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:x="urn:f217"><q:author id="{11111111-1111-1111-1111-111111111111}" name="Ada" userId="ada@example.test" providerId="test"/><x:tail/></q:authorLst>"#;
    let mut authors = CommentAuthorList::from_xml(xml).unwrap();
    authors.authors.push(
        CommentAuthor::new(
            "{22222222-2222-2222-2222-222222222222}",
            "Grace",
            None,
            "grace@example.test",
            "test",
        )
        .unwrap(),
    );
    let written = authors.to_xml().unwrap();
    assert_order(&written, &["name=\"Ada\"", "name=\"Grace\"", "<x:tail/>"]);
    CommentAuthorList::from_xml(&written).unwrap();
}

#[test]
fn comment_fallback_shadows_and_inherited_reply_drawing_namespace_survive() {
    use rpptx_oxml::comments::CommentList;

    let xml = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:p188="urn:f217:primary" xmlns:p188m="urn:f217:fallback" xmlns:a="urn:f217:drawing" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><p188:primary/><p188m:fallback/><a:producer/><q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z"><q:unknownAnchor/><q:replyLst><q:reply id="{33333333-3333-3333-3333-333333333333}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:12:13Z"><q:txBody><d:bodyPr/><d:lstStyle/><d:p/></q:txBody></q:reply></q:replyLst></q:cm></q:cmLst>"#;
    let comments = CommentList::from_xml(xml).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();
    assert!(written.contains(r#"xmlns:p188="urn:f217:primary""#));
    assert!(written.contains(r#"xmlns:p188m="urn:f217:fallback""#));
    assert!(written.contains("<p188:primary/><p188m:fallback/><a:producer/>"));
    assert!(
        written.contains(r#"<q:txBody><d:bodyPr/><d:lstStyle/><d:p/></q:txBody>"#),
        "{written}"
    );
    assert_eq!(comments, CommentList::from_xml(written.as_bytes()).unwrap());
}

#[test]
fn dirty_sections_preserve_producer_owned_p14_shadows_on_typed_shells() {
    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="900" r:id="rId2"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q:sectionLst><q:section name="One" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" xmlns:p14="urn:f217:section"><p14:section-producer/><q:sldIdLst xmlns:p14="urn:f217:list"><p14:list-producer/><q:sldId id="256" xmlns:p14="urn:f217:item"><p14:item-producer/></q:sldId><q:sldId id="900"/></q:sldIdLst></q:section></q:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    presentation.remove_slide_from_sections(900);
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    assert_eq!(written.matches("xmlns:p14=\"urn:f217:").count(), 3);
    assert!(written.contains("<p14:section-producer/>"));
    assert!(written.contains("<p14:list-producer/>"));
    assert!(written.contains("<p14:item-producer/>"));
    assert!(written.contains(&format!(
        r#"<p14m:section xmlns:p14m="{}""#,
        "http://schemas.microsoft.com/office/powerpoint/2010/main"
    )));
    CT_Presentation::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn section_slide_id_raw_boundaries_follow_original_slide_identities() {
    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:x="urn:f217"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="900" r:id="rId2"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst><p14:section name="One" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><p14:sldIdLst><p14:sldId id="256"/><x:between/><p14:sldId id="900"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>"#;

    let mut removed = CT_Presentation::from_xml(xml).unwrap();
    removed.remove_slide_from_sections(256);
    let written = String::from_utf8(removed.to_xml().unwrap()).unwrap();
    let section_list = &written[written.rfind("<p14:sldIdLst").unwrap()..];
    assert_order(section_list.as_bytes(), &["<x:between/>", "id=\"900\""]);
    CT_Presentation::from_xml(written.as_bytes()).unwrap();

    let mut reordered = CT_Presentation::from_xml(xml).unwrap();
    let mut sections = reordered.sections().to_vec();
    sections[0].slide_ids.swap(0, 1);
    reordered.set_sections(sections).unwrap();
    let written = String::from_utf8(reordered.to_xml().unwrap()).unwrap();
    let section_list = &written[written.rfind("<p14:sldIdLst").unwrap()..];
    assert_order(
        section_list.as_bytes(),
        &["<x:between/>", "id=\"900\"", "id=\"256\""],
    );
    CT_Presentation::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn dirty_section_rewrite_recognizes_inherited_alternate_child_aliases() {
    use rpptx_oxml::presentation::Section;

    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:s="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q:sectionLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><s:section name="Old" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><s:sldIdLst><s:sldId id="256"/></s:sldIdLst></s:section></q:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut replaced = CT_Presentation::from_xml(xml).unwrap();
    replaced
        .set_sections(vec![
            Section::new("{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}", "New", vec![256]).unwrap(),
        ])
        .unwrap();
    let written = String::from_utf8(replaced.to_xml().unwrap()).unwrap();
    assert!(!written.contains("name=\"Old\""));
    assert_eq!(written.matches("name=\"New\"").count(), 1);
    let reopened = CT_Presentation::from_xml(written.as_bytes()).unwrap();
    assert_eq!(reopened.sections().len(), 1);
    assert_eq!(reopened.sections()[0].name.as_deref(), Some("New"));

    let mut cleared = CT_Presentation::from_xml(xml).unwrap();
    cleared.set_sections(Vec::new()).unwrap();
    let written = String::from_utf8(cleared.to_xml().unwrap()).unwrap();
    assert!(!written.contains("sectionLst"));
    assert!(
        CT_Presentation::from_xml(written.as_bytes())
            .unwrap()
            .sections()
            .is_empty()
    );
}

#[test]
fn comment_serialization_rejects_when_every_model_prefix_is_producer_owned() {
    use rpptx_oxml::comments::CommentList;

    let xml = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:p188="urn:f217:p188" xmlns:p188m="urn:f217:p188m" xmlns:p188model="urn:f217:p188model"><p188:one/><p188m:two/><p188model:three/><q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z"><q:unknownAnchor/></q:cm></q:cmLst>"#;
    let comments = CommentList::from_xml(xml).unwrap();
    assert_eq!(
        comments.to_xml().unwrap_err().to_string(),
        "invalid value: no unshadowed modern comment model prefix is available"
    );
}

#[test]
fn section_serialization_rejects_when_every_model_prefix_is_producer_owned() {
    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="900" r:id="rId2"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q:sectionLst><q:section name="One" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" xmlns:p14="urn:f217:p14" xmlns:p14m="urn:f217:p14m" xmlns:p14model="urn:f217:p14model"><p14:one/><p14m:two/><p14model:three/><q:sldIdLst><q:sldId id="256"/><q:sldId id="900"/></q:sldIdLst></q:section></q:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    presentation.remove_slide_from_sections(900);
    assert_eq!(
        presentation.to_xml().unwrap_err().to_string(),
        "invalid value: no unshadowed section model prefix is available"
    );
}

#[test]
fn inherited_p14_shadow_is_carried_into_each_typed_section_sidecar() {
    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="900" r:id="rId2"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q:sectionLst xmlns:p14="urn:f217:ancestor"><q:section name="One" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><p14:section-producer/><q:sldIdLst><p14:list-producer/><q:sldId id="256"><p14:item-producer/></q:sldId><q:sldId id="900"/></q:sldIdLst></q:section></q:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    presentation.remove_slide_from_sections(900);
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    assert_eq!(
        written.matches(r#"xmlns:p14="urn:f217:ancestor""#).count(),
        4
    );
    assert!(written.contains("<p14:section-producer/>"));
    assert!(written.contains("<p14:list-producer/>"));
    assert!(written.contains("<p14:item-producer/>"));
    assert_eq!(
        CT_Presentation::from_xml(written.as_bytes())
            .unwrap()
            .sections()[0]
            .slide_ids,
        vec![256]
    );
}

#[test]
fn fixed_comment_prefix_bindings_used_only_by_shell_attributes_survive() {
    use rpptx_oxml::comments::{CommentAuthorList, CommentList};

    let authors = br#"<q:authorLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:p188="urn:f217:authors" p188:listFlag='a&#x20;b'><q:author id="{11111111-1111-1111-1111-111111111111}" name="Ada" userId="ada@example.test" providerId="test" p188:authorFlag='c&#x20;d'/></q:authorLst>"#;
    let authors = CommentAuthorList::from_xml(authors).unwrap();
    let written = String::from_utf8(authors.to_xml().unwrap()).unwrap();
    assert!(written.contains("p188:listFlag='a&#x20;b'"));
    assert!(written.contains("p188:authorFlag='c&#x20;d'"));
    assert_eq!(
        written.matches(r#"xmlns:p188="urn:f217:authors""#).count(),
        2
    );
    assert_eq!(
        authors,
        CommentAuthorList::from_xml(written.as_bytes()).unwrap()
    );

    let comments = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:p188="urn:f217:comments" p188:listFlag='a&#x20;b'><q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z" p188:commentFlag='c&#x20;d'><q:unknownAnchor/><q:replyLst p188:replyListFlag='e&#x20;f'><q:reply id="{33333333-3333-3333-3333-333333333333}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:12:13Z" p188:replyFlag='g&#x20;h'/></q:replyLst></q:cm></q:cmLst>"#;
    let comments = CommentList::from_xml(comments).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();
    for lexical in [
        "p188:listFlag='a&#x20;b'",
        "p188:commentFlag='c&#x20;d'",
        "p188:replyListFlag='e&#x20;f'",
        "p188:replyFlag='g&#x20;h'",
    ] {
        assert!(written.contains(lexical), "missing {lexical}: {written}");
    }
    assert_eq!(
        written.matches(r#"xmlns:p188="urn:f217:comments""#).count(),
        4
    );
    assert_eq!(comments, CommentList::from_xml(written.as_bytes()).unwrap());
}

#[test]
fn fixed_section_prefix_bindings_used_only_by_shell_attributes_survive() {
    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="900" r:id="rId2"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><q:sectionLst xmlns:p14="urn:f217:section-attributes"><q:section name="One" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}" p14:sectionFlag='a&#x20;b'><q:sldIdLst p14:listFlag='c&#x20;d'><q:sldId id="256" p14:itemFlag='e&#x20;f'/><q:sldId id="900"/></q:sldIdLst></q:section></q:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    presentation.remove_slide_from_sections(900);
    let written = String::from_utf8(presentation.to_xml().unwrap()).unwrap();
    for lexical in [
        "p14:sectionFlag='a&#x20;b'",
        "p14:listFlag='c&#x20;d'",
        "p14:itemFlag='e&#x20;f'",
    ] {
        assert!(written.contains(lexical), "missing {lexical}: {written}");
    }
    assert_eq!(
        written
            .matches(r#"xmlns:p14="urn:f217:section-attributes""#)
            .count(),
        4
    );
    CT_Presentation::from_xml(written.as_bytes()).unwrap();
}

#[test]
fn typed_comment_and_reply_text_body_root_sidecars_remain_lexically_exact() {
    use rpptx_oxml::comments::CommentList;

    let xml = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z"><q:unknownAnchor/><q:replyLst><q:reply id="{33333333-3333-3333-3333-333333333333}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:12:13Z"><q:txBody xmlns:y="urn:f217:reply" y:flag='c&#x20;d'><a:bodyPr/><a:lstStyle/><a:p/><y:tail/></q:txBody></q:reply></q:replyLst><q:txBody xmlns:x="urn:f217:comment" x:flag='a&#x20;b'><a:bodyPr/><a:lstStyle/><a:p/><x:tail/></q:txBody></q:cm></q:cmLst>"#;
    let comments = CommentList::from_xml(xml).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();
    assert!(written.contains(r#"xmlns:x="urn:f217:comment" x:flag='a&#x20;b'"#));
    assert!(written.contains(r#"xmlns:y="urn:f217:reply" y:flag='c&#x20;d'"#));
    assert!(written.contains("<x:tail/>"));
    assert!(written.contains("<y:tail/>"));
    assert_eq!(comments, CommentList::from_xml(written.as_bytes()).unwrap());
}

#[test]
fn clearing_sections_with_direct_raw_payload_fails_without_mutation() {
    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:x="urn:f217"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst x:flag='a&#x20;b'><!--keep--><?sections keep?><x:payload/><p14:section name="One" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><p14:sldIdLst><p14:sldId id="256"/></p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    let before = presentation.to_xml().unwrap();
    assert_eq!(
        presentation
            .set_sections(Vec::new())
            .unwrap_err()
            .to_string(),
        "invalid value: cannot clear sections while preserving direct section-list raw payload"
    );
    assert_eq!(presentation.sections().len(), 1);
    assert_eq!(presentation.to_xml().unwrap(), before);
    CT_Presentation::from_xml(&before).unwrap();
}

#[test]
fn inherited_default_p14_section_namespace_supports_dirty_replacement() {
    use rpptx_oxml::presentation::Section;

    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="900" r:id="rId2"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}" xmlns="http://schemas.microsoft.com/office/powerpoint/2010/main"><sectionLst><section name="Old" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><sldIdLst><sldId id="256"/></sldIdLst></section></sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    presentation
        .set_sections(vec![
            Section::new("{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}", "New", vec![900]).unwrap(),
        ])
        .unwrap();

    let written = presentation.to_xml().unwrap();
    let text = String::from_utf8(written.clone()).unwrap();
    assert!(!text.contains("name=\"Old\""), "{text}");
    assert_eq!(text.matches("name=\"New\"").count(), 1, "{text}");
    let reopened = CT_Presentation::from_xml(&written).unwrap();
    assert_eq!(reopened.sections().len(), 1);
    assert_eq!(reopened.sections()[0].name.as_deref(), Some("New"));
    assert_eq!(reopened.sections()[0].slide_ids, vec![900]);
}

#[test]
fn inherited_default_p14_section_namespace_supports_clear() {
    let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:notesSz cx="1" cy="1"/><p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}" xmlns="http://schemas.microsoft.com/office/powerpoint/2010/main"><sectionLst><section name="Old" id="{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}"><sldIdLst><sldId id="256"/></sldIdLst></section></sectionLst></p:ext></p:extLst></p:presentation>"#;
    let mut presentation = CT_Presentation::from_xml(xml).unwrap();
    presentation.set_sections(Vec::new()).unwrap();

    let written = presentation.to_xml().unwrap();
    assert!(!String::from_utf8_lossy(&written).contains("sectionLst"));
    let reopened = CT_Presentation::from_xml(&written).unwrap();
    assert!(reopened.sections().is_empty());
}

#[test]
fn unchanged_text_bodies_keep_inherited_a_raw_bytes_and_unsafe_rewrite_fails() {
    use rpptx_oxml::comments::CommentList;

    let xml = br#"<q:cmLst xmlns:q="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="urn:f217:producer" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main"><q:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:11:12Z"><q:unknownAnchor/><q:txBody><d:bodyPr/><d:lstStyle/><d:p/><a:producer/></q:txBody></q:cm><q:cm id="{33333333-3333-3333-3333-333333333333}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-08-29T10:12:13Z"><q:unknownAnchor/><q:txBody><d:bodyPr/><d:lstStyle/><d:p/><a:second/></q:txBody></q:cm></q:cmLst>"#;
    let mut comments = CommentList::from_xml(xml).unwrap();
    comments.move_comment(0, 1).unwrap();
    let written = String::from_utf8(comments.to_xml().unwrap()).unwrap();
    assert!(written.contains("<q:txBody><d:bodyPr/><d:lstStyle/><d:p/><a:producer/></q:txBody>"));
    assert!(written.contains("<q:txBody><d:bodyPr/><d:lstStyle/><d:p/><a:second/></q:txBody>"));
    CommentList::from_xml(written.as_bytes()).unwrap();

    let mut dirty = CommentList::from_xml(xml).unwrap();
    dirty.comments[0].set_text("changed");
    assert_eq!(
        dirty.to_xml().unwrap_err().to_string(),
        "invalid value: cannot rewrite comment text body with producer-owned a descendants"
    );
    assert!(CommentList::from_xml(xml).unwrap().to_xml().is_ok());
}
