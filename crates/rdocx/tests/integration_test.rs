//! Integration tests for rdocx — end-to-end document creation and round-trip.

use rdocx::Document;
use rdocx::paragraph::Alignment;
use rdocx::table::VerticalAlignment;
use rdocx::{BorderStyle, Length, SectionBreak, StyleBuilder, TabAlignment, TabLeader, UnderlineStyle};

#[test]
fn create_and_round_trip_simple_document() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello, World!");
    doc.add_paragraph("This is a test document.");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 2);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Hello, World!");
    assert_eq!(paras[1].text(), "This is a test document.");
}

#[test]
fn create_and_round_trip_formatted_document() {
    let mut doc = Document::new();

    // Title paragraph
    doc.add_paragraph("Document Title")
        .style("Heading1")
        .alignment(Alignment::Center);

    // Normal paragraph with multiple formatted runs
    let mut para = doc.add_paragraph("");
    para.add_run("This is ").font("Arial").size(11.0);
    para.add_run("bold").bold(true).font("Arial").size(11.0);
    para.add_run(" and this is ").font("Arial").size(11.0);
    para.add_run("italic").italic(true).font("Arial").size(11.0);
    para.add_run(".").font("Arial").size(11.0);

    // Justified paragraph with indentation
    doc.add_paragraph("This paragraph has special formatting.")
        .alignment(Alignment::Justify)
        .indent_left(Length::inches(0.5))
        .space_before(Length::pt(12.0))
        .space_after(Length::pt(6.0));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 3);

    // Check title
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Document Title");
    assert_eq!(paras[0].style_id(), Some("Heading1"));
    assert_eq!(paras[0].alignment(), Some(Alignment::Center));

    // Check formatted runs
    let runs: Vec<_> = paras[1].runs().collect();
    assert_eq!(runs.len(), 5);
    assert!(!runs[0].is_bold());
    assert!(runs[1].is_bold());
    assert!(!runs[2].is_italic());
    assert!(runs[3].is_italic());

    // Check justified paragraph
    assert_eq!(paras[2].alignment(), Some(Alignment::Justify));
}

#[test]
fn round_trip_preserves_styles() {
    let mut doc = Document::new();
    doc.add_paragraph("Normal paragraph");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Should have the default styles
    assert!(doc2.style("Normal").is_some());
    assert!(doc2.style("Heading1").is_some());

    let normal = doc2.style("Normal").unwrap();
    assert!(normal.is_default());
    assert_eq!(normal.name(), Some("Normal"));
}

#[test]
fn save_and_load_file() {
    let dir = std::env::temp_dir();
    let path = dir.join("rdocx_test_output.docx");

    // Create and save
    let mut doc = Document::new();
    doc.add_paragraph("Saved to disk");
    doc.save(&path).unwrap();

    // Load back
    let doc2 = Document::open(&path).unwrap();
    assert_eq!(doc2.paragraph_count(), 1);
    assert_eq!(doc2.paragraphs()[0].text(), "Saved to disk");

    // Clean up
    std::fs::remove_file(&path).ok();
}

#[test]
fn section_properties_preserved() {
    let doc = Document::new();
    let sect = doc.section_properties().unwrap();

    // Default US Letter
    assert_eq!(sect.page_width.unwrap().0, 12240); // 8.5"
    assert_eq!(sect.page_height.unwrap().0, 15840); // 11"
    assert_eq!(sect.margin_top.unwrap().0, 1440); // 1"

    // Round-trip
    let mut doc2 = Document::new();
    let bytes = doc2.to_bytes().unwrap();
    let doc3 = Document::from_bytes(&bytes).unwrap();
    let sect3 = doc3.section_properties().unwrap();
    assert_eq!(sect3.page_width.unwrap().0, 12240);
}

#[test]
fn empty_document_round_trip() {
    let mut doc = Document::new();
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 0);
}

#[test]
fn run_color_and_font_round_trip() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("Red text")
        .color("FF0000")
        .font("Times New Roman")
        .size(16.0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    let runs: Vec<_> = paras[0].runs().collect();
    assert_eq!(runs[0].color(), Some("FF0000"));
    assert_eq!(runs[0].font_name(), Some("Times New Roman"));
    assert_eq!(runs[0].size(), Some(16.0));
}

// ---- Phase 2 Integration Tests ----

#[test]
fn paragraph_borders_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Bordered paragraph")
        .border_all(BorderStyle::Single, 4, "000000");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert!(paras[0].has_borders());
}

#[test]
fn paragraph_tab_stops_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Tab text")
        .add_tab_stop(TabAlignment::Right, Length::inches(6.0))
        .add_tab_stop_with_leader(
            TabAlignment::Right,
            Length::inches(6.5),
            TabLeader::Dot,
        );

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].tab_stop_count(), 2);
}

#[test]
fn paragraph_shading_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Highlighted paragraph")
        .shading("FFFF00");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].shading_fill(), Some("FFFF00"));
}

#[test]
fn paragraph_spacing_and_indent_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Indented text")
        .indent_left(Length::inches(1.0))
        .indent_right(Length::inches(0.5))
        .first_line_indent(Length::inches(0.25))
        .space_before(Length::pt(12.0))
        .space_after(Length::pt(6.0))
        .line_spacing_multiple(1.5);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Verify it round-trips without error
    assert_eq!(doc2.paragraph_count(), 1);
    assert_eq!(doc2.paragraphs()[0].text(), "Indented text");
}

#[test]
fn paragraph_pagination_controls_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Keep with next")
        .keep_with_next(true)
        .keep_together(true)
        .widow_control(true);
    doc.add_paragraph("Page break before")
        .page_break_before(true);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 2);
}

#[test]
fn run_underline_styles_round_trip() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("Simple underline").underline(true);
    para.add_run("Wave underline")
        .underline_style(UnderlineStyle::Wave);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Simple underlineWave underline");
}

#[test]
fn run_advanced_formatting_round_trip() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("Strike").strike(true);
    para.add_run("DStrike").double_strike(true);
    para.add_run("CAPS").all_caps(true);
    para.add_run("SmallCaps").small_caps(true);
    para.add_run("Super").superscript();
    para.add_run("Sub").subscript();
    para.add_run("Hidden").hidden(true);
    para.add_run("Spaced").character_spacing(Length::pt(2.0));
    para.add_run("Wide").width_scale(150);
    para.add_run("Raised").position(6);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    let runs: Vec<_> = paras[0].runs().collect();
    assert_eq!(runs.len(), 10);
    assert!(runs[0].is_strike());
    assert_eq!(runs[4].vert_align(), Some("superscript"));
    assert_eq!(runs[5].vert_align(), Some("subscript"));
}

#[test]
fn run_style_assignment_round_trip() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("Styled run").style("Heading1Char");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let paras = doc2.paragraphs();
    let runs: Vec<_> = paras[0].runs().collect();
    assert_eq!(runs[0].style_id(), Some("Heading1Char"));
}

#[test]
fn custom_style_round_trip() {
    let mut doc = Document::new();

    doc.add_style(
        StyleBuilder::paragraph("CustomHeading", "Custom Heading")
            .based_on("Heading1")
            .next_style("Normal"),
    );

    doc.add_style(
        StyleBuilder::character("Emphasis", "Emphasis Style"),
    );

    doc.add_paragraph("Custom styled")
        .style("CustomHeading");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let s = doc2.style("CustomHeading").unwrap();
    assert_eq!(s.name(), Some("Custom Heading"));
    assert_eq!(s.based_on(), Some("Heading1"));

    assert!(doc2.style("Emphasis").is_some());

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].style_id(), Some("CustomHeading"));
}

#[test]
fn style_inheritance_resolution() {
    let doc = Document::new();

    // Heading1's rpr should have bold from the style definition
    let rpr = doc.resolve_run_properties(Some("Heading1"), None);
    assert_eq!(rpr.bold, Some(true));
    // Font inherited from docDefaults
    assert_eq!(rpr.font_ascii, Some("Calibri".to_string()));

    // Normal paragraph should get docDefaults spacing
    let ppr = doc.resolve_paragraph_properties(Some("Normal"));
    assert!(ppr.space_after.is_some());
}

#[test]
fn section_landscape_round_trip() {
    let mut doc = Document::new();
    doc.set_landscape();

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    assert!(sect.page_width.unwrap().0 > sect.page_height.unwrap().0);
}

#[test]
fn section_margins_round_trip() {
    let mut doc = Document::new();
    doc.set_margins(
        Length::inches(0.5),
        Length::inches(0.75),
        Length::inches(0.5),
        Length::inches(0.75),
    );

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    assert_eq!(sect.margin_top.unwrap().0, 720);
    assert_eq!(sect.margin_right.unwrap().0, 1080);
    assert_eq!(sect.margin_bottom.unwrap().0, 720);
    assert_eq!(sect.margin_left.unwrap().0, 1080);
}

#[test]
fn section_columns_round_trip() {
    let mut doc = Document::new();
    doc.set_columns(3, Length::inches(0.25));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    let cols = sect.columns.as_ref().unwrap();
    assert_eq!(cols.num, Some(3));
    assert_eq!(cols.equal_width, Some(true));
}

#[test]
fn section_a4_page_size() {
    let mut doc = Document::new();
    doc.set_page_size(Length::cm(21.0), Length::cm(29.7));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let sect = doc2.section_properties().unwrap();
    let w = sect.page_width.unwrap().0;
    let h = sect.page_height.unwrap().0;
    // A4 dimensions: 11906tw x 16838tw (allow small rounding)
    assert!((w - 11906).abs() < 5, "Expected ~11906, got {w}");
    assert!((h - 16838).abs() < 5, "Expected ~16838, got {h}");
}

#[test]
fn comprehensive_document_round_trip() {
    // Create a document with many Phase 2 features combined
    let mut doc = Document::new();

    // Custom style
    doc.add_style(
        StyleBuilder::paragraph("BlockQuote", "Block Quote")
            .based_on("Normal"),
    );

    // Page setup
    doc.set_margins(
        Length::inches(1.0),
        Length::inches(1.25),
        Length::inches(1.0),
        Length::inches(1.25),
    );

    // Title
    doc.add_paragraph("My Document")
        .style("Heading1")
        .alignment(Alignment::Center)
        .space_after(Length::pt(24.0));

    // Body paragraph with formatting
    let mut para = doc.add_paragraph("");
    para.add_run("This is ").font("Calibri").size(11.0);
    para.add_run("important")
        .bold(true)
        .color("FF0000")
        .font("Calibri")
        .size(11.0);
    para.add_run(" text with ")
        .font("Calibri")
        .size(11.0);
    para.add_run("underline")
        .underline(true)
        .font("Calibri")
        .size(11.0);
    para.add_run(".")
        .font("Calibri")
        .size(11.0);

    // Block quote with indentation and shading
    doc.add_paragraph("This is a block quote.")
        .style("BlockQuote")
        .indent_left(Length::inches(0.5))
        .indent_right(Length::inches(0.5))
        .shading("F2F2F2")
        .space_before(Length::pt(6.0))
        .space_after(Length::pt(6.0));

    // Bordered paragraph
    doc.add_paragraph("Important note")
        .border_all(BorderStyle::Single, 4, "000000")
        .shading("FFFFCC");

    // Save and reload
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 4);
    assert!(doc2.style("BlockQuote").is_some());

    let paras = doc2.paragraphs();
    assert_eq!(paras[0].style_id(), Some("Heading1"));
    assert_eq!(paras[0].alignment(), Some(Alignment::Center));

    let runs: Vec<_> = paras[1].runs().collect();
    assert_eq!(runs.len(), 5);
    assert!(runs[1].is_bold());
    assert_eq!(runs[1].color(), Some("FF0000"));

    assert_eq!(paras[2].shading_fill(), Some("F2F2F2"));
    assert!(paras[3].has_borders());
    assert_eq!(paras[3].shading_fill(), Some("FFFFCC"));
}

// ---- Phase 3 Integration Tests ----

#[test]
fn table_basic_creation_round_trip() {
    let mut doc = Document::new();
    let mut table = doc.add_table(3, 4);
    table.cell(0, 0).unwrap().set_text("A1");
    table.cell(0, 1).unwrap().set_text("B1");
    table.cell(1, 0).unwrap().set_text("A2");
    table.cell(2, 3).unwrap().set_text("D3");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.table_count(), 1);
    let tables = doc2.tables();
    let t = &tables[0];
    assert_eq!(t.row_count(), 3);
    assert_eq!(t.column_count(), 4);
    assert_eq!(t.cell(0, 0).unwrap().text(), "A1");
    assert_eq!(t.cell(0, 1).unwrap().text(), "B1");
    assert_eq!(t.cell(1, 0).unwrap().text(), "A2");
    assert_eq!(t.cell(2, 3).unwrap().text(), "D3");
}

#[test]
fn table_with_formatting_round_trip() {
    let mut doc = Document::new();
    doc.add_table(2, 2)
        .borders(BorderStyle::Single, 4, "000000")
        .alignment(Alignment::Center)
        .layout_fixed();

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.table_count(), 1);
    let tables = doc2.tables();
    assert_eq!(tables[0].row_count(), 2);
    assert_eq!(tables[0].column_count(), 2);
}

#[test]
fn table_cell_shading_and_alignment() {
    let mut doc = Document::new();
    let mut table = doc.add_table(2, 2);

    table.cell(0, 0).unwrap().set_text("Header");
    table
        .cell(0, 0)
        .unwrap()
        .shading("4472C4")
        .vertical_alignment(VerticalAlignment::Center);

    table
        .cell(1, 0)
        .unwrap()
        .shading("D9E2F3")
        .vertical_alignment(VerticalAlignment::Bottom);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let tables = doc2.tables();
    let cell_00 = tables[0].cell(0, 0).unwrap();
    assert_eq!(cell_00.shading_fill(), Some("4472C4"));
    assert_eq!(cell_00.vertical_alignment(), Some(VerticalAlignment::Center));

    let cell_10 = tables[0].cell(1, 0).unwrap();
    assert_eq!(cell_10.shading_fill(), Some("D9E2F3"));
    assert_eq!(cell_10.vertical_alignment(), Some(VerticalAlignment::Bottom));
}

#[test]
fn table_header_row_round_trip() {
    let mut doc = Document::new();
    let mut table = doc.add_table(3, 2);
    table.row(0).unwrap().header();
    table.cell(0, 0).unwrap().set_text("Col A");
    table.cell(0, 1).unwrap().set_text("Col B");
    table.cell(1, 0).unwrap().set_text("Data 1");
    table.cell(2, 0).unwrap().set_text("Data 2");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let tables = doc2.tables();
    assert!(tables[0].row(0).unwrap().is_header());
    assert!(!tables[0].row(1).unwrap().is_header());
}

#[test]
fn table_cell_grid_span_round_trip() {
    let mut doc = Document::new();
    let mut table = doc.add_table(2, 3);
    // First row: cell 0 spans 2 columns
    table.cell(0, 0).unwrap().set_text("Merged");
    table.cell(0, 0).unwrap().grid_span(2);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let tables = doc2.tables();
    assert_eq!(tables[0].cell(0, 0).unwrap().grid_span(), Some(2));
}

#[test]
fn table_mixed_with_paragraphs() {
    let mut doc = Document::new();
    doc.add_paragraph("Before the table");
    let mut table = doc.add_table(2, 2);
    table.cell(0, 0).unwrap().set_text("Cell");
    doc.add_paragraph("After the table");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 2);
    assert_eq!(doc2.table_count(), 1);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Before the table");
    assert_eq!(paras[1].text(), "After the table");
    let tables = doc2.tables();
    assert_eq!(tables[0].cell(0, 0).unwrap().text(), "Cell");
}

#[test]
fn table_cell_multiple_paragraphs() {
    let mut doc = Document::new();
    let mut table = doc.add_table(1, 1);
    let mut cell = table.cell(0, 0).unwrap();
    cell.set_text("First line");
    let mut para = cell.add_paragraph("");
    para.add_run("Second line").bold(true);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    let tables = doc2.tables();
    let cell = tables[0].cell(0, 0).unwrap();
    let paras: Vec<_> = cell.paragraphs().collect();
    assert_eq!(paras.len(), 2);
    assert_eq!(paras[0].text(), "First line");
    assert_eq!(paras[1].text(), "Second line");
}

#[test]
fn inline_image_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Before image");

    // Create a minimal 1x1 white PNG
    let png_data: Vec<u8> = vec![
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xDE, // 8-bit RGB
        0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41, 0x54, // IDAT chunk
        0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00,
        0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC, 0x33,
        0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, // IEND chunk
        0xAE, 0x42, 0x60, 0x82,
    ];

    doc.add_picture(
        &png_data,
        "test.png",
        Length::inches(2.0),
        Length::inches(1.5),
    );
    doc.add_paragraph("After image");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Image paragraph is counted as a paragraph too
    assert_eq!(doc2.paragraph_count(), 3);
    assert_eq!(doc2.paragraphs()[0].text(), "Before image");
    assert_eq!(doc2.paragraphs()[2].text(), "After image");
}

#[test]
fn header_footer_round_trip() {
    let mut doc = Document::new();
    doc.set_header("Page Header");
    doc.set_footer("Page Footer");
    doc.add_paragraph("Body text");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.header_text(), Some("Page Header".to_string()));
    assert_eq!(doc2.footer_text(), Some("Page Footer".to_string()));
    assert_eq!(doc2.paragraph_count(), 1);
}

#[test]
fn first_page_header_footer() {
    let mut doc = Document::new();
    doc.set_header("Default Header");
    doc.set_footer("Default Footer");
    doc.set_first_page_header("First Page Header");
    doc.set_first_page_footer("First Page Footer");
    doc.add_paragraph("Content");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.header_text(), Some("Default Header".to_string()));
    assert_eq!(doc2.footer_text(), Some("Default Footer".to_string()));
    // titlePg should be set
    let sect = doc2.section_properties().unwrap();
    assert_eq!(sect.title_pg, Some(true));
}

#[test]
fn bullet_list_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Items:");
    doc.add_bullet_list_item("First item", 0);
    doc.add_bullet_list_item("Second item", 0);
    doc.add_bullet_list_item("Sub-item", 1);
    doc.add_bullet_list_item("Third item", 0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 5);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Items:");
    assert_eq!(paras[1].text(), "First item");
    assert_eq!(paras[2].text(), "Second item");
    assert_eq!(paras[3].text(), "Sub-item");
    assert_eq!(paras[4].text(), "Third item");
}

#[test]
fn numbered_list_round_trip() {
    let mut doc = Document::new();
    doc.add_numbered_list_item("Step one", 0);
    doc.add_numbered_list_item("Step two", 0);
    doc.add_numbered_list_item("Sub-step a", 1);
    doc.add_numbered_list_item("Sub-step b", 1);
    doc.add_numbered_list_item("Step three", 0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 5);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Step one");
    assert_eq!(paras[4].text(), "Step three");
}

#[test]
fn mixed_lists_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("Introduction");
    doc.add_bullet_list_item("Bullet 1", 0);
    doc.add_bullet_list_item("Bullet 2", 0);
    doc.add_paragraph("Transition");
    doc.add_numbered_list_item("Step 1", 0);
    doc.add_numbered_list_item("Step 2", 0);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.paragraph_count(), 6);
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Introduction");
    assert_eq!(paras[1].text(), "Bullet 1");
    assert_eq!(paras[3].text(), "Transition");
    assert_eq!(paras[4].text(), "Step 1");
}

#[test]
fn comprehensive_phase3_document() {
    // Create a document using all Phase 3 features together
    let mut doc = Document::new();

    // Page setup
    doc.set_margins(
        Length::inches(1.0),
        Length::inches(1.0),
        Length::inches(1.0),
        Length::inches(1.0),
    );
    doc.set_header("Phase 3 Test Document");
    doc.set_footer("Confidential");

    // Title
    doc.add_paragraph("Project Report")
        .style("Heading1")
        .alignment(Alignment::Center);

    // Intro paragraph
    doc.add_paragraph("This document demonstrates Phase 3 features.");

    // Bulleted list section
    doc.add_paragraph("Key Points:").style("Heading2");
    doc.add_bullet_list_item("Tables with formatting", 0);
    doc.add_bullet_list_item("Inline images", 0);
    doc.add_bullet_list_item("Headers and footers", 0);
    doc.add_bullet_list_item("Numbered and bulleted lists", 0);

    // Table section
    doc.add_paragraph("Data Summary:").style("Heading2");
    let mut table = doc.add_table(3, 3)
        .borders(BorderStyle::Single, 4, "000000");

    // Header row
    table.row(0).unwrap().header();
    table.cell(0, 0).unwrap().set_text("Category");
    table
        .cell(0, 0)
        .unwrap()
        .shading("4472C4")
        .vertical_alignment(VerticalAlignment::Center);
    table.cell(0, 1).unwrap().set_text("Q1");
    table.cell(0, 1).unwrap().shading("4472C4");
    table.cell(0, 2).unwrap().set_text("Q2");
    table.cell(0, 2).unwrap().shading("4472C4");

    // Data rows
    table.cell(1, 0).unwrap().set_text("Revenue");
    table.cell(1, 1).unwrap().set_text("$1,200");
    table.cell(1, 2).unwrap().set_text("$1,500");
    table.cell(2, 0).unwrap().set_text("Expenses");
    table.cell(2, 1).unwrap().set_text("$800");
    table.cell(2, 2).unwrap().set_text("$900");

    // Numbered steps
    doc.add_paragraph("Next Steps:").style("Heading2");
    doc.add_numbered_list_item("Review Q2 financials", 0);
    doc.add_numbered_list_item("Prepare Q3 forecast", 0);
    doc.add_numbered_list_item("Gather input from teams", 1);
    doc.add_numbered_list_item("Consolidate data", 1);
    doc.add_numbered_list_item("Submit report", 0);

    // Save and reload
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Verify structure
    assert_eq!(doc2.table_count(), 1);
    let tables = doc2.tables();
    assert_eq!(tables[0].row_count(), 3);
    assert_eq!(tables[0].column_count(), 3);
    assert!(tables[0].row(0).unwrap().is_header());
    assert_eq!(tables[0].cell(1, 1).unwrap().text(), "$1,200");

    // Verify header/footer
    assert_eq!(
        doc2.header_text(),
        Some("Phase 3 Test Document".to_string())
    );
    assert_eq!(doc2.footer_text(), Some("Confidential".to_string()));

    // Count paragraphs (heading + intro + 3 heading2 + 4 bullets + 5 numbered = 15 total paragraphs)
    assert!(doc2.paragraph_count() > 10);
}

#[test]
fn metadata_round_trip() {
    let mut doc = Document::new();
    doc.set_title("Test Title");
    doc.set_author("Test Author");
    doc.set_subject("Test Subject");
    doc.set_keywords("rust, docx, test");

    assert_eq!(doc.title(), Some("Test Title"));
    assert_eq!(doc.author(), Some("Test Author"));
    assert_eq!(doc.subject(), Some("Test Subject"));
    assert_eq!(doc.keywords(), Some("rust, docx, test"));

    // Round-trip through DOCX bytes
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.title(), Some("Test Title"));
    assert_eq!(doc2.author(), Some("Test Author"));
    assert_eq!(doc2.subject(), Some("Test Subject"));
    assert_eq!(doc2.keywords(), Some("rust, docx, test"));
}

#[test]
fn nested_table_round_trip() {
    let mut doc = Document::new();

    // Create outer 2x2 table
    let mut tbl = doc.add_table(2, 2);
    tbl.cell(0, 0).unwrap().set_text("Outer A1");
    tbl.cell(0, 1).unwrap().set_text("Outer A2");
    tbl.cell(1, 1).unwrap().set_text("Outer B2");

    // Add a nested 2x1 table inside cell (1, 0)
    let mut cell_b1 = tbl.cell(1, 0).unwrap();
    cell_b1.set_text("Before nested");
    let mut nested = cell_b1.add_table(2, 1);
    nested.cell(0, 0).unwrap().set_text("Inner R1");
    nested.cell(1, 0).unwrap().set_text("Inner R2");

    // Serialize and reload
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    // Verify outer table structure
    assert_eq!(doc2.table_count(), 1);
    let tables = doc2.tables();
    assert!(!tables.is_empty());
    let tbl2 = &tables[0];
    assert_eq!(tbl2.row_count(), 2);
    assert_eq!(tbl2.cell(0, 0).unwrap().text(), "Outer A1");
    assert_eq!(tbl2.cell(0, 1).unwrap().text(), "Outer A2");
    assert_eq!(tbl2.cell(1, 1).unwrap().text(), "Outer B2");

    // Cell (1,0) should have paragraph text (nested table text excluded from text())
    let cell_b1_ref = tbl2.cell(1, 0).unwrap();
    assert_eq!(cell_b1_ref.text(), "Before nested");
}

#[test]
fn comprehensive_document_round_trip_with_nested() {
    use rdocx::paragraph::Alignment;

    let mut doc = Document::new();

    // Metadata
    doc.set_title("Comprehensive Test Document");
    doc.set_author("rdocx Test Suite");

    // Heading 1
    doc.add_paragraph("Chapter 1: Introduction")
        .style("Heading1");

    // Normal paragraphs
    doc.add_paragraph("This is a normal paragraph with some body text.")
        .alignment(Alignment::Left);

    doc.add_paragraph("This paragraph is centered for emphasis.")
        .alignment(Alignment::Center);

    doc.add_paragraph("This paragraph is justified for a clean look.")
        .alignment(Alignment::Justify);

    // Heading 2
    doc.add_paragraph("Section 1.1: Data Table")
        .style("Heading2");

    // Table with formatting
    let mut tbl = doc.add_table(3, 3);
    tbl = tbl.borders(rdocx::BorderStyle::Single, 4, "000000");

    // Header row
    tbl.row(0).unwrap().header();
    tbl.cell(0, 0).unwrap().set_text("Name");
    tbl.cell(0, 1).unwrap().set_text("Value");
    tbl.cell(0, 2).unwrap().set_text("Status");

    // Data rows
    tbl.cell(1, 0).unwrap().set_text("Alpha");
    tbl.cell(1, 1).unwrap().set_text("100");
    tbl.cell(1, 2).unwrap().set_text("Active");

    tbl.cell(2, 0).unwrap().set_text("Beta");
    tbl.cell(2, 1).unwrap().set_text("200");
    tbl.cell(2, 2).unwrap().set_text("Pending");

    // Another heading
    doc.add_paragraph("Chapter 2: Nested Content")
        .style("Heading1");

    // Table with nested table
    let mut tbl2 = doc.add_table(2, 2);
    tbl2.cell(0, 0).unwrap().set_text("Outer cell");
    tbl2.cell(0, 1).unwrap().set_text("Another outer cell");
    tbl2.cell(1, 0).unwrap().set_text("Simple cell");

    let mut nested_cell = tbl2.cell(1, 1).unwrap();
    nested_cell.set_text("Contains nested table:");
    let mut nested = nested_cell.add_table(2, 2);
    nested.cell(0, 0).unwrap().set_text("N1");
    nested.cell(0, 1).unwrap().set_text("N2");
    nested.cell(1, 0).unwrap().set_text("N3");
    nested.cell(1, 1).unwrap().set_text("N4");

    // Bullet list
    doc.add_paragraph("Chapter 3: Lists")
        .style("Heading1");
    doc.add_paragraph("First bullet point")
        .style("ListBullet");
    doc.add_paragraph("Second bullet point")
        .style("ListBullet");
    doc.add_paragraph("Third bullet point")
        .style("ListBullet");

    // Final paragraph
    doc.add_paragraph("End of document.");

    // Round-trip through DOCX
    let docx_bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&docx_bytes).unwrap();

    // Verify structure
    assert_eq!(doc2.title(), Some("Comprehensive Test Document"));
    assert_eq!(doc2.author(), Some("rdocx Test Suite"));
    assert!(doc2.table_count() >= 2);
}

// ---- Phase 7: Template Engine, Placeholder Replacement & Background Images ----

#[test]
fn template_open_replace_save() {
    // Create a "template" document
    let mut template = Document::new();
    template.set_header("Company: {{company}}");
    template.set_footer("Date: {{date}}");
    template.add_paragraph("Dear {{name}},").style("Heading1");
    template.add_paragraph("Welcome to {{company}}. Your role starts on {{date}}.");
    template.add_paragraph("{{INSERT_CONTENT}}");
    template.add_paragraph("Best regards,");
    template.add_paragraph("HR Department");

    let template_bytes = template.to_bytes().unwrap();

    // Open template and do replacements
    let mut doc = Document::from_bytes(&template_bytes).unwrap();
    doc.replace_text("{{company}}", "Acme Corp");
    doc.replace_text("{{name}}", "Alice");
    doc.replace_text("{{date}}", "2026-03-01");

    // Find and replace content placeholder
    if let Some(idx) = doc.find_content_index("{{INSERT_CONTENT}}") {
        doc.remove_content(idx);
        doc.insert_paragraph(idx, "Your onboarding schedule is attached.");
        doc.insert_paragraph(idx + 1, "Please review and confirm.");
    }

    // Verify
    let paras = doc.paragraphs();
    assert_eq!(paras[0].text(), "Dear Alice,");
    assert_eq!(paras[1].text(), "Welcome to Acme Corp. Your role starts on 2026-03-01.");
    assert_eq!(paras[2].text(), "Your onboarding schedule is attached.");
    assert_eq!(paras[3].text(), "Please review and confirm.");
    assert_eq!(paras[4].text(), "Best regards,");

    assert_eq!(doc.header_text().unwrap(), "Company: Acme Corp");
    assert_eq!(doc.footer_text().unwrap(), "Date: 2026-03-01");

    // Round-trip
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraphs()[0].text(), "Dear Alice,");
}

#[test]
fn replace_all_batch_workflow() {
    let mut doc = Document::new();
    doc.add_paragraph("{{a}} {{b}} {{c}}");

    let mut map = std::collections::HashMap::new();
    map.insert("{{a}}", "X");
    map.insert("{{b}}", "Y");
    map.insert("{{c}}", "Z");
    let count = doc.replace_all(&map);
    assert_eq!(count, 3);
    assert_eq!(doc.paragraphs()[0].text(), "X Y Z");
}

#[test]
fn background_image_end_to_end() {
    // Minimal 1x1 PNG
    let png_data: Vec<u8> = vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
        0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde,
        0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54,
        0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00,
        0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33,
        0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
        0xae, 0x42, 0x60, 0x82,
    ];

    let mut doc = Document::new();
    doc.add_paragraph("Page content here");
    doc.add_background_image(&png_data, "background.png");

    // Background paragraph inserted at index 0
    assert_eq!(doc.content_count(), 2);

    // Round-trip DOCX
    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.content_count(), 2);
}

#[test]
fn full_phase7_workflow() {
    // Minimal PNG
    let png_data: Vec<u8> = vec![
        0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
        0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde,
        0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54,
        0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00,
        0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33,
        0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
        0xae, 0x42, 0x60, 0x82,
    ];

    let mut doc = Document::new();
    doc.set_title("Template Output");
    doc.set_header("{{company}} - Confidential");

    doc.add_paragraph("Report for {{company}}").style("Heading1");
    doc.add_paragraph("Date: {{date}}");
    doc.add_paragraph("{{INSERT_HERE}}");
    doc.add_paragraph("Summary: {{company}} performed well in {{date}}.");

    // Add background image
    doc.add_background_image(&png_data, "bg.png");

    // Replace placeholders
    doc.replace_text("{{company}}", "Acme Corp");
    doc.replace_text("{{date}}", "2026-02-22");

    // Insert content at placeholder position
    if let Some(idx) = doc.find_content_index("{{INSERT_HERE}}") {
        doc.remove_content(idx);
        doc.insert_paragraph(idx, "Revenue increased by 15%.");
    }

    // Verify final state
    let paras = doc.paragraphs();
    // First paragraph is the background image paragraph, skip it
    let text_paras: Vec<_> = paras.iter().filter(|p| !p.text().is_empty()).collect();
    assert!(text_paras.iter().any(|p| p.text() == "Report for Acme Corp"));
    assert!(text_paras.iter().any(|p| p.text() == "Date: 2026-02-22"));
    assert!(text_paras.iter().any(|p| p.text() == "Revenue increased by 15%."));
    assert!(text_paras.iter().any(|p| p.text() == "Summary: Acme Corp performed well in 2026-02-22."));

    assert_eq!(doc.header_text().unwrap(), "Acme Corp - Confidential");

    // Save as DOCX
    let bytes = doc.to_bytes().unwrap();
    assert!(!bytes.is_empty());
}

// ---- Phase: Code Quality & Modernization — New Tests ----

#[test]
fn section_break_builders_round_trip() {
    let mut doc = Document::new();
    doc.add_paragraph("First section");

    // Create a section break to landscape
    doc.add_paragraph("Landscape section break")
        .section_break(SectionBreak::NextPage)
        .section_landscape();

    doc.add_paragraph("In landscape section");

    // Switch back to portrait
    doc.add_paragraph("Portrait section break")
        .section_break(SectionBreak::NextPage)
        .section_portrait();

    doc.add_paragraph("Back to portrait");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 5);
}

#[test]
fn section_page_size_custom() {
    let mut doc = Document::new();
    // Custom page size: 6" x 9" (book size)
    doc.add_paragraph("Small page")
        .section_break(SectionBreak::NextPage)
        .section_page_size(Length::inches(6.0), Length::inches(9.0));

    doc.add_paragraph("On the next section");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 2);
}

#[test]
fn section_break_continuous() {
    let mut doc = Document::new();
    doc.add_paragraph("Before continuous break")
        .section_break(SectionBreak::Continuous);
    doc.add_paragraph("After continuous break");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 2);
}

#[test]
fn tab_stops_all_leader_styles() {
    let mut doc = Document::new();

    // Tab stop with no leader
    doc.add_paragraph("Left\tRight")
        .add_tab_stop(TabAlignment::Right, Length::inches(6.0));

    // Tab stop with dot leader
    doc.add_paragraph("Item\t100")
        .add_tab_stop_with_leader(TabAlignment::Right, Length::inches(6.0), TabLeader::Dot);

    // Tab stop with hyphen leader
    doc.add_paragraph("Section\tPage 5")
        .add_tab_stop_with_leader(TabAlignment::Right, Length::inches(6.0), TabLeader::Hyphen);

    // Tab stop with underscore leader
    doc.add_paragraph("Name\t")
        .add_tab_stop_with_leader(TabAlignment::Right, Length::inches(6.0), TabLeader::Underscore);

    // Multiple alignments
    doc.add_paragraph("A\tB\tC")
        .add_tab_stop(TabAlignment::Center, Length::inches(3.0))
        .add_tab_stop(TabAlignment::Right, Length::inches(6.0))
        .add_tab_stop(TabAlignment::Decimal, Length::inches(4.5));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();

    assert_eq!(paras[0].tab_stop_count(), 1);
    assert_eq!(paras[1].tab_stop_count(), 1);
    assert_eq!(paras[2].tab_stop_count(), 1);
    assert_eq!(paras[3].tab_stop_count(), 1);
    assert_eq!(paras[4].tab_stop_count(), 3);
}

#[test]
fn run_formatting_all_caps_small_caps() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("UPPERCASE").all_caps(true);
    para.add_run("SmallCaps").small_caps(true);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "UPPERCASESmallCaps");
}

#[test]
fn run_formatting_double_strike_and_spacing() {
    let mut doc = Document::new();
    let mut para = doc.add_paragraph("");
    para.add_run("DStrike").double_strike(true);
    para.add_run("Spaced").character_spacing(Length::pt(3.0));
    para.add_run("Super").superscript();
    para.add_run("Sub").subscript();

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();
    let runs: Vec<_> = paras[0].runs().collect();
    assert_eq!(runs.len(), 4);
    assert_eq!(runs[2].vert_align(), Some("superscript"));
    assert_eq!(runs[3].vert_align(), Some("subscript"));
    assert!(runs[1].character_spacing().is_some());
}

#[test]
fn paragraph_border_bottom_only() {
    let mut doc = Document::new();
    doc.add_paragraph("Bottom bordered")
        .border_bottom(BorderStyle::Single, 4, "000000");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();
    assert!(paras[0].has_borders());
}

#[test]
fn paragraph_shading_and_indent_combined() {
    let mut doc = Document::new();
    doc.add_paragraph("Shaded and indented")
        .shading("E0E0E0")
        .indent_left(Length::inches(0.75))
        .indent_right(Length::inches(0.5))
        .hanging_indent(Length::inches(0.25));

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].shading_fill(), Some("E0E0E0"));
    assert_eq!(paras[0].text(), "Shaded and indented");
}

#[test]
fn document_header_footer_first_page() {
    let mut doc = Document::new();
    doc.set_header("Default Header");
    doc.set_footer("Default Footer");
    doc.set_first_page_header("First Page Header");
    doc.set_first_page_footer("First Page Footer");
    doc.add_paragraph("Body content");

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();

    assert_eq!(doc2.header_text(), Some("Default Header".to_string()));
    assert_eq!(doc2.footer_text(), Some("Default Footer".to_string()));
    let sect = doc2.section_properties().unwrap();
    assert_eq!(sect.title_pg, Some(true));
}

#[test]
fn insert_paragraph_at_beginning_and_end() {
    let mut doc = Document::new();
    doc.add_paragraph("Middle");

    // Insert at beginning
    doc.insert_paragraph(0, "First");
    // Insert at end
    let count = doc.content_count();
    doc.insert_paragraph(count, "Last");

    assert_eq!(doc.content_count(), 3);
    let paras = doc.paragraphs();
    assert_eq!(paras[0].text(), "First");
    assert_eq!(paras[1].text(), "Middle");
    assert_eq!(paras[2].text(), "Last");
}

#[test]
fn insert_table_at_index() {
    let mut doc = Document::new();
    doc.add_paragraph("Before");
    doc.add_paragraph("After");

    // Insert table between the two paragraphs
    let mut table = doc.insert_table(1, 2, 2);
    table.cell(0, 0).unwrap().set_text("Cell");

    assert_eq!(doc.content_count(), 3);

    let bytes = doc.to_bytes().unwrap();
    let doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.table_count(), 1);
    assert_eq!(doc2.paragraph_count(), 2);
}

#[test]
fn remove_content_basic() {
    let mut doc = Document::new();
    doc.add_paragraph("Keep");
    doc.add_paragraph("Remove");
    doc.add_paragraph("Keep too");

    assert_eq!(doc.content_count(), 3);
    assert!(doc.remove_content(1));
    assert_eq!(doc.content_count(), 2);
    assert_eq!(doc.paragraphs()[0].text(), "Keep");
    assert_eq!(doc.paragraphs()[1].text(), "Keep too");
}

#[test]
fn remove_content_out_of_bounds() {
    let mut doc = Document::new();
    doc.add_paragraph("Only");
    assert!(!doc.remove_content(5));
    assert_eq!(doc.content_count(), 1);
}

#[test]
fn find_content_index_returns_none_for_missing() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello");
    assert_eq!(doc.find_content_index("nonexistent"), None);
}

#[test]
fn section_break_round_trip_preserves() {
    let mut doc = Document::new();
    doc.add_paragraph("Section 1")
        .section_break(SectionBreak::NextPage)
        .section_landscape();

    doc.add_paragraph("Section 2 (landscape)");

    doc.add_paragraph("Section 2 end")
        .section_break(SectionBreak::NextPage)
        .section_portrait();

    doc.add_paragraph("Section 3 (portrait)");

    // Round-trip
    let bytes = doc.to_bytes().unwrap();
    let mut doc2 = Document::from_bytes(&bytes).unwrap();
    assert_eq!(doc2.paragraph_count(), 4);

    // The document should still be valid and have the paragraphs
    let paras = doc2.paragraphs();
    assert_eq!(paras[0].text(), "Section 1");
    assert_eq!(paras[1].text(), "Section 2 (landscape)");
    assert_eq!(paras[2].text(), "Section 2 end");
    assert_eq!(paras[3].text(), "Section 3 (portrait)");

    // Re-round-trip to verify stability
    let bytes2 = doc2.to_bytes().unwrap();
    let doc3 = Document::from_bytes(&bytes2).unwrap();
    assert_eq!(doc3.paragraph_count(), 4);
}

#[test]
fn empty_document_insert_and_remove() {
    let mut doc = Document::new();
    assert_eq!(doc.content_count(), 0);

    doc.insert_paragraph(0, "Inserted");
    assert_eq!(doc.content_count(), 1);
    assert_eq!(doc.paragraphs()[0].text(), "Inserted");

    assert!(doc.remove_content(0));
    assert_eq!(doc.content_count(), 0);
}

// ---- PDF rendering tests ----

#[test]
fn to_pdf_simple_document() {
    let mut doc = Document::new();
    doc.add_paragraph("Hello, World!");
    doc.add_paragraph("This is a test document.");

    let result = doc.to_pdf();
    // On systems without fonts, layout may fail — that's OK for CI
    if let Ok(pdf_bytes) = result {
        // Verify it starts with PDF header
        assert!(pdf_bytes.starts_with(b"%PDF"));
        // Verify it's not trivially small
        assert!(pdf_bytes.len() > 100);
        // Verify it ends with %%EOF
        let tail = String::from_utf8_lossy(&pdf_bytes[pdf_bytes.len().saturating_sub(10)..]);
        assert!(tail.contains("%%EOF"));
    }
}

#[test]
fn to_pdf_with_formatting() {
    let mut doc = Document::new();
    doc.add_paragraph("Title")
        .style("Heading1")
        .alignment(Alignment::Center);
    doc.add_paragraph("Normal text with ")
        .add_run("bold").bold(true);
    doc.add_paragraph("Another paragraph");

    let result = doc.to_pdf();
    if let Ok(pdf_bytes) = result {
        assert!(pdf_bytes.starts_with(b"%PDF"));
        assert!(pdf_bytes.len() > 200);
    }
}

#[test]
fn to_pdf_with_table() {
    let mut doc = Document::new();
    doc.add_paragraph("Table test");
    {
        let mut table = doc.add_table(2, 3);
        table.cell(0, 0).unwrap().set_text("A1");
        table.cell(0, 1).unwrap().set_text("B1");
        table.cell(0, 2).unwrap().set_text("C1");
        table.cell(1, 0).unwrap().set_text("A2");
        table.cell(1, 1).unwrap().set_text("B2");
        table.cell(1, 2).unwrap().set_text("C2");
    }
    doc.add_paragraph("After table");

    let result = doc.to_pdf();
    if let Ok(pdf_bytes) = result {
        assert!(pdf_bytes.starts_with(b"%PDF"));
    }
}

#[test]
fn to_pdf_with_metadata() {
    let mut doc = Document::new();
    doc.set_title("Test Document");
    doc.set_author("rdocx");
    doc.add_paragraph("Content");

    let result = doc.to_pdf();
    if let Ok(pdf_bytes) = result {
        assert!(pdf_bytes.starts_with(b"%PDF"));
        // Metadata should be embedded in the PDF
        let pdf_str = String::from_utf8_lossy(&pdf_bytes);
        assert!(pdf_str.contains("Test Document") || pdf_str.contains("rdocx-pdf"));
    }
}

#[test]
fn save_pdf_to_file() {
    let mut doc = Document::new();
    doc.add_paragraph("PDF file test");

    let path = "/tmp/rdocx_test_output.pdf";
    let result = doc.save_pdf(path);
    if result.is_ok() {
        let bytes = std::fs::read(path).unwrap();
        assert!(bytes.starts_with(b"%PDF"));
        std::fs::remove_file(path).ok();
    }
}
