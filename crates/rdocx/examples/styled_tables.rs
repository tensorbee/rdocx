//! Styled tables showcase — demonstrates all table formatting options.
//!
//! Run with: cargo run --example styled_tables

use std::path::Path;

use rdocx::{BorderStyle, Document, Length, VerticalAlignment};

fn main() {
    let samples_dir = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("samples");
    std::fs::create_dir_all(&samples_dir).unwrap();

    let out = samples_dir.join("styled_tables.docx");
    generate_styled_tables(&out);
    println!("  Created: styled_tables.docx");
    println!("\nDone!");
}

fn generate_styled_tables(path: &Path) {
    let mut doc = Document::new();
    doc.set_page_size(Length::inches(8.5), Length::inches(11.0));
    doc.set_margins(
        Length::inches(0.75), Length::inches(0.75),
        Length::inches(0.75), Length::inches(0.75),
    );

    doc.add_paragraph("Styled Tables Showcase")
        .style("Heading1");

    doc.add_paragraph("");

    // =========================================================================
    // 1. Professional report table with alternating rows
    // =========================================================================
    doc.add_paragraph("1. Report Table with Alternating Row Colors")
        .style("Heading2");

    {
        let mut tbl = doc.add_table(8, 4);
        tbl = tbl.borders(BorderStyle::Single, 2, "BFBFBF");
        tbl = tbl.width_pct(100.0);

        // Header row
        let headers = ["Product", "Q1 Sales", "Q2 Sales", "Growth"];
        for (col, h) in headers.iter().enumerate() {
            tbl.cell(0, col).unwrap().shading("2E75B6");
            tbl.cell(0, col).unwrap().set_text(h);
        }
        tbl.row(0).unwrap().header();

        // Data with alternating shading
        let data = [
            ["Enterprise Suite", "$245,000", "$312,000", "+27.3%"],
            ["Professional",     "$189,000", "$201,000", "+6.3%"],
            ["Starter Pack",     "$67,000",  "$84,500",  "+26.1%"],
            ["Add-ons",          "$34,000",  "$41,200",  "+21.2%"],
            ["Training",         "$22,000",  "$28,900",  "+31.4%"],
            ["Support Plans",    "$56,000",  "$62,300",  "+11.3%"],
        ];

        for (i, row) in data.iter().enumerate() {
            let row_idx = i + 1;
            for (col, val) in row.iter().enumerate() {
                tbl.cell(row_idx, col).unwrap().set_text(val);
                // Alternate row colors
                if i % 2 == 0 {
                    tbl.cell(row_idx, col).unwrap().shading("F2F7FB");
                }
            }
        }

        // Total row
        tbl.cell(7, 0).unwrap().set_text("TOTAL");
        tbl.cell(7, 0).unwrap().shading("D6E4F0");
        tbl.cell(7, 1).unwrap().set_text("$613,000");
        tbl.cell(7, 1).unwrap().shading("D6E4F0");
        tbl.cell(7, 2).unwrap().set_text("$729,900");
        tbl.cell(7, 2).unwrap().shading("D6E4F0");
        tbl.cell(7, 3).unwrap().set_text("+19.1%");
        tbl.cell(7, 3).unwrap().shading("D6E4F0");
    }

    doc.add_paragraph("");

    // =========================================================================
    // 2. Invoice-style table with merged header
    // =========================================================================
    doc.add_paragraph("2. Invoice Table with Merged Header & Row Spans")
        .style("Heading2");

    {
        let mut tbl = doc.add_table(7, 4);
        tbl = tbl.borders(BorderStyle::Single, 4, "000000");
        tbl = tbl.width_pct(100.0);

        // Merged title row
        tbl.cell(0, 0).unwrap().set_text("INVOICE #2026-0042");
        tbl.cell(0, 0).unwrap().grid_span(4);
        tbl.cell(0, 0).unwrap().shading("1F4E79");

        // Column headers
        let headers = ["Item", "Description", "Qty", "Amount"];
        for (col, h) in headers.iter().enumerate() {
            tbl.cell(1, col).unwrap().set_text(h);
            tbl.cell(1, col).unwrap().shading("D6E4F0");
        }

        // Line items
        tbl.cell(2, 0).unwrap().set_text("LIC-ENT-500");
        tbl.cell(2, 1).unwrap().set_text("Enterprise License (500 seats)");
        tbl.cell(2, 2).unwrap().set_text("1");
        tbl.cell(2, 3).unwrap().set_text("$60,000");

        tbl.cell(3, 0).unwrap().set_text("SVC-IMPL");
        tbl.cell(3, 1).unwrap().set_text("Implementation Services");
        tbl.cell(3, 2).unwrap().set_text("1");
        tbl.cell(3, 3).unwrap().set_text("$25,000");

        tbl.cell(4, 0).unwrap().set_text("SVC-TRAIN");
        tbl.cell(4, 1).unwrap().set_text("On-site Training (3 days)");
        tbl.cell(4, 2).unwrap().set_text("1");
        tbl.cell(4, 3).unwrap().set_text("$4,500");

        // Subtotal
        tbl.cell(5, 0).unwrap().set_text("Subtotal");
        tbl.cell(5, 0).unwrap().grid_span(3);
        tbl.cell(5, 0).unwrap().shading("F2F2F2");
        tbl.cell(5, 3).unwrap().set_text("$89,500");
        tbl.cell(5, 3).unwrap().shading("F2F2F2");

        // Total
        tbl.cell(6, 0).unwrap().set_text("TOTAL DUE");
        tbl.cell(6, 0).unwrap().grid_span(3);
        tbl.cell(6, 0).unwrap().shading("1F4E79");
        tbl.cell(6, 3).unwrap().set_text("$89,500");
        tbl.cell(6, 3).unwrap().shading("1F4E79");
    }

    doc.add_paragraph("");

    // =========================================================================
    // 3. Specification table with vertical merge
    // =========================================================================
    doc.add_paragraph("3. Specification Table with Vertical Merges")
        .style("Heading2");

    {
        let mut tbl = doc.add_table(8, 3);
        tbl = tbl.borders(BorderStyle::Single, 4, "000000");
        tbl = tbl.width_pct(100.0);

        // Header
        tbl.cell(0, 0).unwrap().set_text("Category");
        tbl.cell(0, 0).unwrap().shading("2E75B6");
        tbl.cell(0, 1).unwrap().set_text("Specification");
        tbl.cell(0, 1).unwrap().shading("2E75B6");
        tbl.cell(0, 2).unwrap().set_text("Value");
        tbl.cell(0, 2).unwrap().shading("2E75B6");

        // "Hardware" spans 3 rows
        tbl.cell(1, 0).unwrap().set_text("Hardware");
        tbl.cell(1, 0).unwrap().v_merge_restart();
        tbl.cell(1, 0).unwrap().shading("E2EFDA");
        tbl.cell(1, 0).unwrap().vertical_alignment(VerticalAlignment::Center);
        tbl.cell(1, 1).unwrap().set_text("Processor");
        tbl.cell(1, 2).unwrap().set_text("Intel Xeon E-2388G");

        tbl.cell(2, 0).unwrap().v_merge_continue();
        tbl.cell(2, 1).unwrap().set_text("Memory");
        tbl.cell(2, 2).unwrap().set_text("64 GB DDR4 ECC");

        tbl.cell(3, 0).unwrap().v_merge_continue();
        tbl.cell(3, 1).unwrap().set_text("Storage");
        tbl.cell(3, 2).unwrap().set_text("2x 1TB NVMe SSD (RAID 1)");

        // "Network" spans 2 rows
        tbl.cell(4, 0).unwrap().set_text("Network");
        tbl.cell(4, 0).unwrap().v_merge_restart();
        tbl.cell(4, 0).unwrap().shading("FCE4D6");
        tbl.cell(4, 0).unwrap().vertical_alignment(VerticalAlignment::Center);
        tbl.cell(4, 1).unwrap().set_text("Ethernet");
        tbl.cell(4, 2).unwrap().set_text("4x 10GbE SFP+");

        tbl.cell(5, 0).unwrap().v_merge_continue();
        tbl.cell(5, 1).unwrap().set_text("Management");
        tbl.cell(5, 2).unwrap().set_text("1x 1GbE IPMI");

        // "Software" spans 2 rows
        tbl.cell(6, 0).unwrap().set_text("Software");
        tbl.cell(6, 0).unwrap().v_merge_restart();
        tbl.cell(6, 0).unwrap().shading("D6E4F0");
        tbl.cell(6, 0).unwrap().vertical_alignment(VerticalAlignment::Center);
        tbl.cell(6, 1).unwrap().set_text("Operating System");
        tbl.cell(6, 2).unwrap().set_text("Ubuntu 24.04 LTS");

        tbl.cell(7, 0).unwrap().v_merge_continue();
        tbl.cell(7, 1).unwrap().set_text("Monitoring");
        tbl.cell(7, 2).unwrap().set_text("Prometheus + Grafana");
    }

    doc.add_paragraph("");

    // =========================================================================
    // 4. Nested table (table inside a cell)
    // =========================================================================
    doc.add_paragraph("4. Nested Table")
        .style("Heading2");

    {
        let mut tbl = doc.add_table(2, 2);
        tbl = tbl.borders(BorderStyle::Single, 6, "2E75B6");
        tbl = tbl.width_pct(100.0);
        tbl = tbl.cell_margins(
            Length::twips(72), Length::twips(108),
            Length::twips(72), Length::twips(108),
        );

        tbl.cell(0, 0).unwrap().set_text("Project Alpha");
        tbl.cell(0, 0).unwrap().shading("2E75B6");
        tbl.cell(0, 1).unwrap().set_text("Project Beta");
        tbl.cell(0, 1).unwrap().shading("2E75B6");

        // Nested table in cell (1,0)
        {
            let mut cell = tbl.cell(1, 0).unwrap();
            cell.set_text("Milestones:");
            let mut inner = cell.add_table(3, 2);
            inner = inner.borders(BorderStyle::Single, 2, "70AD47");
            inner.cell(0, 0).unwrap().set_text("Phase");
            inner.cell(0, 0).unwrap().shading("E2EFDA");
            inner.cell(0, 1).unwrap().set_text("Status");
            inner.cell(0, 1).unwrap().shading("E2EFDA");
            inner.cell(1, 0).unwrap().set_text("Design");
            inner.cell(1, 1).unwrap().set_text("Complete");
            inner.cell(2, 0).unwrap().set_text("Build");
            inner.cell(2, 1).unwrap().set_text("In Progress");
        }

        // Nested table in cell (1,1)
        {
            let mut cell = tbl.cell(1, 1).unwrap();
            cell.set_text("Budget:");
            let mut inner = cell.add_table(3, 2);
            inner = inner.borders(BorderStyle::Single, 2, "ED7D31");
            inner.cell(0, 0).unwrap().set_text("Category");
            inner.cell(0, 0).unwrap().shading("FCE4D6");
            inner.cell(0, 1).unwrap().set_text("Amount");
            inner.cell(0, 1).unwrap().shading("FCE4D6");
            inner.cell(1, 0).unwrap().set_text("Development");
            inner.cell(1, 1).unwrap().set_text("$120,000");
            inner.cell(2, 0).unwrap().set_text("Testing");
            inner.cell(2, 1).unwrap().set_text("$35,000");
        }
    }

    doc.add_paragraph("");

    // =========================================================================
    // 5. Form-style table with labels
    // =========================================================================
    doc.add_paragraph("5. Form-Style Table")
        .style("Heading2");

    {
        let mut tbl = doc.add_table(6, 4);
        tbl = tbl.borders(BorderStyle::Single, 4, "808080");
        tbl = tbl.width_pct(100.0);

        // Row 0: Full-width title
        tbl.cell(0, 0).unwrap().set_text("Customer Registration Form");
        tbl.cell(0, 0).unwrap().grid_span(4);
        tbl.cell(0, 0).unwrap().shading("404040");

        // Row 1: Name fields
        tbl.cell(1, 0).unwrap().set_text("First Name");
        tbl.cell(1, 0).unwrap().shading("E8E8E8");
        tbl.cell(1, 1).unwrap().set_text("John");
        tbl.cell(1, 2).unwrap().set_text("Last Name");
        tbl.cell(1, 2).unwrap().shading("E8E8E8");
        tbl.cell(1, 3).unwrap().set_text("Smith");

        // Row 2: Contact
        tbl.cell(2, 0).unwrap().set_text("Email");
        tbl.cell(2, 0).unwrap().shading("E8E8E8");
        tbl.cell(2, 1).unwrap().set_text("john.smith@example.com");
        tbl.cell(2, 1).unwrap().grid_span(3);

        // Row 3: Phone
        tbl.cell(3, 0).unwrap().set_text("Phone");
        tbl.cell(3, 0).unwrap().shading("E8E8E8");
        tbl.cell(3, 1).unwrap().set_text("+1 (555) 123-4567");
        tbl.cell(3, 2).unwrap().set_text("Company");
        tbl.cell(3, 2).unwrap().shading("E8E8E8");
        tbl.cell(3, 3).unwrap().set_text("Acme Corp");

        // Row 4: Address (spanning)
        tbl.cell(4, 0).unwrap().set_text("Address");
        tbl.cell(4, 0).unwrap().shading("E8E8E8");
        tbl.cell(4, 1).unwrap().set_text("123 Business Ave, Suite 400, Portland, OR 97201");
        tbl.cell(4, 1).unwrap().grid_span(3);

        // Row 5: Notes
        tbl.cell(5, 0).unwrap().set_text("Notes");
        tbl.cell(5, 0).unwrap().shading("E8E8E8");
        tbl.cell(5, 0).unwrap().vertical_alignment(VerticalAlignment::Top);
        {
            let mut cell = tbl.cell(5, 1).unwrap().grid_span(3);
            cell.set_text("Premium customer since 2020. Preferred contact method: email.");
            cell.add_paragraph("Annual review scheduled for March 2026.");
        }
    }

    doc.add_paragraph("");

    // =========================================================================
    // 6. Comparison table with border styles
    // =========================================================================
    doc.add_paragraph("6. Comparison Table with Custom Borders")
        .style("Heading2");

    {
        let mut tbl = doc.add_table(5, 3);
        tbl = tbl.borders(BorderStyle::Double, 4, "2E75B6");
        tbl = tbl.width_pct(100.0);

        // Header
        tbl.cell(0, 0).unwrap().set_text("Feature");
        tbl.cell(0, 0).unwrap().shading("2E75B6");
        tbl.cell(0, 1).unwrap().set_text("Basic Plan");
        tbl.cell(0, 1).unwrap().shading("2E75B6");
        tbl.cell(0, 2).unwrap().set_text("Enterprise Plan");
        tbl.cell(0, 2).unwrap().shading("2E75B6");

        tbl.cell(1, 0).unwrap().set_text("Users");
        tbl.cell(1, 1).unwrap().set_text("Up to 10");
        tbl.cell(1, 2).unwrap().set_text("Unlimited");
        tbl.cell(1, 2).unwrap().shading("E2EFDA");

        tbl.cell(2, 0).unwrap().set_text("Storage");
        tbl.cell(2, 1).unwrap().set_text("50 GB");
        tbl.cell(2, 2).unwrap().set_text("5 TB");
        tbl.cell(2, 2).unwrap().shading("E2EFDA");

        tbl.cell(3, 0).unwrap().set_text("Support");
        tbl.cell(3, 1).unwrap().set_text("Email only");
        tbl.cell(3, 2).unwrap().set_text("24/7 Phone + Email");
        tbl.cell(3, 2).unwrap().shading("E2EFDA");

        tbl.cell(4, 0).unwrap().set_text("Price");
        tbl.cell(4, 0).unwrap().shading("F2F2F2");
        tbl.cell(4, 1).unwrap().set_text("$29/month");
        tbl.cell(4, 1).unwrap().shading("F2F2F2");
        tbl.cell(4, 2).unwrap().set_text("$199/month");
        tbl.cell(4, 2).unwrap().shading("C6EFCE");
    }

    doc.add_paragraph("");

    // =========================================================================
    // 7. Wide table with fixed layout and row height
    // =========================================================================
    doc.add_paragraph("7. Fixed Layout Table with Row Height Control")
        .style("Heading2");

    {
        let mut tbl = doc.add_table(4, 5);
        tbl = tbl.borders(BorderStyle::Single, 4, "000000");
        tbl = tbl.width(Length::inches(7.0));
        tbl = tbl.layout_fixed();

        // Set column widths
        for col in 0..5 {
            tbl.cell(0, col).unwrap().width(Length::inches(1.4));
        }

        // Header with exact height
        tbl.row(0).unwrap().height_exact(Length::twips(480));
        tbl.row(0).unwrap().header();
        tbl.row(0).unwrap().cant_split();

        let headers = ["Mon", "Tue", "Wed", "Thu", "Fri"];
        for (col, h) in headers.iter().enumerate() {
            tbl.cell(0, col).unwrap().set_text(h);
            tbl.cell(0, col).unwrap().shading("404040");
            tbl.cell(0, col).unwrap().vertical_alignment(VerticalAlignment::Center);
        }

        // Schedule rows with minimum height
        tbl.row(1).unwrap().height(Length::twips(600));
        tbl.cell(1, 0).unwrap().set_text("9:00 Standup");
        tbl.cell(1, 1).unwrap().set_text("9:00 Standup");
        tbl.cell(1, 2).unwrap().set_text("9:00 Standup");
        tbl.cell(1, 3).unwrap().set_text("9:00 Standup");
        tbl.cell(1, 4).unwrap().set_text("9:00 Standup");

        tbl.row(2).unwrap().height(Length::twips(600));
        tbl.cell(2, 0).unwrap().set_text("10:00 Dev");
        tbl.cell(2, 0).unwrap().shading("D6E4F0");
        tbl.cell(2, 1).unwrap().set_text("10:00 Design Review");
        tbl.cell(2, 1).unwrap().shading("FCE4D6");
        tbl.cell(2, 2).unwrap().set_text("10:00 Dev");
        tbl.cell(2, 2).unwrap().shading("D6E4F0");
        tbl.cell(2, 3).unwrap().set_text("10:00 Sprint Planning");
        tbl.cell(2, 3).unwrap().shading("E2EFDA");
        tbl.cell(2, 4).unwrap().set_text("10:00 Dev");
        tbl.cell(2, 4).unwrap().shading("D6E4F0");

        tbl.row(3).unwrap().height(Length::twips(600));
        tbl.cell(3, 0).unwrap().set_text("14:00 Code Review");
        tbl.cell(3, 1).unwrap().set_text("14:00 Dev");
        tbl.cell(3, 1).unwrap().shading("D6E4F0");
        tbl.cell(3, 2).unwrap().set_text("14:00 Demo");
        tbl.cell(3, 2).unwrap().shading("FCE4D6");
        tbl.cell(3, 3).unwrap().set_text("14:00 Dev");
        tbl.cell(3, 3).unwrap().shading("D6E4F0");
        tbl.cell(3, 4).unwrap().set_text("14:00 Retro");
        tbl.cell(3, 4).unwrap().shading("E2EFDA");
    }

    doc.set_title("Styled Tables Showcase");
    doc.set_author("rdocx");

    doc.save(path).unwrap();
}
