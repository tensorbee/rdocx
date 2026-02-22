//! Template-based placeholder replacement workflow.
//!
//! This example demonstrates the typical template workflow:
//! 1. Create a template document with placeholders
//! 2. Open the template
//! 3. Replace placeholders in body text, tables, and across formatting runs
//! 4. Insert additional content at specific positions
//! 5. Save the final document
//!
//! Run with: cargo run --example template_replace

use std::collections::HashMap;
use std::path::Path;

use rdocx::{Alignment, BorderStyle, Document, Length};

fn main() {
    let samples_dir = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap()
        .join("samples");
    std::fs::create_dir_all(&samples_dir).unwrap();

    // Step 1: Create a template document with placeholders
    let template_path = samples_dir.join("_template.docx");
    create_template(&template_path);
    println!("  Created template: _template.docx");

    // Step 2: Open the template and replace placeholders
    let output_path = samples_dir.join("from_template.docx");
    fill_template(&template_path, &output_path);
    println!("  Created filled document: from_template.docx");

    // Clean up the intermediate template
    std::fs::remove_file(&template_path).ok();
    println!("\nDone!");
}

/// Create a template document with placeholders in body text, tables, and
/// cross-run formatting.
fn create_template(path: &Path) {
    let mut doc = Document::new();
    doc.set_page_size(Length::inches(8.5), Length::inches(11.0));
    doc.set_margins(
        Length::inches(1.0), Length::inches(1.0),
        Length::inches(1.0), Length::inches(1.0),
    );

    doc.set_header("{{company_name}} — Confidential");
    doc.set_footer("Prepared by {{author_name}} on {{date}}");

    // ── Title ──
    doc.add_paragraph("{{company_name}}")
        .style("Heading1")
        .alignment(Alignment::Center);

    doc.add_paragraph("Project Proposal")
        .alignment(Alignment::Center);

    doc.add_paragraph("");

    // ── Summary section ──
    doc.add_paragraph("Executive Summary")
        .style("Heading2");

    doc.add_paragraph(
        "This proposal outlines the {{project_name}} project for {{company_name}}. \
         The primary contact is {{contact_name}} ({{contact_email}}). \
         The proposed start date is {{start_date}} with an estimated duration of {{duration}}."
    );

    doc.add_paragraph("");

    // ── Cross-run placeholder (bold label + normal value) ──
    doc.add_paragraph("Key Details")
        .style("Heading2");

    {
        let mut p = doc.add_paragraph("");
        p.add_run("Project: ").bold(true);
        p.add_run("{{project_name}}");
    }
    {
        let mut p = doc.add_paragraph("");
        p.add_run("Budget: ").bold(true);
        p.add_run("{{budget}}");
    }
    {
        let mut p = doc.add_paragraph("");
        p.add_run("Status: ").bold(true);
        p.add_run("{{status}}");
    }

    doc.add_paragraph("");

    // ── Table with placeholders ──
    doc.add_paragraph("Team Members")
        .style("Heading2");

    {
        let mut tbl = doc.add_table(4, 3);
        tbl = tbl.borders(BorderStyle::Single, 4, "000000");

        // Header row
        for col in 0..3 {
            tbl.cell(0, col).unwrap().shading("2E75B6");
        }
        tbl.cell(0, 0).unwrap().set_text("Name");
        tbl.cell(0, 1).unwrap().set_text("Role");
        tbl.cell(0, 2).unwrap().set_text("Email");

        tbl.cell(1, 0).unwrap().set_text("{{member1_name}}");
        tbl.cell(1, 1).unwrap().set_text("{{member1_role}}");
        tbl.cell(1, 2).unwrap().set_text("{{member1_email}}");

        tbl.cell(2, 0).unwrap().set_text("{{member2_name}}");
        tbl.cell(2, 1).unwrap().set_text("{{member2_role}}");
        tbl.cell(2, 2).unwrap().set_text("{{member2_email}}");

        tbl.cell(3, 0).unwrap().set_text("{{member3_name}}");
        tbl.cell(3, 1).unwrap().set_text("{{member3_role}}");
        tbl.cell(3, 2).unwrap().set_text("{{member3_email}}");
    }

    doc.add_paragraph("");

    // ── Deliverables section ──
    doc.add_paragraph("Deliverables")
        .style("Heading2");

    doc.add_paragraph("INSERTION_POINT");

    doc.add_paragraph("");

    // ── Signature block ──
    doc.add_paragraph("Acceptance")
        .style("Heading2");

    doc.add_paragraph(
        "By signing below, {{company_name}} agrees to the terms outlined in this proposal."
    );

    {
        let mut tbl = doc.add_table(2, 2);
        tbl = tbl.borders(BorderStyle::Single, 4, "000000");
        tbl.cell(0, 0).unwrap().set_text("Customer: ___________________");
        tbl.cell(0, 1).unwrap().set_text("Provider: ___________________");
        tbl.cell(1, 0).unwrap().set_text("Date: {{date}}");
        tbl.cell(1, 1).unwrap().set_text("Date: {{date}}");
    }

    doc.set_title("{{company_name}} — Project Proposal Template");
    doc.set_author("Template Generator");

    doc.save(path).unwrap();
}

/// Open the template, replace all placeholders, insert content, and save.
fn fill_template(template_path: &Path, output_path: &Path) {
    let mut doc = Document::open(template_path).unwrap();

    // ── Batch replacement ──
    let mut replacements = HashMap::new();
    replacements.insert("{{company_name}}", "Riverside Medical Center");
    replacements.insert("{{project_name}}", "Network Security Upgrade");
    replacements.insert("{{contact_name}}", "Dr. Sarah Chen");
    replacements.insert("{{contact_email}}", "s.chen@riverside.org");
    replacements.insert("{{start_date}}", "March 1, 2026");
    replacements.insert("{{duration}}", "12 weeks");
    replacements.insert("{{budget}}", "$185,000");
    replacements.insert("{{status}}", "Pending Approval");
    replacements.insert("{{author_name}}", "James Wilson");
    replacements.insert("{{date}}", "February 22, 2026");

    // Team members
    replacements.insert("{{member1_name}}", "James Wilson");
    replacements.insert("{{member1_role}}", "Project Lead");
    replacements.insert("{{member1_email}}", "j.wilson@provider.com");
    replacements.insert("{{member2_name}}", "Maria Garcia");
    replacements.insert("{{member2_role}}", "Security Architect");
    replacements.insert("{{member2_email}}", "m.garcia@provider.com");
    replacements.insert("{{member3_name}}", "David Park");
    replacements.insert("{{member3_role}}", "Network Engineer");
    replacements.insert("{{member3_email}}", "d.park@provider.com");

    let count = doc.replace_all(&replacements);
    println!("  Replaced {} placeholders", count);

    // ── Insert deliverables at the insertion point ──
    if let Some(idx) = doc.find_content_index("INSERTION_POINT") {
        // Remove the placeholder paragraph
        doc.remove_content(idx);

        // Insert deliverables list
        doc.insert_paragraph(idx, "The following deliverables are included:");

        // Insert a deliverables table
        let mut tbl = doc.insert_table(idx + 1, 5, 3);
        tbl = tbl.borders(BorderStyle::Single, 4, "000000");

        for col in 0..3 {
            tbl.cell(0, col).unwrap().shading("E2EFDA");
        }
        tbl.cell(0, 0).unwrap().set_text("Phase");
        tbl.cell(0, 1).unwrap().set_text("Description");
        tbl.cell(0, 2).unwrap().set_text("Timeline");

        tbl.cell(1, 0).unwrap().set_text("1. Discovery");
        tbl.cell(1, 1).unwrap().set_text("Network assessment and asset inventory");
        tbl.cell(1, 2).unwrap().set_text("Weeks 1-3");

        tbl.cell(2, 0).unwrap().set_text("2. Design");
        tbl.cell(2, 1).unwrap().set_text("Security architecture and policy design");
        tbl.cell(2, 2).unwrap().set_text("Weeks 4-6");

        tbl.cell(3, 0).unwrap().set_text("3. Implementation");
        tbl.cell(3, 1).unwrap().set_text("Deploy monitoring and access controls");
        tbl.cell(3, 2).unwrap().set_text("Weeks 7-10");

        tbl.cell(4, 0).unwrap().set_text("4. Validation");
        tbl.cell(4, 1).unwrap().set_text("Testing, training, and handover");
        tbl.cell(4, 2).unwrap().set_text("Weeks 11-12");

        println!("  Inserted deliverables table at position {}", idx);
    }

    // ── Update metadata ──
    doc.set_title("Riverside Medical Center — Network Security Upgrade Proposal");
    doc.set_author("James Wilson");
    doc.set_subject("Project Proposal");
    doc.set_keywords("security, network, medical, proposal");

    doc.save(output_path).unwrap();
}
