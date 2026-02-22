use rdocx::Document;

fn main() {
    let doc = Document::open("samples/feature_showcase.docx").expect("Failed to open document");

    let html = doc.to_html();
    std::fs::write("/tmp/feature_showcase.html", &html).expect("Failed to write HTML");
    println!("HTML: {} bytes -> /tmp/feature_showcase.html", html.len());

    let md = doc.to_markdown();
    std::fs::write("/tmp/feature_showcase.md", &md).expect("Failed to write Markdown");
    println!(
        "Markdown: {} bytes -> /tmp/feature_showcase.md",
        md.len()
    );

    // Simple document
    let mut simple = Document::new();
    simple.add_paragraph("Hello, World!");
    let html = simple.to_html();
    println!("\n--- Simple HTML ---\n{html}");

    let md = simple.to_markdown();
    println!("--- Simple Markdown ---\n{md}");
}
