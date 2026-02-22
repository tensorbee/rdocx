//! CLI command implementations.

use std::path::Path;

use rdocx::Document;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

/// Inspect a DOCX file and print structure information.
pub fn inspect(file: &Path, json: bool) -> Result<()> {
    let doc = Document::open(file)?;

    let paragraph_count = doc.paragraph_count();
    let table_count = doc.table_count();
    let content_count = doc.content_count();

    let title = doc.title().map(|s| s.to_string());
    let author = doc.author().map(|s| s.to_string());
    let subject = doc.subject().map(|s| s.to_string());
    let keywords = doc.keywords().map(|s| s.to_string());

    // Collect style IDs used in paragraphs
    let mut style_ids: Vec<String> = Vec::new();
    for para in doc.paragraphs() {
        if let Some(style) = para.style_id() {
            let s = style.to_string();
            if !style_ids.contains(&s) {
                style_ids.push(s);
            }
        }
    }

    if json {
        let obj = serde_json::json!({
            "file": file.display().to_string(),
            "paragraphs": paragraph_count,
            "tables": table_count,
            "content_elements": content_count,
            "metadata": {
                "title": title,
                "author": author,
                "subject": subject,
                "keywords": keywords,
            },
            "styles_used": style_ids,
        });
        println!("{}", serde_json::to_string_pretty(&obj)?);
    } else {
        println!("File: {}", file.display());
        println!("Paragraphs: {paragraph_count}");
        println!("Tables: {table_count}");
        println!("Content elements: {content_count}");
        println!();
        println!("Metadata:");
        if let Some(t) = &title {
            println!("  Title: {t}");
        }
        if let Some(a) = &author {
            println!("  Author: {a}");
        }
        if let Some(s) = &subject {
            println!("  Subject: {s}");
        }
        if let Some(k) = &keywords {
            println!("  Keywords: {k}");
        }
        if title.is_none() && author.is_none() && subject.is_none() && keywords.is_none() {
            println!("  (none)");
        }
        println!();
        println!("Styles used:");
        if style_ids.is_empty() {
            println!("  (none)");
        } else {
            for sid in &style_ids {
                println!("  - {sid}");
            }
        }
    }

    Ok(())
}

/// Extract plain text from a DOCX file.
pub fn text(file: &Path) -> Result<()> {
    let doc = Document::open(file)?;

    for para in doc.paragraphs() {
        println!("{}", para.text());
    }

    Ok(())
}

/// Convert a DOCX file to another format.
pub fn convert(file: &Path, to: &str, output: Option<&Path>, dpi: u32, font_dir: Option<&Path>) -> Result<()> {
    let doc = Document::open(file)?;

    let default_ext = match to {
        "pdf" => "pdf",
        "html" => "html",
        "md" | "markdown" => "md",
        "png" => "png",
        other => {
            return Err(format!("Unknown format: {other}. Supported: pdf, html, md, png").into());
        }
    };

    let output_path = match output {
        Some(p) => p.to_path_buf(),
        None => file.with_extension(default_ext),
    };

    match to {
        "pdf" => {
            let bytes = if let Some(dir) = font_dir {
                let font_files = Document::load_fonts_from_dir(dir);
                let font_refs: Vec<(&str, &[u8])> = font_files
                    .iter()
                    .map(|f| (f.family.as_str(), f.data.as_slice()))
                    .collect();
                doc.to_pdf_with_fonts(&font_refs)?
            } else {
                doc.to_pdf()?
            };
            std::fs::write(&output_path, bytes)?;
        }
        "html" => {
            let html = doc.to_html();
            std::fs::write(&output_path, html)?;
        }
        "md" | "markdown" => {
            let md = doc.to_markdown();
            std::fs::write(&output_path, md)?;
        }
        "png" => {
            let pages = doc.render_all_pages(dpi as f64)?;
            if pages.len() == 1 {
                std::fs::write(&output_path, &pages[0])?;
            } else {
                let stem = output_path.file_stem().unwrap_or_default().to_string_lossy();
                let parent = output_path.parent().unwrap_or(Path::new("."));
                for (i, page) in pages.iter().enumerate() {
                    let page_path = parent.join(format!("{stem}_{:03}.png", i + 1));
                    std::fs::write(&page_path, page)?;
                }
                println!(
                    "Written {} pages to {}/{stem}_NNN.png",
                    pages.len(),
                    parent.display()
                );
                return Ok(());
            }
        }
        _ => unreachable!(),
    }

    println!("Written to {}", output_path.display());
    Ok(())
}

/// Structural diff between two DOCX files.
pub fn diff(file_a: &Path, file_b: &Path) -> Result<()> {
    let doc_a = Document::open(file_a)?;
    let doc_b = Document::open(file_b)?;

    let paras_a: Vec<String> = doc_a.paragraphs().iter().map(|p| p.text()).collect();
    let paras_b: Vec<String> = doc_b.paragraphs().iter().map(|p| p.text()).collect();

    println!(
        "--- {} ({} paragraphs, {} tables)",
        file_a.display(),
        doc_a.paragraph_count(),
        doc_a.table_count()
    );
    println!(
        "+++ {} ({} paragraphs, {} tables)",
        file_b.display(),
        doc_b.paragraph_count(),
        doc_b.table_count()
    );
    println!();

    // Simple LCS-based diff on paragraph texts
    let lcs = compute_lcs(&paras_a, &paras_b);
    let mut i = 0;
    let mut j = 0;
    let mut k = 0;

    while k < lcs.len() {
        // Output removed lines before the match
        while i < paras_a.len() && paras_a[i] != lcs[k] {
            println!("- [{}] {}", i + 1, paras_a[i]);
            i += 1;
        }
        // Output added lines before the match
        while j < paras_b.len() && paras_b[j] != lcs[k] {
            println!("+ [{}] {}", j + 1, paras_b[j]);
            j += 1;
        }
        // Skip the common line
        i += 1;
        j += 1;
        k += 1;
    }

    // Remaining lines
    while i < paras_a.len() {
        println!("- [{}] {}", i + 1, paras_a[i]);
        i += 1;
    }
    while j < paras_b.len() {
        println!("+ [{}] {}", j + 1, paras_b[j]);
        j += 1;
    }

    let changes = paras_a.len() + paras_b.len() - 2 * lcs.len();
    if changes == 0 {
        println!("(no differences in paragraph text)");
    } else {
        println!("\n{changes} paragraph(s) differ.");
    }

    Ok(())
}

/// Compute the longest common subsequence of two string slices.
fn compute_lcs(a: &[String], b: &[String]) -> Vec<String> {
    let m = a.len();
    let n = b.len();
    let mut dp = vec![vec![0u32; n + 1]; m + 1];

    for i in 1..=m {
        for j in 1..=n {
            dp[i][j] = if a[i - 1] == b[j - 1] {
                dp[i - 1][j - 1] + 1
            } else {
                dp[i - 1][j].max(dp[i][j - 1])
            };
        }
    }

    // Backtrack
    let mut result = Vec::new();
    let mut i = m;
    let mut j = n;
    while i > 0 && j > 0 {
        if a[i - 1] == b[j - 1] {
            result.push(a[i - 1].clone());
            i -= 1;
            j -= 1;
        } else if dp[i - 1][j] >= dp[i][j - 1] {
            i -= 1;
        } else {
            j -= 1;
        }
    }

    result.reverse();
    result
}

/// Replace a placeholder in a DOCX file and save to output.
pub fn replace(
    file: &Path,
    placeholder: &str,
    value: &str,
    output: &Path,
) -> Result<()> {
    let mut doc = Document::open(file)?;
    let count = doc.replace_text(placeholder, value);
    doc.save(output)?;
    println!("Replaced {count} occurrence(s) of \"{placeholder}\" -> \"{value}\"");
    println!("Written to {}", output.display());
    Ok(())
}

/// Render document pages to PNG images.
pub fn render(file: &Path, output_dir: Option<&Path>, dpi: f64, page: Option<usize>) -> Result<()> {
    let doc = Document::open(file)?;
    let out_dir = output_dir.unwrap_or_else(|| Path::new("."));

    let stem = file.file_stem().unwrap_or_default().to_string_lossy();

    if let Some(page_idx) = page {
        let png = doc
            .render_page_to_png(page_idx, dpi)?
            .ok_or_else(|| format!("Page {page_idx} not found"))?;
        let out_path = out_dir.join(format!("{stem}_page{}.png", page_idx + 1));
        std::fs::write(&out_path, &png)?;
        println!(
            "Page {} -> {} ({} bytes)",
            page_idx + 1,
            out_path.display(),
            png.len()
        );
    } else {
        let pages = doc.render_all_pages(dpi)?;
        for (i, png) in pages.iter().enumerate() {
            let out_path = out_dir.join(format!("{stem}_page{}.png", i + 1));
            std::fs::write(&out_path, png)?;
            println!(
                "Page {} -> {} ({} bytes)",
                i + 1,
                out_path.display(),
                png.len()
            );
        }
        println!("Rendered {} page(s) at {dpi} DPI", pages.len());
    }

    Ok(())
}

/// Validate OOXML conformance of a DOCX file.
pub fn validate(file: &Path) -> Result<()> {
    let doc = Document::open(file)?;

    let mut issues: Vec<String> = Vec::new();

    // Check: Empty document
    if doc.content_count() == 0 {
        issues.push("Document has no content (no paragraphs or tables)".to_string());
    }

    // Check: Empty paragraphs (warning)
    let mut empty_count = 0;
    for para in doc.paragraphs() {
        if para.text().trim().is_empty() {
            empty_count += 1;
        }
    }
    if empty_count > 0 {
        issues.push(format!("{empty_count} empty paragraph(s) found"));
    }

    // Check: Heading level gaps
    let mut prev_level: Option<u32> = None;
    for para in doc.paragraphs() {
        if let Some(style_id) = para.style_id() {
            if let Some(level_str) = style_id.strip_prefix("Heading") {
                if let Ok(level) = level_str.parse::<u32>() {
                    if let Some(prev) = prev_level {
                        if level > prev + 1 {
                            issues.push(format!(
                                "Heading level gap: Heading{prev} -> Heading{level} (skipped level(s))"
                            ));
                        }
                    }
                    prev_level = Some(level);
                }
            }
        }
    }

    // Check: Missing metadata
    if doc.title().is_none() {
        issues.push("Missing document title".to_string());
    }
    if doc.author().is_none() {
        issues.push("Missing document author".to_string());
    }

    if issues.is_empty() {
        println!("OK — no issues found in {}", file.display());
    } else {
        println!("Found {} issue(s) in {}:", issues.len(), file.display());
        for (i, issue) in issues.iter().enumerate() {
            println!("  {}. {issue}", i + 1);
        }
    }

    Ok(())
}
