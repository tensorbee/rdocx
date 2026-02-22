//! rdocx CLI — "jq for DOCX"
//!
//! Inspect, convert, diff, and manipulate DOCX files from the command line.

use std::path::PathBuf;
use std::process;

use clap::{Parser, Subcommand};

mod commands;

#[derive(Parser)]
#[command(name = "rdocx", version, about = "CLI tool for DOCX files")]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Print document structure: paragraph/table count, styles, images, metadata
    Inspect {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output as JSON
        #[arg(long)]
        json: bool,
    },
    /// Extract plain text from a DOCX file
    Text {
        /// Path to the DOCX file
        file: PathBuf,
    },
    /// Convert DOCX to another format (pdf, html, md, png)
    Convert {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output format: pdf, html, md, png
        #[arg(long, short = 't')]
        to: String,
        /// Output file path (defaults to input with new extension)
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,
        /// DPI for image rendering (default: 150)
        #[arg(long, default_value = "150")]
        dpi: u32,
        /// Directory containing font files (.ttf/.otf) to use for PDF rendering
        #[arg(long)]
        font_dir: Option<PathBuf>,
    },
    /// Structural diff between two DOCX files
    Diff {
        /// First DOCX file
        file_a: PathBuf,
        /// Second DOCX file
        file_b: PathBuf,
    },
    /// Replace placeholders in a DOCX file
    Replace {
        /// Path to the DOCX file
        file: PathBuf,
        /// Placeholder string
        #[arg(long, short = 'p')]
        placeholder: String,
        /// Replacement value
        #[arg(long, short = 'v')]
        value: String,
        /// Output file path
        #[arg(long, short = 'o')]
        output: PathBuf,
    },
    /// Validate OOXML conformance
    Validate {
        /// Path to the DOCX file
        file: PathBuf,
    },
    /// Render pages to PNG images
    Render {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output directory (defaults to current directory)
        #[arg(long, short = 'o')]
        output_dir: Option<PathBuf>,
        /// DPI resolution (default: 150)
        #[arg(long, default_value = "150")]
        dpi: f64,
        /// Render only a specific page (0-based index)
        #[arg(long)]
        page: Option<usize>,
    },
}

fn main() {
    let cli = Cli::parse();

    let result = match cli.command {
        Command::Inspect { file, json } => commands::inspect(&file, json),
        Command::Text { file } => commands::text(&file),
        Command::Convert {
            file,
            to,
            output,
            dpi,
            font_dir,
        } => commands::convert(&file, &to, output.as_deref(), dpi, font_dir.as_deref()),
        Command::Diff { file_a, file_b } => commands::diff(&file_a, &file_b),
        Command::Replace {
            file,
            placeholder,
            value,
            output,
        } => commands::replace(&file, &placeholder, &value, &output),
        Command::Validate { file } => commands::validate(&file),
        Command::Render {
            file,
            output_dir,
            dpi,
            page,
        } => commands::render(&file, output_dir.as_deref(), dpi, page),
    };

    if let Err(e) = result {
        eprintln!("Error: {e}");
        process::exit(1);
    }
}
