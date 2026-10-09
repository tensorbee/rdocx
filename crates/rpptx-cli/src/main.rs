//! Command-line access to the public `rpptx` facade.

use std::io::{self, Write};
use std::path::PathBuf;
use std::process;

use clap::{Parser, Subcommand};

mod commands;

#[derive(Parser)]
#[command(name = "rpptx", version, about = "CLI tool for PPTX files")]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Print presentation structure and metadata
    Inspect {
        file: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// Extract slide text in presentation order
    Text {
        file: PathBuf,
        /// Output paragraphs, runs, and speaker notes as schema-1 JSON
        #[arg(long)]
        json: bool,
        /// Print speaker notes after each slide's text (JSON always includes them)
        #[arg(long)]
        notes: bool,
    },
    /// Convert a presentation to deterministic PDF or image output
    Convert {
        file: PathBuf,
        #[arg(long, short = 't')]
        to: String,
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,
        /// Replace existing output files, but never the input file
        #[arg(long)]
        force: bool,
        #[arg(long, default_value = "150")]
        dpi: f64,
        /// One-based slide range for image output, such as 1,3-5
        #[arg(long)]
        slides: Option<String>,
        /// JPEG quality from 1 through 100
        #[arg(long, default_value = "90")]
        quality: u8,
        /// Preserve unpainted PNG pixels as transparent
        #[arg(long)]
        transparent: bool,
    },
    /// Compare slide text using a longest-common-subsequence diff
    Diff { file_a: PathBuf, file_b: PathBuf },
    /// Replace literal presentation text while retaining run formatting
    Replace {
        file: PathBuf,
        #[arg(long, short = 'p')]
        placeholder: String,
        #[arg(long, short = 'v')]
        value: String,
        /// Require exactly this many slide and speaker-note replacements
        #[arg(long)]
        expect: Option<usize>,
        #[arg(long, short = 'o')]
        output: PathBuf,
    },
    /// Validate package and PresentationML invariants
    Validate { file: PathBuf },
    /// Render selected slides to deterministic image files
    Render {
        file: PathBuf,
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,
        /// Replace existing slide images, but never the input file
        #[arg(long)]
        force: bool,
        #[arg(long, default_value = "150")]
        dpi: f64,
        #[arg(long)]
        slide: Option<String>,
        /// Output format: png, jpeg, tiff
        #[arg(long, default_value = "png")]
        format: String,
        /// JPEG quality from 1 through 100
        #[arg(long, default_value = "90")]
        quality: u8,
        /// Preserve unpainted PNG pixels as transparent
        #[arg(long)]
        transparent: bool,
    },
    /// Render slide one as a proportional 320-pixel-wide PNG
    Thumbnail {
        file: PathBuf,
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,
        /// Replace an existing output file, but never the input file
        #[arg(long)]
        force: bool,
    },
    /// Print each slide title and recursive paragraph outline
    Outline {
        file: PathBuf,
        /// Output titles, outline items, and speaker notes as schema-1 JSON
        #[arg(long)]
        json: bool,
        /// Print speaker notes after each slide's outline (JSON always includes them)
        #[arg(long)]
        notes: bool,
    },
    /// Set the slide number, footer and date of every slide, as PowerPoint's
    /// Header and Footer dialog with Apply to All. Every flag is opt-in, while
    /// Python's set_header_footer turns the slide number and the title-slide
    /// skip on by default
    Footer {
        file: PathBuf,
        /// Show the slide number
        #[arg(long)]
        slide_number: bool,
        /// Footer text, no footer when omitted
        #[arg(long)]
        footer: Option<String>,
        /// off, auto for a date PowerPoint refreshes, cached with today's local
        /// date, or fixed text
        #[arg(long, default_value = "off")]
        date: String,
        /// Date field format for --date auto, datetime1 to datetime7
        #[arg(long, default_value = "datetime1")]
        date_format: String,
        /// Show none of them on slides with a title slide layout
        #[arg(long)]
        skip_title: bool,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Inspect and mutate modern PowerPoint comment threads
    Comment {
        #[command(subcommand)]
        command: CommentCommand,
    },
}

#[derive(Subcommand)]
enum CommentCommand {
    /// List modern comments and their replies in slide order
    List {
        file: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// Add a comment to one slide
    Add {
        file: PathBuf,
        /// One-based slide number
        #[arg(long)]
        slide: usize,
        /// Comment author, reused by name or added to the author list
        #[arg(long)]
        author: String,
        /// Initials recorded when the author is added
        #[arg(long)]
        initials: Option<String>,
        #[arg(long)]
        text: String,
        /// RFC 3339 creation timestamp
        #[arg(long)]
        date: String,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Reply to an existing comment thread
    Reply {
        file: PathBuf,
        /// Comment id of the thread
        #[arg(long)]
        id: String,
        /// Reply author, reused by name or added to the author list
        #[arg(long)]
        author: String,
        #[arg(long)]
        text: String,
        /// RFC 3339 creation timestamp
        #[arg(long)]
        date: String,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Mark one comment thread resolved
    Resolve {
        file: PathBuf,
        /// Comment id of the thread
        #[arg(long)]
        id: String,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Remove one comment with its replies, or one reply
    Remove {
        file: PathBuf,
        /// Comment or reply id
        #[arg(long)]
        id: String,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
}

fn main() {
    #[cfg(windows)]
    {
        // A full validation can exceed the one MiB Windows main-thread stack.
        std::thread::Builder::new()
            .stack_size(8 * 1024 * 1024)
            .spawn(run_cli)
            .expect("start rpptx CLI thread")
            .join()
            .expect("rpptx CLI thread panicked");
    }
    #[cfg(not(windows))]
    run_cli();
}

fn run_cli() {
    let cli = Cli::parse();
    if let Command::Validate { file } = &cli.command {
        match commands::validate(file) {
            Ok(true) => return,
            Ok(false) => process::exit(1),
            Err(error) => {
                eprintln!("Error: {error}");
                process::exit(1);
            }
        }
    }

    let result = match cli.command {
        Command::Inspect { file, json } => commands::inspect(&file, json),
        Command::Text { file, json, notes } => commands::text(&file, json, notes),
        Command::Convert {
            file,
            to,
            output,
            force,
            dpi,
            slides,
            quality,
            transparent,
        } => commands::convert(
            &file,
            &to,
            output.as_deref(),
            force,
            dpi,
            commands::ImageOptions {
                slides: slides.as_deref(),
                quality,
                transparent,
            },
        ),
        Command::Diff { file_a, file_b } => commands::diff(&file_a, &file_b),
        Command::Replace {
            file,
            placeholder,
            value,
            expect,
            output,
        } => commands::replace(&file, &placeholder, &value, expect, &output),
        Command::Validate { .. } => unreachable!("validate is dispatched above"),
        Command::Render {
            file,
            output,
            force,
            dpi,
            slide,
            format,
            quality,
            transparent,
        } => commands::render(
            &file,
            output.as_deref(),
            force,
            dpi,
            &format,
            commands::ImageOptions {
                slides: slide.as_deref(),
                quality,
                transparent,
            },
        ),
        Command::Thumbnail {
            file,
            output,
            force,
        } => commands::thumbnail(&file, output.as_deref(), force),
        Command::Outline { file, json, notes } => commands::outline(&file, json, notes),
        Command::Footer {
            file,
            slide_number,
            footer,
            date,
            date_format,
            skip_title,
            output,
            json,
        } => commands::footer(
            &file,
            commands::FooterInput {
                slide_number,
                footer: footer.as_deref(),
                date: &date,
                date_format: &date_format,
                skip_title,
            },
            &output,
            json,
        ),
        Command::Comment { command } => match command {
            CommentCommand::List { file, json } => commands::comment_list(&file, json),
            CommentCommand::Add {
                file,
                slide,
                author,
                initials,
                text,
                date,
                output,
                json,
            } => commands::comment_add(
                &file,
                slide,
                commands::CommentInput {
                    author: &author,
                    initials: initials.as_deref(),
                    text: &text,
                    date: &date,
                },
                &output,
                json,
            ),
            CommentCommand::Reply {
                file,
                id,
                author,
                text,
                date,
                output,
                json,
            } => commands::comment_reply(
                &file,
                &id,
                commands::CommentInput {
                    author: &author,
                    initials: None,
                    text: &text,
                    date: &date,
                },
                &output,
                json,
            ),
            CommentCommand::Resolve {
                file,
                id,
                output,
                json,
            } => commands::comment_resolve(&file, &id, &output, json),
            CommentCommand::Remove {
                file,
                id,
                output,
                json,
            } => commands::comment_remove(&file, &id, &output, json),
        },
    };
    // Standard output is line buffered, so a last line without a newline is
    // only written, and can only fail, when it is flushed.
    let result = result.and_then(|()| io::stdout().flush().map_err(Into::into));
    if let Err(error) = result {
        // A reader that closes standard output early, as `| head` does, ends
        // the output. That is not a failure of the command.
        let closed_stdout = error
            .downcast_ref::<io::Error>()
            .is_some_and(|error| error.kind() == io::ErrorKind::BrokenPipe);
        if !closed_stdout {
            eprintln!("Error: {error}");
            process::exit(1);
        }
    }
}
