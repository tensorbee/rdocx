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
    ///
    /// Give one pair with -p and -v, or many with --map. Nothing is written
    /// unless every pair replaced its expected count, or at least one
    /// occurrence when it gives no count.
    Replace {
        file: PathBuf,
        #[arg(
            long,
            short = 'p',
            required_unless_present = "map",
            requires = "value",
            conflicts_with = "map"
        )]
        placeholder: Option<String>,
        #[arg(long, short = 'v', requires = "placeholder")]
        value: Option<String>,
        /// JSON file holding an array of pairs, applied in order, such as
        /// [{"placeholder": "{{name}}", "value": "Ada", "expect": 2}]
        #[arg(long, value_name = "PAIRS_JSON")]
        map: Option<PathBuf>,
        /// Require exactly this many slide and speaker-note replacements
        #[arg(long, conflicts_with = "map")]
        expect: Option<usize>,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the count of each pair as JSON
        #[arg(long)]
        json: bool,
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
    /// Inspect and mutate modern PowerPoint comment threads
    Comment {
        #[command(subcommand)]
        command: CommentCommand,
    },
    /// Add, duplicate, remove, move, hide and show slides
    Slide {
        #[command(subcommand)]
        command: SlideCommand,
    },
    /// Write speaker notes
    Notes {
        #[command(subcommand)]
        command: NotesCommand,
    },
    /// Report every text frame whose text overflows it, exit 1 when one does
    /// and 2 on an error
    ///
    /// Lays out the text of the slide shapes with rpptx's renderer. Each
    /// overflowing frame is listed with the largest font scale, in steps of
    /// 2.5% down to 25%, at which its text fits, as PowerPoint's "shrink text
    /// on overflow" computes it, or none when even 25% overflows. Tables and
    /// SmartArt are not checked.
    Fit {
        file: PathBuf,
        /// Output the report as JSON
        #[arg(long)]
        json: bool,
    },
    /// Read and write the core document properties
    Meta {
        #[command(subcommand)]
        command: MetaCommand,
    },
}

#[derive(Subcommand)]
enum SlideCommand {
    /// Add a slide from a layout, at the end or at a one-based position
    Add {
        file: PathBuf,
        /// Layout name, or its one-based number in master order
        #[arg(long)]
        layout: String,
        /// One-based position of the new slide, the end when absent
        #[arg(long)]
        at: Option<usize>,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Duplicate one slide, inserting the copy right after it
    Duplicate {
        file: PathBuf,
        /// One-based slide number
        slide: usize,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Remove one slide with its notes and comments
    Remove {
        file: PathBuf,
        /// One-based slide number
        slide: usize,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Move one slide to another one-based position
    Move {
        file: PathBuf,
        /// One-based slide number
        slide: usize,
        /// One-based final position
        #[arg(long)]
        to: usize,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Hide one slide from the slide show
    Hide {
        file: PathBuf,
        /// One-based slide number
        slide: usize,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Show one hidden slide in the slide show again
    Show {
        file: PathBuf,
        /// One-based slide number
        slide: usize,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
}

#[derive(Subcommand)]
enum NotesCommand {
    /// Replace the speaker notes of one slide, creating them when absent
    Set {
        file: PathBuf,
        /// One-based slide number
        slide: usize,
        /// Notes text, one paragraph per line
        #[arg(
            long,
            required_unless_present = "from_file",
            conflicts_with = "from_file"
        )]
        text: Option<String>,
        /// Read the notes text from a UTF-8 file
        #[arg(long, value_name = "PATH")]
        from_file: Option<PathBuf>,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
}

#[derive(Subcommand)]
enum MetaCommand {
    /// Print the core properties
    Get {
        file: PathBuf,
        /// Output as JSON
        #[arg(long)]
        json: bool,
    },
    /// Set core properties
    Set {
        file: PathBuf,
        #[arg(long)]
        title: Option<String>,
        /// Author (dc:creator)
        #[arg(long)]
        author: Option<String>,
        #[arg(long)]
        subject: Option<String>,
        #[arg(long)]
        keywords: Option<String>,
        /// Description (comments)
        #[arg(long)]
        description: Option<String>,
        #[arg(long)]
        category: Option<String>,
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the properties written as JSON
        #[arg(long)]
        json: bool,
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
    // `validate` and `fit` carry a verdict in their exit status. `fit` exits
    // 2 on an error, so that 1 always means an overflow.
    let error_status = if matches!(cli.command, Command::Fit { .. }) {
        2
    } else {
        1
    };
    let verdict = match &cli.command {
        Command::Validate { file } => Some(commands::validate(file)),
        Command::Fit { file, json } => Some(commands::fit(file, *json)),
        _ => None,
    };
    if let Some(verdict) = verdict {
        match verdict {
            Ok(true) => return,
            Ok(false) => process::exit(1),
            Err(error) => {
                eprintln!("Error: {error}");
                process::exit(error_status);
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
            map,
            expect,
            output,
            json,
        } => commands::replace(
            &file,
            placeholder.zip(value),
            map.as_deref(),
            expect,
            &output,
            json,
        ),
        Command::Validate { .. } | Command::Fit { .. } => {
            unreachable!("verdict commands are dispatched above")
        }
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
        Command::Slide { command } => match command {
            SlideCommand::Add {
                file,
                layout,
                at,
                output,
                json,
            } => commands::slide_add(&file, &layout, at, &output, json),
            SlideCommand::Duplicate {
                file,
                slide,
                output,
                json,
            } => commands::slide_edit(&file, commands::SlideEdit::Duplicate(slide), &output, json),
            SlideCommand::Remove {
                file,
                slide,
                output,
                json,
            } => commands::slide_edit(&file, commands::SlideEdit::Remove(slide), &output, json),
            SlideCommand::Move {
                file,
                slide,
                to,
                output,
                json,
            } => commands::slide_edit(&file, commands::SlideEdit::Move(slide, to), &output, json),
            SlideCommand::Hide {
                file,
                slide,
                output,
                json,
            } => commands::slide_edit(
                &file,
                commands::SlideEdit::Hidden(slide, true),
                &output,
                json,
            ),
            SlideCommand::Show {
                file,
                slide,
                output,
                json,
            } => commands::slide_edit(
                &file,
                commands::SlideEdit::Hidden(slide, false),
                &output,
                json,
            ),
        },
        Command::Notes { command } => match command {
            NotesCommand::Set {
                file,
                slide,
                text,
                from_file,
                output,
                json,
            } => commands::notes_set(
                &file,
                slide,
                text.as_deref(),
                from_file.as_deref(),
                &output,
                json,
            ),
        },
        Command::Meta { command } => match command {
            MetaCommand::Get { file, json } => commands::meta_get(&file, json),
            MetaCommand::Set {
                file,
                title,
                author,
                subject,
                keywords,
                description,
                category,
                output,
                json,
            } => commands::meta_set(
                &file,
                [title, author, subject, keywords, description, category],
                &output,
                json,
            ),
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
