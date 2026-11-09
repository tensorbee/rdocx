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
enum FontCommand {
    /// List native font records and faces
    List {
        file: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// Embed one caller-owned font face with explicit license authorization
    Embed {
        file: PathBuf,
        /// Exact font record/typeface name
        #[arg(long)]
        name: String,
        /// Local TrueType or OpenType font file
        #[arg(long)]
        font: PathBuf,
        #[arg(long, default_value = "regular", value_parser = ["regular", "bold", "italic", "bold-italic"])]
        kind: String,
        /// Explicitly authorize embedding under the identified license
        #[arg(long, required = true)]
        license_authorized: bool,
        #[arg(long)]
        license_identity: String,
        #[arg(long, short = 'o')]
        output: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// Remove a named font record, or only the selected embedded face
    Remove {
        file: PathBuf,
        #[arg(long)]
        name: String,
        #[arg(long, value_parser = ["regular", "bold", "italic", "bold-italic"])]
        kind: Option<String>,
        #[arg(long, short = 'o')]
        output: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// Read flags, or set explicitly named tri-state flags into a new output
    Settings {
        file: PathBuf,
        #[arg(long, value_parser = ["true", "false", "none"])]
        embed_true_type_fonts: Option<String>,
        #[arg(long, value_parser = ["true", "false", "none"])]
        save_subset_fonts: Option<String>,
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,
        #[arg(long)]
        json: bool,
    },
}

#[derive(Subcommand)]
enum Command {
    /// Sign current package content with a local PKCS#8 key and X.509 certificate
    Sign {
        file: PathBuf,
        #[arg(long)]
        key: PathBuf,
        #[arg(long)]
        cert: PathBuf,
        #[arg(short = 'o', long)]
        output: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// Verify signatures and complete coverage, without asserting chain trust
    VerifySignatures {
        file: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// List, embed, remove and configure explicitly licensed fonts
    Font {
        #[command(subcommand)]
        command: FontCommand,
    },
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
    /// Convert a presentation to PDF, images, semantic HTML, or Markdown
    Convert {
        file: PathBuf,
        /// Directory containing caller font files (.ttf/.otf/.ttc)
        #[arg(long)]
        font_dir: Option<PathBuf>,
        #[arg(long, short = 't')]
        to: String,
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,
        /// Replace existing output files, but never the input file
        #[arg(long)]
        force: bool,
        /// Raster resolution, default 150
        #[arg(long)]
        dpi: Option<f64>,
        /// One-based slide range for image output, such as 1,3-5
        #[arg(long)]
        slides: Option<String>,
        /// JPEG quality from 1 through 100
        #[arg(long)]
        quality: Option<u8>,
        /// Preserve unpainted PNG pixels as transparent
        #[arg(long)]
        transparent: bool,
        /// Include speaker notes in semantic HTML or Markdown
        #[arg(long)]
        notes: bool,
        /// Omit hidden slides from semantic HTML or Markdown
        #[arg(long)]
        skip_hidden: bool,
        /// Embed semantic images (default true), or use --embed-images=false for adjacent assets
        #[arg(long, num_args = 0..=1, default_missing_value = "true", require_equals = true)]
        embed_images: Option<bool>,
    },
    /// Compare slide text using a longest-common-subsequence diff
    Diff { file_a: PathBuf, file_b: PathBuf },
    /// Replace literal presentation text while retaining run formatting
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
        #[arg(long, short = 'v', requires = "placeholder", conflicts_with = "map")]
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
    /// Validate a stored presentation package through its original XML
    ///
    /// Checks a finite stored subset: XML 1.0 lexical and character faults,
    /// one complete root per XML part including package metadata, the
    /// p:presentation root child sequence, unknown direct shape-tree
    /// children, required unqualified cNvPr id and name on physical owners,
    /// table row against grid column equality, PowerPoint role content-type
    /// overrides, and the existing parsed-model issues. This is not a full
    /// schema validator.
    Validate { file: PathBuf },
    /// Render selected slides to deterministic image files
    Render {
        file: PathBuf,
        /// Directory containing caller font files (.ttf/.otf/.ttc)
        #[arg(long)]
        font_dir: Option<PathBuf>,
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
        /// Directory containing caller font files (.ttf/.otf/.ttc)
        #[arg(long)]
        font_dir: Option<PathBuf>,
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
    /// Apply slide-owned footer fields to the presentation
    Footer {
        file: PathBuf,
        #[arg(long)]
        slide_number: bool,
        #[arg(long)]
        footer: Option<String>,
        /// off, auto for an uncached automatic field, or fixed text
        #[arg(long, default_value = "off")]
        date: String,
        #[arg(long, default_value = "datetime1")]
        date_format: String,
        /// Explicit automatic date cache, without consulting the clock
        #[arg(long)]
        date_cache: Option<String>,
        #[arg(long)]
        skip_title: bool,
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,
        #[arg(long)]
        force: bool,
        #[arg(long)]
        json: bool,
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
    /// Write speaker notes without trimming or newline conversion
    Notes {
        #[command(subcommand)]
        command: NotesCommand,
    },
    /// Report every admitted text owner through the shared native fit solver
    Fit {
        file: PathBuf,
        #[arg(long)]
        json: bool,
        /// Directory containing caller fonts, otherwise deterministic bundled fonts
        #[arg(long)]
        font_dir: Option<PathBuf>,
    },
    /// Read and write core document properties
    Meta {
        #[command(subcommand)]
        command: MetaCommand,
    },
}

#[derive(Subcommand)]
enum SlideCommand {
    /// Import selected slides from another deck into a new output file
    Import {
        destination: PathBuf,
        source: PathBuf,
        /// One-based source selection in caller order, such as 1,3-5
        #[arg(long)]
        slides: Option<String>,
        /// One-based insertion position, append when absent
        #[arg(long)]
        at: Option<usize>,
        #[arg(long)]
        keep_source_formatting: bool,
        #[arg(long, short = 'o')]
        output: PathBuf,
        #[arg(long)]
        json: bool,
    },
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
    let signing_command = matches!(&cli.command, Command::Sign { .. });
    if let Command::VerifySignatures { file, json } = &cli.command {
        let result = commands::verify_signatures(file, *json);
        let flushed = io::stdout().flush();
        match (result, flushed) {
            (Ok(true), Ok(())) => return,
            (Ok(false), Ok(())) => process::exit(1),
            (Err(error), _) => {
                eprintln!("Error: {error}");
                process::exit(1);
            }
            (_, Err(error)) => {
                eprintln!("Error: {error}");
                process::exit(1);
            }
        }
    }

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

    if let Command::Fit {
        file,
        json,
        font_dir,
    } = &cli.command
    {
        match commands::fit(file, *json, font_dir.as_deref()) {
            Ok(false) => return,
            Ok(true) => process::exit(1),
            Err(error) => {
                eprintln!("Error: {error}");
                process::exit(2);
            }
        }
    }
    let result = match cli.command {
        Command::VerifySignatures { .. } => unreachable!("verification is dispatched above"),
        Command::Sign {
            file,
            key,
            cert,
            output,
            json,
        } => commands::sign(&file, &key, &cert, &output, json),
        Command::Font { command } => match command {
            FontCommand::List { file, json } => commands::font_list(&file, json),
            FontCommand::Embed {
                file,
                name,
                font,
                kind,
                license_authorized,
                license_identity,
                output,
                json,
            } => commands::font_embed(
                &file,
                &commands::FontEmbedOptions {
                    name: &name,
                    font: &font,
                    kind: &kind,
                    license_authorized,
                    license_identity: &license_identity,
                    output: &output,
                    json,
                },
            ),
            FontCommand::Remove {
                file,
                name,
                kind,
                output,
                json,
            } => commands::font_remove(&file, &name, kind.as_deref(), &output, json),
            FontCommand::Settings {
                file,
                embed_true_type_fonts,
                save_subset_fonts,
                output,
                json,
            } => commands::font_settings(
                &file,
                embed_true_type_fonts.as_deref(),
                save_subset_fonts.as_deref(),
                output.as_deref(),
                json,
            ),
        },

        Command::Inspect { file, json } => commands::inspect(&file, json),
        Command::Text { file, json, notes } => commands::text(&file, json, notes),
        Command::Convert {
            file,
            font_dir,
            to,
            output,
            force,
            dpi,
            slides,
            quality,
            transparent,
            notes,
            skip_hidden,
            embed_images,
        } => {
            if matches!(to.as_str(), "html" | "md") {
                if font_dir.is_some()
                    || dpi.is_some()
                    || slides.is_some()
                    || quality.is_some()
                    || transparent
                {
                    Err("semantic HTML and Markdown do not accept rendering options".into())
                } else {
                    commands::convert_semantic(
                        &file,
                        &to,
                        output.as_deref(),
                        force,
                        rpptx::PresentationTextExportOptions {
                            include_notes: notes,
                            skip_hidden,
                            inline_images: embed_images.unwrap_or(true),
                        },
                    )
                }
            } else if notes || skip_hidden || embed_images.is_some() {
                Err(
                    "--notes, --skip-hidden, and --embed-images require --to html or --to md"
                        .into(),
                )
            } else {
                commands::convert(
                    &file,
                    &to,
                    output.as_deref(),
                    force,
                    dpi.unwrap_or(150.0),
                    commands::ImageOptions {
                        font_dir: font_dir.as_deref(),
                        slides: slides.as_deref(),
                        quality: quality.unwrap_or(90),
                        transparent,
                    },
                )
            }
        }
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
        Command::Validate { .. } => unreachable!("validate is dispatched above"),
        Command::Render {
            file,
            font_dir,
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
                font_dir: font_dir.as_deref(),
                slides: slide.as_deref(),
                quality,
                transparent,
            },
        ),
        Command::Thumbnail {
            file,
            font_dir,
            output,
            force,
        } => commands::thumbnail(&file, output.as_deref(), force, font_dir.as_deref()),
        Command::Outline { file, json, notes } => commands::outline(&file, json, notes),
        Command::Footer {
            file,
            slide_number,
            footer,
            date,
            date_format,
            date_cache,
            skip_title,
            output,
            force,
            json,
        } => commands::footer(
            &file,
            commands::FooterInput {
                slide_number,
                footer: footer.as_deref(),
                date: &date,
                date_format: &date_format,
                date_cache: date_cache.as_deref(),
                skip_title,
            },
            output.as_deref(),
            force,
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
        Command::Slide { command } => match command {
            SlideCommand::Import {
                destination,
                source,
                slides,
                at,
                keep_source_formatting,
                output,
                json,
            } => commands::slide_import(
                &destination,
                &source,
                slides.as_deref(),
                at,
                keep_source_formatting,
                &output,
                json,
            ),
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
        Command::Fit { .. } => unreachable!("fit verdict dispatched above"),
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
        if !closed_stdout || signing_command {
            eprintln!("Error: {error}");
            process::exit(1);
        }
    }
}
