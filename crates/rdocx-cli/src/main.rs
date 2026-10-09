//! rdocx CLI — "jq for DOCX"
//!
//! Inspect, convert, diff, and manipulate DOCX files from the command line.

use std::io::{self, Write};
use std::path::PathBuf;
use std::process;

use clap::{Args, Parser, Subcommand};

mod commands;

#[derive(Parser)]
#[command(name = "rdocx", version, about = "CLI tool for DOCX files")]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    /// Print document structure: paragraph, table, word, character and page
    /// counts, styles, pictures, content controls, metadata
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
        /// Output accepted-view rich text as schema-1 JSON
        #[arg(long)]
        json: bool,
    },
    /// Inspect deterministic top-level body layout geometry
    Layout {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output point-space body fragments as schema-1 JSON
        #[arg(long)]
        json: bool,
    },
    /// Convert DOCX to another format (pdf, html, md, png, jpeg, tiff)
    Convert {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output format: pdf, html, md, png, jpeg, tiff
        #[arg(long, short = 't')]
        to: String,
        /// Output file path (defaults to input with new extension)
        #[arg(long, short = 'o')]
        output: Option<PathBuf>,
        /// Replace existing output files, but never the input file
        #[arg(long)]
        force: bool,
        /// DPI for image rendering (default: 150)
        #[arg(long, default_value = "150")]
        dpi: u32,
        /// Directory containing font files (.ttf/.otf) to use for PDF rendering
        #[arg(long)]
        font_dir: Option<PathBuf>,
        /// Revision view for PDF and image output
        #[arg(long, default_value = "accepted", value_parser = commands::parse_revision_view)]
        revision_view: rdocx::RevisionView,
        /// One-based page range for image output, such as 1,3-5
        #[arg(long)]
        pages: Option<String>,
        /// JPEG quality from 1 through 100
        #[arg(long, default_value = "90")]
        quality: u8,
        /// Preserve unpainted PNG pixels as transparent
        #[arg(long)]
        transparent: bool,
    },
    /// Compare the paragraph text of every story of two DOCX files
    ///
    /// Compares the body paragraphs, the table cells of the body, the text
    /// boxes, the headers and footers of each section, the footnotes, the
    /// endnotes, and the comments. A changed paragraph prints a `-` and a `+`
    /// line, each located between brackets: `[2]` for the second body
    /// paragraph, or a story location such as `[table 1, row 1, cell 2,
    /// paragraph 1]` or `[header default, section 1, paragraph 1]`. A story
    /// that cannot be read is named as not compared.
    Diff {
        /// First DOCX file
        file_a: PathBuf,
        /// Second DOCX file
        file_b: PathBuf,
        /// Output the differences as schema-1 JSON
        #[arg(long)]
        json: bool,
        /// Exit with 1 when the files differ and 2 on an error, as `diff` does
        #[arg(long)]
        exit_code: bool,
    },
    /// Replace placeholders in a DOCX file
    ///
    /// Give one pair with -p and -v, or many with --map. Nothing is written
    /// unless every pair with an expected count replaced exactly that count.
    Replace {
        /// Path to the DOCX file
        file: PathBuf,
        /// Placeholder string, a regular expression with --regex
        #[arg(
            long,
            short = 'p',
            required_unless_present = "map",
            requires = "value",
            conflicts_with = "map"
        )]
        placeholder: Option<String>,
        /// Replacement value, where $1 names a capture group with --regex
        #[arg(long, short = 'v', requires = "placeholder")]
        value: Option<String>,
        /// JSON file holding an array of pairs, applied in order, such as
        /// [{"placeholder": "{{name}}", "value": "Ada", "expect": 2}]
        #[arg(long, value_name = "PAIRS_JSON")]
        map: Option<PathBuf>,
        /// Read each placeholder as a regular expression
        #[arg(long)]
        regex: bool,
        /// Output file path
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Require exactly this many replacements before publishing output
        #[arg(long, conflicts_with = "map")]
        expect: Option<usize>,
        /// Output the count of each pair as JSON
        #[arg(long)]
        json: bool,
    },
    /// Validate OOXML conformance
    Validate {
        /// Path to the DOCX file
        file: PathBuf,
    },
    /// Render pages to image files
    Render {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output directory (defaults to current directory)
        #[arg(long, short = 'o')]
        output_dir: Option<PathBuf>,
        /// Replace existing page images, but never the input file
        #[arg(long)]
        force: bool,
        /// DPI resolution (default: 150)
        #[arg(long, default_value = "150")]
        dpi: f64,
        /// Render only a specific page (0-based index)
        #[arg(long, conflicts_with = "pages")]
        page: Option<usize>,
        /// One-based page range, such as 1,3-5
        #[arg(long)]
        pages: Option<String>,
        /// Output format: png, jpeg, tiff
        #[arg(long, default_value = "png")]
        format: String,
        /// Revision view to render
        #[arg(long, default_value = "accepted", value_parser = commands::parse_revision_view)]
        revision_view: rdocx::RevisionView,
        /// JPEG quality from 1 through 100
        #[arg(long, default_value = "90")]
        quality: u8,
        /// Preserve unpainted PNG pixels as transparent
        #[arg(long)]
        transparent: bool,
    },
    /// Inspect and mutate Word comment threads
    Comment {
        #[command(subcommand)]
        command: CommentCommand,
    },
    /// Inspect and resolve tracked Word revisions
    Revision {
        #[command(subcommand)]
        command: RevisionCommand,
    },
    /// Create a tracked-changes document from an original and edited file
    Compare {
        /// Original DOCX file
        original: PathBuf,
        /// Edited DOCX file
        edited: PathBuf,
        /// Revision author recorded in the comparison
        #[arg(long)]
        author: String,
        /// RFC 3339 revision timestamp
        #[arg(long)]
        timestamp: String,
        /// Unit of a text change: run, word, or character
        #[arg(
            long,
            value_name = "UNIT",
            default_value = "run",
            value_parser = commands::parse_comparison_granularity
        )]
        granularity: rdocx::ComparisonGranularity,
        /// Keep the original formatting and record no formatting change
        #[arg(long)]
        ignore_formatting: bool,
        /// Keep the original whitespace and record no whitespace-only change
        #[arg(long)]
        ignore_whitespace: bool,
        /// Keep the original field results and record no field change
        #[arg(long)]
        ignore_fields: bool,
        /// Keep the original comments and anchors, dropping the edited ones
        #[arg(long)]
        ignore_comments: bool,
        /// Keep one story of the original, repeatable: body, header, footer,
        /// comment, text_box, footnote, or endnote
        #[arg(
            long = "ignore-story",
            value_name = "KIND",
            value_parser = commands::parse_comparison_story
        )]
        ignore_stories: Vec<rdocx::ComparisonStoryKind>,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Rebuild supported fields
    Toc {
        #[command(subcommand)]
        command: TocCommand,
    },
    /// Update field results
    Fields {
        #[command(subcommand)]
        command: FieldsCommand,
    },
    /// List and extract the pictures of the main story
    Images {
        #[command(subcommand)]
        command: ImagesCommand,
    },
    /// Read and write core and custom document properties
    Meta {
        #[command(subcommand)]
        command: MetaCommand,
    },
    /// Set the value of content controls by tag or alias
    ///
    /// Every --tag and --alias must name at least one control of the body,
    /// or nothing is written. A control bound to custom XML gets the value in
    /// its bound part too.
    Fill {
        /// Path to the DOCX file
        file: PathBuf,
        /// Set every control whose tag is NAME, repeatable
        #[arg(long = "tag", value_name = "NAME=VALUE", value_parser = commands::parse_assignment)]
        tags: Vec<(String, String)>,
        /// Set every control whose alias (title) is NAME, repeatable
        #[arg(long = "alias", value_name = "NAME=VALUE", value_parser = commands::parse_assignment)]
        aliases: Vec<(String, String)>,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the count of each assignment as JSON
        #[arg(long)]
        json: bool,
    },
}

#[derive(Subcommand)]
enum FieldsCommand {
    /// Update PAGE, NUMPAGES, PAGEREF, SECTIONPAGES, DATE, TIME, SEQ, REF and
    /// the other supported field results
    ///
    /// Field results that need a value the document does not hold, such as a
    /// mail-merge field, keep their cached result. Page numbers come from
    /// rdocx's own pagination.
    Update {
        /// Path to the DOCX file
        file: PathBuf,
        /// Date and time for DATE and TIME fields, as YYYY-MM-DD or
        /// YYYY-MM-DDTHH:MM:SS, the current UTC time when absent
        #[arg(long, value_parser = commands::parse_field_date_time)]
        now: Option<rdocx::FieldDateTime>,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
}

#[derive(Subcommand)]
enum ImagesCommand {
    /// Write every picture of the main story into a directory and list them
    ///
    /// Pictures in headers, footers, notes and text boxes are not listed.
    Extract {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output directory, created when absent
        dir: PathBuf,
        /// Replace existing image files, but never the input file
        #[arg(long)]
        force: bool,
        /// Output the listing as JSON
        #[arg(long)]
        json: bool,
    },
}

// Parsed once per run, so the size of the `set` variant costs nothing.
#[allow(clippy::large_enum_variant)]
#[derive(Subcommand)]
enum MetaCommand {
    /// Print the core and custom properties
    Get {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output as JSON
        #[arg(long)]
        json: bool,
    },
    /// Set core properties and add, replace or remove custom properties
    Set {
        /// Path to the DOCX file
        file: PathBuf,
        #[command(flatten)]
        core: CoreArgs,
        /// Add or replace a custom property, repeatable. An existing number,
        /// integer or Boolean property keeps its type
        #[arg(long = "custom", value_name = "NAME=VALUE", value_parser = commands::parse_assignment)]
        custom: Vec<(String, String)>,
        /// Remove a custom property, repeatable
        #[arg(long = "remove-custom", value_name = "NAME")]
        remove_custom: Vec<String>,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the properties written as JSON
        #[arg(long)]
        json: bool,
    },
}

#[derive(Args)]
struct CoreArgs {
    /// Title
    #[arg(long)]
    title: Option<String>,
    /// Author (dc:creator)
    #[arg(long)]
    author: Option<String>,
    /// Subject
    #[arg(long)]
    subject: Option<String>,
    /// Keywords
    #[arg(long)]
    keywords: Option<String>,
    /// Description (comments)
    #[arg(long)]
    description: Option<String>,
    /// Category
    #[arg(long)]
    category: Option<String>,
}

#[derive(Subcommand)]
enum CommentCommand {
    /// List comments in package order
    List {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output as JSON
        #[arg(long)]
        json: bool,
    },
    /// Add a comment over a zero-based half-open body run range, or on a
    /// piece of text with --anchor
    Add {
        /// Path to the DOCX file
        file: PathBuf,
        #[command(flatten)]
        range: CommentRangeArgs,
        /// Anchor the comment on this literal, case-sensitive text of the
        /// main story instead of a run range
        #[arg(
            long,
            conflicts_with_all = ["start_paragraph", "start_run", "end_paragraph", "end_run"]
        )]
        anchor: Option<String>,
        /// Zero-based occurrence of the --anchor text in document order,
        /// 0 when absent
        #[arg(
            long,
            requires = "anchor",
            conflicts_with_all = ["start_paragraph", "start_run", "end_paragraph", "end_run"]
        )]
        occurrence: Option<usize>,
        /// Comment author
        #[arg(long)]
        author: String,
        /// Optional comment author initials
        #[arg(long)]
        initials: Option<String>,
        /// Comment text, one paragraph per line
        #[arg(long)]
        text: String,
        /// RFC 3339 comment timestamp, omitted from the comment when absent
        #[arg(long)]
        date: Option<String>,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Move an existing root comment onto literal main-story text
    Move {
        /// Path to the DOCX file
        file: PathBuf,
        /// Existing root comment id
        id: i32,
        /// Literal case-sensitive destination text
        #[arg(long)]
        text: String,
        /// Zero-based destination occurrence
        #[arg(long, default_value_t = 0)]
        occurrence: usize,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Reply to an existing comment
    Reply {
        /// Path to the DOCX file
        file: PathBuf,
        /// Parent comment id
        #[arg(long)]
        id: i32,
        /// Reply author
        #[arg(long)]
        author: String,
        /// Reply text, one paragraph per line
        #[arg(long)]
        text: String,
        /// RFC 3339 reply timestamp, omitted from the reply when absent
        #[arg(long)]
        date: Option<String>,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Mark one comment thread resolved
    Resolve {
        /// Path to the DOCX file
        file: PathBuf,
        /// Comment id
        #[arg(long)]
        id: i32,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Remove one comment and its replies
    Remove {
        /// Path to the DOCX file
        file: PathBuf,
        /// Comment id
        #[arg(long)]
        id: i32,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
}

#[derive(Args)]
struct CommentRangeArgs {
    /// Zero-based body paragraph index at the inclusive start
    #[arg(long, required_unless_present = "anchor")]
    start_paragraph: Option<usize>,
    /// Zero-based run boundary at the inclusive start, counting the runs that
    /// `text --json` lists
    #[arg(long, required_unless_present = "anchor")]
    start_run: Option<usize>,
    /// Zero-based body paragraph index at the exclusive end
    #[arg(long, required_unless_present = "anchor")]
    end_paragraph: Option<usize>,
    /// Zero-based run boundary at the exclusive end, counting the runs that
    /// `text --json` lists
    #[arg(long, required_unless_present = "anchor")]
    end_run: Option<usize>,
}

#[derive(Subcommand)]
enum RevisionCommand {
    /// List modeled revisions from every supported story
    List {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output as JSON
        #[arg(long)]
        json: bool,
    },
    /// Accept revisions from every supported story
    Accept {
        /// Path to the DOCX file
        file: PathBuf,
        #[command(flatten)]
        selector: RevisionSelectorArgs,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
    /// Reject revisions from every supported story
    Reject {
        /// Path to the DOCX file
        file: PathBuf,
        #[command(flatten)]
        selector: RevisionSelectorArgs,
        /// Output DOCX file
        #[arg(long, short = 'o')]
        output: PathBuf,
        /// Output the operation record as JSON
        #[arg(long)]
        json: bool,
    },
}

#[derive(Args)]
struct RevisionSelectorArgs {
    /// Select one shared revision id
    #[arg(long, conflicts_with_all = ["author", "start_date", "end_date"])]
    id: Option<i32>,
    /// Select revisions by exact case-sensitive author
    #[arg(long, conflicts_with_all = ["id", "start_date", "end_date"])]
    author: Option<String>,
    /// Inclusive RFC 3339 lower date bound
    #[arg(long, requires = "end_date", conflicts_with_all = ["id", "author"])]
    start_date: Option<String>,
    /// Inclusive RFC 3339 upper date bound
    #[arg(long, requires = "start_date", conflicts_with_all = ["id", "author"])]
    end_date: Option<String>,
}

#[derive(Subcommand)]
enum TocCommand {
    /// Rebuild supported existing table-of-contents fields
    Rebuild {
        /// Path to the DOCX file
        file: PathBuf,
        /// Output DOCX file
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
        // Full document comparison can exceed the one MiB Windows main-thread stack.
        std::thread::Builder::new()
            .stack_size(8 * 1024 * 1024)
            .spawn(run_cli)
            .expect("start rdocx CLI thread")
            .join()
            .expect("rdocx CLI thread panicked");
    }
    #[cfg(not(windows))]
    run_cli();
}

fn run_cli() {
    let cli = Cli::parse();

    // `validate` always carries a verdict in its exit status, so it is
    // dispatched separately from the commands that only report errors.
    if let Command::Validate { file } = &cli.command {
        match commands::validate(file) {
            Ok(true) => return,
            Ok(false) => process::exit(1),
            Err(e) => {
                eprintln!("Error: {e}");
                process::exit(1);
            }
        }
    }

    // `diff --exit-code` carries a verdict too: 1 for files that differ, which
    // survives a closed standard output, and 2 for an error.
    let diff_exit_code = matches!(
        cli.command,
        Command::Diff {
            exit_code: true,
            ..
        }
    );
    let mut diff_status = 0;

    let result = match cli.command {
        Command::Inspect { file, json } => commands::inspect(&file, json),
        Command::Text { file, json } => commands::text(&file, json),
        Command::Layout { file, json } => commands::layout(&file, json),
        Command::Convert {
            file,
            to,
            output,
            force,
            dpi,
            font_dir,
            revision_view,
            pages,
            quality,
            transparent,
        } => commands::convert(
            &file,
            &to,
            output.as_deref(),
            force,
            dpi,
            font_dir.as_deref(),
            revision_view,
            commands::ImageOptions {
                pages: pages.as_deref(),
                quality,
                transparent,
            },
        ),
        Command::Diff {
            file_a,
            file_b,
            json,
            exit_code: _,
        } => commands::diff(&file_a, &file_b, json, &mut diff_status),
        Command::Replace {
            file,
            placeholder,
            value,
            map,
            regex,
            output,
            expect,
            json,
        } => commands::replace(
            &file,
            commands::ReplaceInput {
                pair: placeholder.zip(value),
                map: map.as_deref(),
                expect,
                regex,
            },
            &output,
            json,
        ),
        // Handled above so its exit code can reflect the verdict.
        Command::Validate { .. } => unreachable!(),
        Command::Render {
            file,
            output_dir,
            force,
            dpi,
            page,
            pages,
            format,
            revision_view,
            quality,
            transparent,
        } => commands::render(
            &file,
            output_dir.as_deref(),
            force,
            dpi,
            commands::RenderOptions {
                page,
                pages: pages.as_deref(),
                format: &format,
                revision_view,
                quality,
                transparent,
            },
        ),
        Command::Comment { command } => match command {
            CommentCommand::List { file, json } => commands::comment_list(&file, json),
            CommentCommand::Add {
                file,
                range,
                anchor,
                occurrence,
                author,
                initials,
                text,
                date,
                output,
                json,
            } => commands::comment_add(
                &file,
                match (
                    anchor.as_deref(),
                    range.start_paragraph,
                    range.start_run,
                    range.end_paragraph,
                    range.end_run,
                ) {
                    (Some(anchor), ..) => commands::CommentAnchor::Text {
                        text: anchor,
                        occurrence: occurrence.unwrap_or(0),
                    },
                    (
                        None,
                        Some(start_paragraph),
                        Some(start_run),
                        Some(end_paragraph),
                        Some(end_run),
                    ) => commands::CommentAnchor::Range(rdocx::RunRange {
                        start: rdocx::RunPosition {
                            body_index: start_paragraph,
                            run_index: start_run,
                        },
                        end: rdocx::RunPosition {
                            body_index: end_paragraph,
                            run_index: end_run,
                        },
                    }),
                    (None, ..) => unreachable!("clap requires the range without --anchor"),
                },
                &author,
                initials.as_deref(),
                &text,
                date.as_deref(),
                &output,
                json,
            ),
            CommentCommand::Move {
                file,
                id,
                text,
                occurrence,
                output,
                json,
            } => commands::comment_move(&file, id, &text, occurrence, &output, json),
            CommentCommand::Reply {
                file,
                id,
                author,
                text,
                date,
                output,
                json,
            } => commands::comment_reply(&file, id, &author, &text, date.as_deref(), &output, json),
            CommentCommand::Resolve {
                file,
                id,
                output,
                json,
            } => commands::comment_resolve(&file, id, &output, json),
            CommentCommand::Remove {
                file,
                id,
                output,
                json,
            } => commands::comment_remove(&file, id, &output, json),
        },
        Command::Revision { command } => match command {
            RevisionCommand::List { file, json } => commands::revision_list(&file, json),
            RevisionCommand::Accept {
                file,
                selector,
                output,
                json,
            } => commands::resolve_revisions(
                &file,
                commands::RevisionAction::Accept,
                commands::RevisionSelector {
                    id: selector.id,
                    author: selector.author.as_deref(),
                    start_date: selector.start_date.as_deref(),
                    end_date: selector.end_date.as_deref(),
                },
                &output,
                json,
            ),
            RevisionCommand::Reject {
                file,
                selector,
                output,
                json,
            } => commands::resolve_revisions(
                &file,
                commands::RevisionAction::Reject,
                commands::RevisionSelector {
                    id: selector.id,
                    author: selector.author.as_deref(),
                    start_date: selector.start_date.as_deref(),
                    end_date: selector.end_date.as_deref(),
                },
                &output,
                json,
            ),
        },
        Command::Compare {
            original,
            edited,
            author,
            timestamp,
            granularity,
            ignore_formatting,
            ignore_whitespace,
            ignore_fields,
            ignore_comments,
            ignore_stories,
            output,
            json,
        } => commands::compare(
            &original,
            &edited,
            &author,
            &timestamp,
            &rdocx::ComparisonOptions {
                granularity,
                ignore_formatting,
                ignore_whitespace,
                ignore_fields,
                ignore_comments,
                ignored_stories: ignore_stories,
            },
            &output,
            json,
        ),
        Command::Toc { command } => match command {
            TocCommand::Rebuild { file, output, json } => {
                commands::toc_rebuild(&file, &output, json)
            }
        },
        Command::Fields { command } => match command {
            FieldsCommand::Update {
                file,
                now,
                output,
                json,
            } => commands::fields_update(&file, now, &output, json),
        },
        Command::Images { command } => match command {
            ImagesCommand::Extract {
                file,
                dir,
                force,
                json,
            } => commands::images_extract(&file, &dir, force, json),
        },
        Command::Meta { command } => match command {
            MetaCommand::Get { file, json } => commands::meta_get(&file, json),
            MetaCommand::Set {
                file,
                core,
                custom,
                remove_custom,
                output,
                json,
            } => commands::meta_set(
                &file,
                &commands::MetaChanges {
                    title: core.title,
                    author: core.author,
                    subject: core.subject,
                    keywords: core.keywords,
                    description: core.description,
                    category: core.category,
                    custom,
                    remove_custom,
                },
                &output,
                json,
            ),
        },
        Command::Fill {
            file,
            tags,
            aliases,
            output,
            json,
        } => commands::fill(&file, &tags, &aliases, &output, json),
    };

    // Standard output is line buffered, so a last line without a newline is
    // only written, and can only fail, when it is flushed.
    let result = result.and_then(|()| io::stdout().flush().map_err(Into::into));
    if let Err(e) = result {
        // A reader that closes standard output early, as `| head` does, ends
        // the output. That is not a failure of the command.
        let closed_stdout = e
            .downcast_ref::<io::Error>()
            .is_some_and(|e| e.kind() == io::ErrorKind::BrokenPipe);
        if !closed_stdout {
            eprintln!("Error: {e}");
            process::exit(if diff_exit_code { 2 } else { 1 });
        }
    }
    if diff_exit_code && diff_status != 0 {
        process::exit(diff_status);
    }
}
