//! Shared command-line contracts for OOXML tools.

use std::collections::BTreeSet;
use std::fs::{self, File, OpenOptions};
use std::io::{self, Write};
use std::path::{Path, PathBuf};

use serde_json::Value;

const MAX_RANGE_VALUES: usize = 100_000;

/// An invalid shared command-line value.
#[derive(Clone, Debug, Eq, PartialEq, thiserror::Error)]
pub enum Error {
    /// A range expression or one of its components is invalid.
    #[error("invalid one-based range component: {0:?}")]
    InvalidRange(String),
    /// A range expression would materialize too many values.
    #[error("range selection exceeds the maximum of {limit} values")]
    RangeTooLarge { limit: usize },
    /// A JSON envelope payload was not an object.
    #[error("JSON payload must be an object")]
    JsonPayloadNotObject,
    /// A JSON envelope payload tried to define the reserved schema field.
    #[error("JSON payload must not define reserved field \"schema\"")]
    ReservedSchemaField,
    /// A replacement map is not a JSON array of replacement pairs.
    #[error("invalid replacement map: {0}")]
    InvalidReplacementMap(String),
}

/// A civil date and time on the wall clock.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct LocalDateTime {
    pub year: i32,
    pub month: u8,
    pub day: u8,
    pub hour: u8,
    pub minute: u8,
    pub second: u8,
}

/// The current local date and time, as Word reads it for a DATE field.
///
/// Unix asks the C library, which applies the `TZ` time zone. Elsewhere it
/// is the UTC time.
pub fn local_date_time() -> io::Result<LocalDateTime> {
    let invalid = |message: &str| io::Error::other(message.to_owned());
    let seconds = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_err(|_| invalid("the system clock is before 1970"))?
        .as_secs();
    #[cfg(unix)]
    {
        let time = libc::time_t::try_from(seconds).map_err(|_| invalid("the clock overflows"))?;
        // SAFETY: `localtime_r` writes only the `tm` it is given, which is
        // fully owned here and zero-initialisable plain data.
        let mut local: libc::tm = unsafe { std::mem::zeroed() };
        if unsafe { libc::localtime_r(&time, &mut local) }.is_null() {
            return Err(invalid("could not read the local time"));
        }
        let field = |value: libc::c_int| u8::try_from(value).map_err(|_| invalid("bad local time"));
        Ok(LocalDateTime {
            year: local.tm_year + 1900,
            month: field(local.tm_mon + 1)?,
            day: field(local.tm_mday)?,
            hour: field(local.tm_hour)?,
            minute: field(local.tm_min)?,
            // A leap second reads as the last second of its minute.
            second: field(local.tm_sec.min(59))?,
        })
    }
    #[cfg(not(unix))]
    {
        let days = i64::try_from(seconds / 86_400).map_err(|_| invalid("the clock overflows"))?;
        let (year, month, day) = civil_from_days(days);
        let second_of_day = seconds % 86_400;
        Ok(LocalDateTime {
            year: i32::try_from(year).map_err(|_| invalid("the clock overflows"))?,
            month,
            day,
            hour: (second_of_day / 3_600) as u8,
            minute: (second_of_day / 60 % 60) as u8,
            second: (second_of_day % 60) as u8,
        })
    }
}

/// The proleptic Gregorian date of a count of days since 1970-01-01.
pub fn civil_from_days(days: i64) -> (i64, u8, u8) {
    // Howard Hinnant's days-to-civil conversion.
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let day_of_era = z.rem_euclid(146_097);
    let year_of_era =
        (day_of_era - day_of_era / 1_460 + day_of_era / 36_524 - day_of_era / 146_096) / 365;
    let day_of_year = day_of_era - (365 * year_of_era + year_of_era / 4 - year_of_era / 100);
    let month_index = (5 * day_of_year + 2) / 153;
    let day = day_of_year - (153 * month_index + 2) / 5 + 1;
    let month = if month_index < 10 {
        month_index + 3
    } else {
        month_index - 9
    };
    let year = year_of_era + era * 400 + i64::from(month <= 2);
    (year, month as u8, day as u8)
}

/// The count of days since 1970-01-01 of a proleptic Gregorian date.
pub fn days_from_civil(year: i64, month: u8, day: u8) -> i64 {
    let year = year - i64::from(month <= 2);
    let era = year.div_euclid(400);
    let year_of_era = year.rem_euclid(400);
    let month = i64::from(month);
    let day_of_year =
        (153 * (if month > 2 { month - 3 } else { month + 9 }) + 2) / 5 + i64::from(day) - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    era * 146_097 + day_of_era - 719_468
}

/// One pair of a replacement map: the text to find, its replacement, and
/// the count the caller expects, when it gives one.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct ReplacementPair {
    pub placeholder: String,
    pub value: String,
    pub expect: Option<usize>,
}

/// Parses a replacement map, a JSON array of objects such as
/// `{"placeholder": "{{name}}", "value": "Ada", "expect": 2}`.
///
/// `expect` is optional. The pairs keep their order, because a later pair
/// sees the text an earlier one wrote. An empty array, a missing or
/// non-string `placeholder` or `value`, an empty `placeholder`, a negative or
/// fractional `expect`, and any other key are refused, naming the
/// zero-based pair.
pub fn parse_replacement_map(json: &str) -> Result<Vec<ReplacementPair>, Error> {
    let invalid = |message: String| Error::InvalidReplacementMap(message);
    let value: Value =
        serde_json::from_str(json).map_err(|error| invalid(format!("not JSON: {error}")))?;
    let Value::Array(entries) = value else {
        return Err(invalid(
            "expected an array of {\"placeholder\", \"value\", \"expect\"} objects".to_owned(),
        ));
    };
    if entries.is_empty() {
        return Err(invalid("the array holds no pair".to_owned()));
    }
    entries
        .iter()
        .enumerate()
        .map(|(index, entry)| {
            let Value::Object(fields) = entry else {
                return Err(invalid(format!("pair {index} is not an object")));
            };
            if let Some(key) = fields
                .keys()
                .find(|key| !matches!(key.as_str(), "placeholder" | "value" | "expect"))
            {
                return Err(invalid(format!(
                    "pair {index} has unknown key \"{key}\" (use placeholder, value, expect)"
                )));
            }
            let text = |key: &str| match fields.get(key) {
                Some(Value::String(text)) => Ok(text.clone()),
                _ => Err(invalid(format!("pair {index} needs a string \"{key}\""))),
            };
            let placeholder = text("placeholder")?;
            if placeholder.is_empty() {
                return Err(invalid(format!("pair {index} has an empty placeholder")));
            }
            let expect = match fields.get("expect") {
                None | Some(Value::Null) => None,
                Some(count) => Some(
                    count
                        .as_u64()
                        .and_then(|count| usize::try_from(count).ok())
                        .ok_or_else(|| {
                            invalid(format!(
                                "pair {index} needs a non-negative integer \"expect\""
                            ))
                        })?,
                ),
            };
            Ok(ReplacementPair {
                placeholder,
                value: text("value")?,
                expect,
            })
        })
        .collect()
}

/// Parses positive one-based values and inclusive ranges.
///
/// Components are comma-separated. Whitespace around components and range
/// endpoints is ignored. The result is sorted and deduplicated. At most
/// 100,000 values may be requested across all components before
/// deduplication.
pub fn parse_range(input: &str) -> Result<Vec<usize>, Error> {
    let mut values = BTreeSet::new();
    let mut expansion_work = 0;

    if input.trim().is_empty() {
        return Err(Error::InvalidRange(input.to_owned()));
    }

    for raw_component in input.split(',') {
        let component = raw_component.trim();
        if component.is_empty() {
            return Err(Error::InvalidRange(raw_component.to_owned()));
        }

        let hyphen_count = component.bytes().filter(|byte| *byte == b'-').count();
        match hyphen_count {
            0 => {
                let value = parse_positive(component)?;
                charge_expansion_work(&mut expansion_work, 1)?;
                values.insert(value);
            }
            1 => {
                let (start, end) = component
                    .split_once('-')
                    .expect("one counted hyphen must split");
                let start = parse_positive(start.trim())?;
                let end = parse_positive(end.trim())?;
                if start > end {
                    return Err(Error::InvalidRange(component.to_owned()));
                }
                let cardinality = end
                    .checked_sub(start)
                    .and_then(|width| width.checked_add(1))
                    .ok_or(Error::RangeTooLarge {
                        limit: MAX_RANGE_VALUES,
                    })?;
                charge_expansion_work(&mut expansion_work, cardinality)?;
                values.extend(start..=end);
            }
            _ => return Err(Error::InvalidRange(component.to_owned())),
        }
    }

    Ok(values.into_iter().collect())
}

fn charge_expansion_work(total: &mut usize, additional: usize) -> Result<(), Error> {
    let charged = total.checked_add(additional).ok_or(Error::RangeTooLarge {
        limit: MAX_RANGE_VALUES,
    })?;
    if charged > MAX_RANGE_VALUES {
        return Err(Error::RangeTooLarge {
            limit: MAX_RANGE_VALUES,
        });
    }
    *total = charged;
    Ok(())
}

fn parse_positive(value: &str) -> Result<usize, Error> {
    let parsed = value
        .parse::<usize>()
        .map_err(|_| Error::InvalidRange(value.to_owned()))?;
    if parsed == 0 {
        return Err(Error::InvalidRange(value.to_owned()));
    }
    Ok(parsed)
}

/// Replaces or adds the requested extension while preserving the input path.
pub fn default_output_path(input: &Path, extension: &str) -> PathBuf {
    let mut output = input.to_path_buf();
    output.set_extension(extension.trim_start_matches('.'));
    output
}

/// Fails before publication when any requested output already exists.
pub fn ensure_output_paths_available(paths: &[PathBuf]) -> io::Result<()> {
    let mut unique = BTreeSet::new();
    for path in paths {
        if !unique.insert(path.clone()) {
            return Err(io::Error::new(
                io::ErrorKind::AlreadyExists,
                format!("duplicate output path {}", path.display()),
            ));
        }
        if path.exists() {
            return Err(io::Error::new(
                io::ErrorKind::AlreadyExists,
                format!("output already exists: {}", path.display()),
            ));
        }
    }
    Ok(())
}

/// Fails before publication when a requested output is the input file, is not
/// a regular file, or already exists and `force` is false.
///
/// On Unix an output is the input when both name the same device and inode,
/// so a relative, `..`, symlinked or case-folded spelling of the input, a path
/// through another mount of its directory, and a hard link to it are refused
/// as well. Other platforms compare canonical paths.
///
/// An existing output that is not a regular file, such as a directory, a
/// symbolic link, a device, a FIFO or a socket, is refused even with `force`,
/// since a replacing publication would put a regular file in its place. An
/// output that does not exist yet cannot be the input.
pub fn ensure_output_paths_allowed(paths: &[PathBuf], input: &Path, force: bool) -> io::Result<()> {
    let input = file_identity(input)?;
    let mut unique = BTreeSet::new();
    for path in paths {
        if !unique.insert(path) {
            return Err(io::Error::new(
                io::ErrorKind::AlreadyExists,
                format!("duplicate output path {}", path.display()),
            ));
        }
        let Ok(metadata) = fs::symlink_metadata(path) else {
            continue;
        };
        if file_identity(path).is_ok_and(|output| output == input) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("output is the input file: {}", path.display()),
            ));
        }
        if !metadata.is_file() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                format!("output is not a regular file: {}", path.display()),
            ));
        }
        if !force {
            return Err(io::Error::new(
                io::ErrorKind::AlreadyExists,
                format!(
                    "output already exists: {} (pass --force to replace it)",
                    path.display()
                ),
            ));
        }
    }
    Ok(())
}

/// Identifies the file `path` names, whatever the spelling of `path`.
#[cfg(unix)]
fn file_identity(path: &Path) -> io::Result<(u64, u64)> {
    use std::os::unix::fs::MetadataExt;

    let metadata = fs::metadata(path)?;
    Ok((metadata.dev(), metadata.ino()))
}

/// Identifies the file `path` names, whatever the spelling of `path`.
#[cfg(not(unix))]
fn file_identity(path: &Path) -> io::Result<PathBuf> {
    fs::canonicalize(path)
}

/// Stages output files beside their final destinations and publishes them as a set.
pub struct StagedOutputSet {
    staged: Vec<StagedOutput>,
    replace_existing: bool,
}

struct StagedOutput {
    final_path: PathBuf,
    temp_path: PathBuf,
}

impl StagedOutputSet {
    pub fn new() -> Self {
        Self::with_replace_existing(false)
    }

    /// Creates a set that replaces existing destinations when
    /// `replace_existing` is true, and refuses them like `new` otherwise.
    ///
    /// Replacement renames each staged file over its destination, so a
    /// destination holds either its previous or its new complete content. If a
    /// later output fails, the outputs already published stay, since the files
    /// they replaced cannot be restored. Callers check the destinations with
    /// `ensure_output_paths_allowed` first, which refuses the input file and
    /// anything that is not a regular file.
    pub fn with_replace_existing(replace_existing: bool) -> Self {
        Self {
            staged: Vec::new(),
            replace_existing,
        }
    }

    pub fn stage_bytes(&mut self, final_path: &Path, bytes: &[u8]) -> io::Result<()> {
        if self
            .staged
            .iter()
            .any(|staged| staged.final_path == final_path)
        {
            return Err(io::Error::new(
                io::ErrorKind::AlreadyExists,
                format!("duplicate output path {}", final_path.display()),
            ));
        }
        if !self.replace_existing && final_path.exists() {
            return Err(io::Error::new(
                io::ErrorKind::AlreadyExists,
                format!("output already exists: {}", final_path.display()),
            ));
        }
        let (temp_path, mut file) = create_temp_file(final_path)?;
        if let Err(error) = file.write_all(bytes).and_then(|()| file.sync_all()) {
            let _ = fs::remove_file(&temp_path);
            return Err(error);
        }
        self.staged.push(StagedOutput {
            final_path: final_path.to_path_buf(),
            temp_path,
        });
        Ok(())
    }

    pub fn publish(mut self) -> io::Result<()> {
        let mut published = Vec::new();
        for staged in &self.staged {
            let result = if self.replace_existing {
                fs::rename(&staged.temp_path, &staged.final_path)
            } else {
                publish_staged_output(staged)
            };
            match result {
                Ok(()) => {
                    published.push(staged.final_path.clone());
                }
                Err(error) => {
                    // Removing a replaced output would not bring back what it
                    // replaced.
                    if !self.replace_existing {
                        for path in published {
                            let _ = fs::remove_file(path);
                        }
                    }
                    for staged in &self.staged {
                        let _ = fs::remove_file(&staged.temp_path);
                    }
                    self.staged.clear();
                    return Err(error);
                }
            }
        }
        for staged in &self.staged {
            let _ = fs::remove_file(&staged.temp_path);
        }
        self.staged.clear();
        Ok(())
    }
}

impl Default for StagedOutputSet {
    fn default() -> Self {
        Self::new()
    }
}

impl Drop for StagedOutputSet {
    fn drop(&mut self) {
        for staged in &self.staged {
            let _ = fs::remove_file(&staged.temp_path);
        }
    }
}

fn create_temp_file(final_path: &Path) -> io::Result<(PathBuf, File)> {
    let parent = final_path.parent().unwrap_or_else(|| Path::new("."));
    let file_name = final_path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("output");
    for attempt in 0..1000_u32 {
        let temp_path = parent.join(format!(
            ".{file_name}.{}.{}.tmp",
            std::process::id(),
            attempt
        ));
        match OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&temp_path)
        {
            Ok(file) => return Ok((temp_path, file)),
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => {}
            Err(error) => return Err(error),
        }
    }
    Err(io::Error::new(
        io::ErrorKind::AlreadyExists,
        format!(
            "could not create a unique temporary file beside {}",
            final_path.display()
        ),
    ))
}

fn publish_staged_output(staged: &StagedOutput) -> io::Result<()> {
    let mut input = File::open(&staged.temp_path)?;
    let mut output = OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&staged.final_path)?;
    if let Err(error) = io::copy(&mut input, &mut output)
        .and_then(|_| output.sync_all())
        .map(|_| ())
    {
        let _ = fs::remove_file(&staged.final_path);
        return Err(error);
    }
    fs::remove_file(&staged.temp_path)
}

/// Adds the versioned CLI schema field to an object payload.
pub fn json_envelope(mut payload: Value) -> Result<Value, Error> {
    let Value::Object(object) = &mut payload else {
        return Err(Error::JsonPayloadNotObject);
    };
    if object.contains_key("schema") {
        return Err(Error::ReservedSchemaField);
    }
    object.insert("schema".to_owned(), Value::from(1));
    Ok(payload)
}

#[cfg(test)]
mod tests {
    use std::fs;
    use std::path::{Path, PathBuf};

    use serde_json::json;

    use super::*;

    #[test]
    fn civil_dates_round_trip_through_day_counts() {
        assert_eq!(days_from_civil(1970, 1, 1), 0);
        assert_eq!(civil_from_days(0), (1970, 1, 1));
        assert_eq!(days_from_civil(2024, 2, 29), 19_782);
        for days in [-719_468, -1, 0, 59, 19_782, 2_932_896] {
            let (year, month, day) = civil_from_days(days);
            assert_eq!(days_from_civil(year, month, day), days);
        }
        let now = local_date_time().unwrap();
        assert!(now.year >= 2026 && (1..=12).contains(&now.month), "{now:?}");
    }

    #[test]
    fn replacement_maps_keep_pair_order_and_refuse_ambiguous_entries() {
        assert_eq!(
            parse_replacement_map(
                r#"[{"placeholder": "b", "value": "c", "expect": 2}, {"placeholder": "a", "value": "b"}]"#
            )
            .unwrap(),
            vec![
                ReplacementPair {
                    placeholder: "b".to_owned(),
                    value: "c".to_owned(),
                    expect: Some(2),
                },
                ReplacementPair {
                    placeholder: "a".to_owned(),
                    value: "b".to_owned(),
                    expect: None,
                },
            ]
        );
        for (json, message) in [
            ("{}", "expected an array"),
            ("[]", "holds no pair"),
            (
                r#"[{"placeholder": "a"}]"#,
                "pair 0 needs a string \"value\"",
            ),
            (
                r#"[{"placeholder": "", "value": "b"}]"#,
                "empty placeholder",
            ),
            (
                r#"[{"placeholder": "a", "value": "b", "expect": -1}]"#,
                "non-negative integer",
            ),
            (
                r#"[{"placeholder": "a", "value": "b", "count": 1}]"#,
                "unknown key \"count\"",
            ),
            ("[1]", "pair 0 is not an object"),
        ] {
            let error = parse_replacement_map(json).unwrap_err().to_string();
            assert!(error.contains(message), "{json}: {error}");
        }
    }

    fn temp_dir(label: &str) -> PathBuf {
        let temp = std::env::temp_dir().join(format!("oxml-cli-{label}-{}", std::process::id()));
        let _ = fs::remove_dir_all(&temp);
        fs::create_dir(&temp).unwrap();
        temp
    }

    fn temp_entries(path: &Path) -> Vec<PathBuf> {
        fs::read_dir(path)
            .unwrap()
            .map(|entry| entry.unwrap().path())
            .filter(|path| {
                path.file_name()
                    .and_then(|name| name.to_str())
                    .is_some_and(|name| name.ends_with(".tmp"))
            })
            .collect()
    }

    #[test]
    fn range_2_4_through_6_is_the_expected_set() {
        assert_eq!(parse_range("2,4-6").unwrap(), [2, 4, 5, 6]);
    }

    #[test]
    fn invalid_ranges_are_rejected_and_duplicates_are_normalized() {
        assert_eq!(parse_range("6, 2,4-6,2").unwrap(), [2, 4, 5, 6]);
        for invalid in ["", "0", "1,,2", "6-4", "2-", "-2", "two", "1-2-3"] {
            assert!(parse_range(invalid).is_err(), "accepted {invalid:?}");
        }
    }

    #[test]
    fn ranges_too_large_to_materialize_are_rejected() {
        let expected = Err(Error::RangeTooLarge { limit: 100_000 });
        assert_eq!(parse_range("1-100001"), expected);
        assert_eq!(parse_range(&format!("1-{}", usize::MAX)), expected);
    }

    #[test]
    fn exactly_one_hundred_thousand_requested_values_are_accepted() {
        let values = parse_range("1-100000").unwrap();
        assert_eq!(values.len(), 100_000);
        assert_eq!(values.first(), Some(&1));
        assert_eq!(values.last(), Some(&100_000));
    }

    #[test]
    fn overlapping_ranges_cannot_amplify_expansion_work() {
        assert_eq!(
            parse_range("1-50001,1-50001"),
            Err(Error::RangeTooLarge { limit: 100_000 })
        );
    }

    #[test]
    fn json_envelope_has_schema_one_and_preserves_payload_fields() {
        let value = json_envelope(json!({"slides": 3, "metadata": {"title": "Deck"}}))
            .expect("object payload");
        assert_eq!(value["schema"], 1);
        assert_eq!(value["slides"], 3);
        assert_eq!(value["metadata"]["title"], "Deck");
        assert!(json_envelope(json!({"schema": 9})).is_err());
        assert!(json_envelope(json!([1, 2, 3])).is_err());
    }

    #[test]
    fn output_paths_replace_or_add_only_the_extension() {
        assert_eq!(
            default_output_path(Path::new("relative/report.docx"), "pdf"),
            Path::new("relative/report.pdf")
        );
        assert_eq!(
            default_output_path(Path::new("relative/report"), ".html"),
            Path::new("relative/report.html")
        );
        assert_eq!(
            default_output_path(Path::new("relative/report.final.docx"), "md"),
            Path::new("relative/report.final.md")
        );
    }

    #[test]
    fn staged_outputs_roll_back_published_files_when_later_publication_fails() {
        let temp = temp_dir("staged-rollback");
        let first = temp.join("first.png");
        let second = temp.join("second.png");
        let mut staged = StagedOutputSet::new();
        staged.stage_bytes(&first, b"first").unwrap();
        staged.stage_bytes(&second, b"second").unwrap();
        fs::create_dir(&second).unwrap();

        let error = staged.publish().unwrap_err();
        assert_eq!(error.kind(), std::io::ErrorKind::AlreadyExists);
        assert!(!first.exists());
        assert!(second.is_dir());
        assert!(temp_entries(&temp).is_empty());
        fs::remove_dir_all(&temp).unwrap();
    }

    #[test]
    fn staged_outputs_leave_no_temps_on_success_and_reject_duplicate_targets() {
        let temp = temp_dir("staged-success");
        let first = temp.join("first.png");
        let second = temp.join("second.png");
        ensure_output_paths_available(&[first.clone(), second.clone()]).unwrap();
        assert_eq!(
            ensure_output_paths_available(&[first.clone(), first.clone()])
                .unwrap_err()
                .kind(),
            std::io::ErrorKind::AlreadyExists
        );

        let mut staged = StagedOutputSet::new();
        staged.stage_bytes(&first, b"first").unwrap();
        assert_eq!(
            staged.stage_bytes(&first, b"duplicate").unwrap_err().kind(),
            std::io::ErrorKind::AlreadyExists
        );
        staged.stage_bytes(&second, b"second").unwrap();
        staged.publish().unwrap();

        assert_eq!(fs::read(first).unwrap(), b"first");
        assert_eq!(fs::read(second).unwrap(), b"second");
        assert!(temp_entries(&temp).is_empty());
        fs::remove_dir_all(&temp).unwrap();
    }

    #[test]
    fn allowed_outputs_refuse_any_spelling_of_the_input_even_when_forced() {
        let temp = temp_dir("allowed-outputs");
        let input = temp.join("input.docx");
        let existing = temp.join("existing.pdf");
        let fresh = temp.join("fresh.pdf");
        fs::write(&input, b"input").unwrap();
        fs::write(&existing, b"existing").unwrap();
        fs::create_dir(temp.join("nested")).unwrap();
        #[cfg_attr(not(unix), allow(unused_mut))]
        let mut spellings = vec![input.clone(), temp.join("nested/../input.docx")];
        #[cfg(unix)]
        {
            let link = temp.join("link.docx");
            std::os::unix::fs::symlink(&input, &link).unwrap();
            spellings.push(link);
            // A hard link has its own canonical path, as a mount alias of the
            // input's directory does, so only the file identity matches it.
            let hard_link = temp.join("nested/hard-link.pdf");
            fs::hard_link(&input, &hard_link).unwrap();
            spellings.push(hard_link);
        }

        for force in [false, true] {
            for output in &spellings {
                let error =
                    ensure_output_paths_allowed(std::slice::from_ref(output), &input, force)
                        .unwrap_err();
                assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
                assert_eq!(
                    error.to_string(),
                    format!("output is the input file: {}", output.display())
                );
            }
            ensure_output_paths_allowed(std::slice::from_ref(&fresh), &input, force).unwrap();
            assert_eq!(
                ensure_output_paths_allowed(&[fresh.clone(), fresh.clone()], &input, force)
                    .unwrap_err()
                    .kind(),
                std::io::ErrorKind::AlreadyExists
            );
        }
        let error = ensure_output_paths_allowed(std::slice::from_ref(&existing), &input, false)
            .unwrap_err();
        assert_eq!(error.kind(), std::io::ErrorKind::AlreadyExists);
        assert_eq!(
            error.to_string(),
            format!(
                "output already exists: {} (pass --force to replace it)",
                existing.display()
            )
        );
        ensure_output_paths_allowed(&[existing], &input, true).unwrap();
        fs::remove_dir_all(&temp).unwrap();
    }

    #[test]
    fn allowed_outputs_refuse_anything_but_a_regular_file_even_when_forced() {
        let temp = temp_dir("allowed-special-outputs");
        let input = temp.join("input.docx");
        let folder = temp.join("folder.pdf");
        fs::write(&input, b"input").unwrap();
        fs::create_dir(&folder).unwrap();
        #[cfg_attr(not(unix), allow(unused_mut))]
        let mut outputs = vec![folder];
        #[cfg(unix)]
        {
            // Renaming over a link would replace the link, not its target.
            let target = temp.join("target.pdf");
            let link = temp.join("link.pdf");
            fs::write(&target, b"target").unwrap();
            std::os::unix::fs::symlink(&target, &link).unwrap();
            outputs.extend([link, PathBuf::from("/dev/null")]);
        }

        for force in [false, true] {
            for output in &outputs {
                let error =
                    ensure_output_paths_allowed(std::slice::from_ref(output), &input, force)
                        .unwrap_err();
                assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
                assert_eq!(
                    error.to_string(),
                    format!("output is not a regular file: {}", output.display())
                );
            }
        }
        fs::remove_dir_all(&temp).unwrap();
    }

    #[test]
    fn replacing_outputs_swap_in_complete_files_and_keep_them_when_a_later_one_fails() {
        let temp = temp_dir("staged-replace");
        let existing = temp.join("existing.png");
        let fresh = temp.join("fresh.png");
        fs::write(&existing, b"old").unwrap();
        assert_eq!(
            StagedOutputSet::new()
                .stage_bytes(&existing, b"new")
                .unwrap_err()
                .kind(),
            std::io::ErrorKind::AlreadyExists
        );

        let mut staged = StagedOutputSet::with_replace_existing(true);
        staged.stage_bytes(&existing, b"new").unwrap();
        staged.stage_bytes(&fresh, b"fresh").unwrap();
        staged.publish().unwrap();
        assert_eq!(fs::read(&existing).unwrap(), b"new");
        assert_eq!(fs::read(&fresh).unwrap(), b"fresh");
        assert!(temp_entries(&temp).is_empty());

        let blocked = temp.join("blocked.png");
        fs::create_dir(&blocked).unwrap();
        fs::write(blocked.join("keep"), b"keep").unwrap();
        let mut staged = StagedOutputSet::with_replace_existing(true);
        staged.stage_bytes(&existing, b"newer").unwrap();
        staged.stage_bytes(&blocked, b"blocked").unwrap();
        assert!(staged.publish().is_err());
        assert_eq!(fs::read(&existing).unwrap(), b"newer");
        assert_eq!(fs::read(blocked.join("keep")).unwrap(), b"keep");
        assert!(temp_entries(&temp).is_empty());
        fs::remove_dir_all(&temp).unwrap();
    }
}
