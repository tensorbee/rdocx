#!/usr/bin/env python3
"""Sprint state machine.

Holds the resumable state of a sprint in `.claude/scratch/SNN-run.json`, and
enforces the transitions that `.claude/WORKFLOW.md` describes as law. The point
is that "an agent may not merge" is checkable rather than aspirational.

Subcommands:
    init SNN                     create the run state from CURRENT_SPRINT.md
    status [SNN]                 print the current state
    set-phase SNN PHASE          advance the sprint phase
    mark-feature SNN F-ID STATE  advance one feature's state and its worker facts
    record-review SNN N ...      record a sprint-review pass result
    record-verification SNN ...  record a /verify result
    validate-handoff PATH ...    check a worker handoff before integration
    close-preflight SNN          the checks /close-sprint requires
    release-notes TAG            validate or render reviewed release notes
    unified-release-family TAG      validate matching Rust and Python versions
    release-semver all|TAG ...    execute or consume strict native compatibility proof
    published-unified-release TAG DIR
                                 verify exact selected registry entries
    python-release-artifacts TAG DIR
                                 validate one selected Python release set

Exit codes: 0 ok, 1 refused, 2 usage.
"""

from __future__ import annotations

import argparse
import hashlib
import gzip
import shutil
import tempfile
import platform
from datetime import datetime, timezone
from email.message import Message
from email.parser import Parser
import json
import os
import re
import signal
import stat
import string
import subprocess
import sys
import tarfile
import time
import urllib.error
import urllib.request
import urllib.parse
import tomllib
import unicodedata
import zipfile
from html.parser import HTMLParser
from pathlib import Path, PurePosixPath, PureWindowsPath

try:
    from . import prose_check
except ImportError:
    import prose_check

SCHEMA_VERSION = 2

REPO = Path(__file__).resolve().parent.parent
SCRATCH = REPO / ".claude" / "scratch"
HANDOFFS = REPO / ".claude" / "handoffs"
CURRENT_SPRINT = REPO / "docs" / "sprints" / "CURRENT_SPRINT.md"
BACKLOG = REPO / "docs" / "sprints" / "BACKLOG.md"
SPRINT_TRACKER = REPO / "docs" / "sprints" / "SPRINT_TRACKER.md"
AS_BUILT = REPO / "docs" / "sprints" / "AS_BUILT.md"
PLANS = REPO / ".claude" / "plans"
REVIEWS = REPO / ".claude" / "reviews"
CHANGELOG = REPO / "CHANGELOG.md"

SPRINT_RE = re.compile(r"^# Current Sprint, (S\d+(?:\.\d+)?)$", re.MULTILINE)
VALIDATION_ONLY_RE = re.compile(
    r"^\*\*Validation-only\*\*:\s*yes\s*$", re.MULTILINE | re.IGNORECASE
)
WAVE_ROW_RE = re.compile(
    r"^\|\s*(F-[A-Za-z0-9]+)\s*\|\s*([^|]+?)\s*\|\s*([SMLX]+)\s*\|\s*"
    r"(pending|in-progress|done|archived)\s*\|\s*([^|]*?)\s*\|$",
    re.MULTILINE,
)
BACKLOG_ROW_RE = re.compile(
    r"^\|\s*(F-[A-Za-z0-9]+)\s*\|[^|]*\|[^|]*\|\s*([SMLX]+)\s*\|\s*"
    r"(pending|in-progress|done|archived)\s*\|$",
    re.MULTILINE,
)
FID_RE = re.compile(r"^F-[A-Za-z0-9]+$")
SPRINT_ID_RE = re.compile(r"^S\d+(?:\.\d+)?$")
HANDOFF_FIELD_RE = re.compile(r"^\*\*([A-Za-z][A-Za-z -]*)\*\*:\s*(.+?)\s*$", re.MULTILINE)
SEMVER_COMPONENT_RE = r"(?:0|[1-9][0-9]*)"
RELEASE_TAG_RE = re.compile(
    rf"^(?:(?:rpptx|py-(?:rdocx|rpptx))-)?v{SEMVER_COMPONENT_RE}\."
    rf"{SEMVER_COMPONENT_RE}\.{SEMVER_COMPONENT_RE}$"
)
PYTHON_RELEASE_TAG_RE = re.compile(
    rf"^(?:(?P<stable>v)|(?P<incubating>rpptx-v)|py-(?P<legacy>rdocx|rpptx)-v)(?P<version>{SEMVER_COMPONENT_RE}\."
    rf"{SEMVER_COMPONENT_RE}\.{SEMVER_COMPONENT_RE})$"
)
PYTHON_RELEASE_PLATFORMS = (
    "manylinux_2_28_x86_64",
    "manylinux_2_28_aarch64",
    "musllinux_1_2_x86_64",
    "macosx_10_12_x86_64",
    "macosx_11_0_arm64",
    "win_amd64",
)
PYTHON_RELEASE_METADATA = {
    "rdocx": {
        "readme": "crates/rdocx-py/README.md",
        "summary": (
            "Native DOCX creation, editing, comparison, layout, and "
            "rendering for Python"
        ),
        "keywords": ("docx", "word", "ooxml", "documents", "pdf"),
    },
    "rpptx": {
        "readme": "crates/rpptx-py/README.md",
        "summary": (
            "Native PPTX creation, editing, comments, notes, and rendering "
            "for Python"
        ),
        "keywords": ("pptx", "powerpoint", "ooxml", "presentations", "pdf"),
    },
}
PYTHON_RELEASE_CLASSIFIERS = {
    "Development Status :: 4 - Beta",
    "Intended Audience :: Developers",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3 :: Only",
    "Programming Language :: Python :: Implementation :: CPython",
    "Programming Language :: Rust",
    "Topic :: Office/Business",
    "Topic :: Software Development :: Libraries",
}
PYTHON_RELEASE_PROJECT_URLS = {
    "Homepage, https://github.com/tensorbee/rdocx",
    "Repository, https://github.com/tensorbee/rdocx",
    "Issues, https://github.com/tensorbee/rdocx/issues",
    "Changelog, https://github.com/tensorbee/rdocx/blob/main/CHANGELOG.md",
}
RELEASE_NOTE_PLACEHOLDER_RE = re.compile(
    r"\b(?:TBD|TODO|FIXME|CHANGEME|PLACEHOLDER)\b|\?\?\?|\[insert\b",
    re.IGNORECASE,
)
RELEASE_NOTE_HEADINGS = (
    "Highlights",
    "Added",
    "Fixed",
    "Compatibility",
    "Contributors",
)
COMMONMARK_BLOCK_TAGS = (
    "address|article|aside|base|basefont|blockquote|body|caption|center|col|"
    "colgroup|dd|details|dialog|dir|div|dl|dt|fieldset|figcaption|figure|footer|"
    "form|frame|frameset|h1|h2|h3|h4|h5|h6|head|header|hr|html|iframe|legend|"
    "li|link|main|menu|menuitem|nav|noframes|ol|optgroup|option|p|param|search|"
    "section|summary|table|tbody|td|tfoot|th|thead|title|tr|track|ul"
)
COMMONMARK_BLOCK_TAG_RE = re.compile(
    rf"^ {{0,3}}</?(?:{COMMONMARK_BLOCK_TAGS})(?:[ \t/>]|$)",
    re.IGNORECASE,
)
COMMONMARK_COMPLETE_TAG_RE = re.compile(
    r"^ {0,3}(?:"
    r"</[A-Za-z][A-Za-z0-9-]*[ \t]*>|"
    r"<[A-Za-z][A-Za-z0-9-]*(?:[ \t]+[^<>]*)?[ \t]*/?>"
    r")[ \t]*$"
)
COMMONMARK_BOUNDED_RAW_HTML = (
    (re.compile(r"^ {0,3}<\?"), re.compile(r"\?>")),
    (re.compile(r"^ {0,3}<![A-Z]"), re.compile(r">")),
    (re.compile(r"^ {0,3}<!\[CDATA\["), re.compile(r"\]\]>")),
)
HTML_VOID_ELEMENTS = frozenset(
    (
        "area",
        "base",
        "br",
        "col",
        "embed",
        "hr",
        "img",
        "input",
        "link",
        "meta",
        "param",
        "source",
        "track",
        "wbr",
    )
)
COMMONMARK_REFERENCE_DEFINITION_RE = re.compile(
    r"^ {0,3}\[(?:\\.|[^\[\]\r\n]){1,999}\]:"
)
COMMONMARK_REFERENCE_TITLE_RE = re.compile(
    r"^ {0,3}(?:"
    r'"(?:\\.|[^"\r\n])*"|'
    r"'(?:\\.|[^'\r\n])*'|"
    r"\((?:\\.|[^)\r\n])*\)"
    r")[ \t]*$"
)
COMMONMARK_AUTOLINK_RE = re.compile(
    r"<(?P<label>"
    r"(?:[A-Za-z][A-Za-z0-9+.-]{1,31}:[^<>\s]*|"
    r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[^<>\s@]+)"
    r")>"
)
COMMONMARK_ESCAPABLE_PUNCTUATION = frozenset(string.punctuation)
MARKDOWN_ESCAPE_GUARD = "\N{WORD JOINER}"

# Every field a worker must have filled in before /integrate-feature will look
# at its branch. A missing one means the worker stopped early, and integrating
# on the strength of a half-written handoff is how unreviewed code lands.
HANDOFF_REQUIRED = (
    "F-ID",
    "Owner",
    "Branch",
    "Worktree",
    "Base",
    "Head",
    "Design plan",
    "Microscope",
    "Verify",
    "Hash harness",
    "Test gate",
)

PHASES = (
    "design",
    "questions",
    "implementation",
    "integration",
    "verification",
    "review",
    "ready_to_close",
    "blocked",
    "closed",
)

FEATURE_STATES = (
    "pending",
    "draft",
    "approved",
    "running",
    "reviewed",
    "completed",
    "blocked",
    "carried",
)

FEATURE_TRANSITIONS: dict[str, set[str]] = {
    "pending": {"draft", "approved", "blocked", "carried"},
    "draft": {"approved", "blocked", "carried"},
    "approved": {"running", "blocked", "carried"},
    "running": {"reviewed", "blocked", "carried"},
    "reviewed": {"running", "completed", "blocked"},
    "completed": set(),
    "blocked": {"draft", "approved", "running", "reviewed", "completed", "carried"},
    "carried": set(),
}

MAX_REVIEW_PASSES = 3

# Worker facts `mark-feature` records alongside the state. The CLI flag is the
# lowercase name with hyphens, so `--integration-commit` writes
# `integration_commit`.
WORKER_FIELDS = (
    "wave",
    "branch",
    "worktree",
    "base",
    "head",
    "handoff",
    "integration_commit",
)


def die(msg: str) -> None:
    print(f"sprint_workflow: {msg}", file=sys.stderr)
    sys.exit(1)


def protect_markdown_escapes(source: str) -> str:
    """Protect visible CommonMark escapes before classifying raw HTML."""
    protected: list[str] = []
    index = 0
    while index < len(source):
        if (
            source[index] == "\\"
            and index + 1 < len(source)
            and source[index + 1] in COMMONMARK_ESCAPABLE_PUNCTUATION
        ):
            # The format-only guard cannot collide with Markdown syntax and
            # does not make an otherwise empty section meaningful.
            protected.append(MARKDOWN_ESCAPE_GUARD)
            index += 2
            continue
        protected.append(source[index])
        index += 1
    return "".join(protected)


def markdown_heading_lines(lines: list[str], prefix: str) -> list[tuple[int, str]]:
    """Return exact Markdown headings outside non-rendered Markdown contexts."""
    headings: list[tuple[int, str]] = []
    fence: tuple[str, int] | None = None
    in_comment = False
    raw_html_end: re.Pattern[str] | None = None
    raw_html_until_blank = False
    for index, line in enumerate(lines):
        without_newline = protect_markdown_escapes(line.rstrip("\r\n"))
        if fence is not None:
            marker, minimum_length = fence
            close = re.fullmatch(rf" {{0,3}}({re.escape(marker)}+)[ \t]*", without_newline)
            if close is not None and len(close.group(1)) >= minimum_length:
                fence = None
            continue
        if raw_html_end is not None:
            if raw_html_end.search(without_newline):
                raw_html_end = None
            continue
        if raw_html_until_blank:
            if not without_newline.strip():
                raw_html_until_blank = False
            continue

        visible: list[str] = []
        cursor = 0
        while cursor < len(without_newline):
            if in_comment:
                end = without_newline.find("-->", cursor)
                if end < 0:
                    visible.append(" " * (len(without_newline) - cursor))
                    cursor = len(without_newline)
                    continue
                visible.append(" " * (end + 3 - cursor))
                cursor = end + 3
                in_comment = False
                continue
            start = without_newline.find("<!--", cursor)
            if start < 0:
                visible.append(without_newline[cursor:])
                break
            visible.append(without_newline[cursor:start])
            visible.append(" " * 4)
            cursor = start + 4
            in_comment = True

        visible_line = "".join(visible)
        opening = re.match(r"^ {0,3}(`{3,}|~{3,})", visible_line)
        if opening is not None:
            marker = opening.group(1)
            fence = (marker[0], len(marker))
            continue
        raw_html = re.match(
            r"^ {0,3}<(?P<tag>script|pre|style|textarea)(?:[ \t>]|$)",
            visible_line,
            re.IGNORECASE,
        )
        if raw_html is not None:
            tag = raw_html.group("tag").lower()
            close = re.compile(rf"</{re.escape(tag)}\s*>", re.IGNORECASE)
            if close.search(visible_line[raw_html.end() :]) is None:
                raw_html_end = close
            continue
        matched_bounded_raw_html = False
        for start_pattern, close in COMMONMARK_BOUNDED_RAW_HTML:
            start = start_pattern.match(visible_line)
            if start is None:
                continue
            if close.search(visible_line[start.end() :]) is None:
                raw_html_end = close
            matched_bounded_raw_html = True
            break
        if matched_bounded_raw_html:
            continue
        if COMMONMARK_BLOCK_TAG_RE.match(visible_line) is not None:
            raw_html_until_blank = True
            continue
        if COMMONMARK_COMPLETE_TAG_RE.fullmatch(visible_line) is not None:
            raw_html_until_blank = True
            continue
        if visible_line.startswith(prefix):
            headings.append((index, visible_line.rstrip()[len(prefix) :]))
    return headings


class MarkdownOutsideRawHtml(HTMLParser):
    """Collect Markdown source that is not contained by a raw HTML element."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.hidden_tags: list[str] = []
        self.text: list[str] = []

    def handle_starttag(
        self, tag: str, attrs: list[tuple[str, str | None]]
    ) -> None:
        del attrs
        tag = tag.lower()
        if tag not in HTML_VOID_ELEMENTS:
            self.hidden_tags.append(tag)

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        for index in range(len(self.hidden_tags) - 1, -1, -1):
            if self.hidden_tags[index] == tag:
                del self.hidden_tags[index:]
                break

    def handle_data(self, data: str) -> None:
        if not self.hidden_tags:
            self.text.append(data)


def closing_markdown_delimiter(
    source: str, opening_index: int, opening: str, closing: str
) -> int | None:
    """Find one escaped, balanced Markdown delimiter within bounded source."""
    depth = 1
    index = opening_index + 1
    while index < len(source):
        character = source[index]
        if character == "\\":
            index += 2
            continue
        if character == opening:
            depth += 1
        elif character == closing:
            depth -= 1
            if depth == 0:
                return index
        index += 1
    return None


def markdown_without_link_destinations(source: str) -> str:
    """Keep rendered labels while removing inline and reference destinations."""
    visible: list[str] = []
    index = 0
    while index < len(source):
        if source[index] == "\\" and index + 1 < len(source):
            visible.append(source[index : index + 2])
            index += 2
            continue
        image = source[index] == "!" and source[index + 1 : index + 2] == "["
        label_start = index + 1 if image else index
        if source[label_start : label_start + 1] != "[":
            visible.append(source[index])
            index += 1
            continue
        label_end = closing_markdown_delimiter(source, label_start, "[", "]")
        if label_end is None:
            visible.append(source[index])
            index += 1
            continue
        suffix_start = label_end + 1
        suffix_end: int | None = None
        if source[suffix_start : suffix_start + 1] == "(":
            suffix_end = closing_markdown_delimiter(source, suffix_start, "(", ")")
        elif source[suffix_start : suffix_start + 1] == "[":
            suffix_end = closing_markdown_delimiter(source, suffix_start, "[", "]")
        if suffix_end is None:
            visible.append(source[index : label_end + 1])
            index = label_end + 1
            continue
        visible.append(source[label_start + 1 : label_end])
        index = suffix_end + 1
    return "".join(visible)


def markdown_code_span(source: str, index: int) -> tuple[int, int] | None:
    """Return the content bounds for a code span beginning at index."""
    run_end = index
    while run_end < len(source) and source[run_end] == "`":
        run_end += 1
    marker = source[index:run_end]
    close = source.find(marker, run_end)
    while close >= 0 and (
        source[close - 1 : close] == "`"
        or source[close + len(marker) : close + len(marker) + 1] == "`"
    ):
        close = source.find(marker, close + len(marker))
    if close < 0:
        return None
    return run_end, close


def markdown_without_inline_code(source: str) -> tuple[str, bool]:
    """Remove code-span syntax and report whether a span renders code content."""
    visible: list[str] = []
    index = 0
    while index < len(source):
        if source[index] != "`":
            visible.append(source[index])
            index += 1
            continue
        bounds = markdown_code_span(source, index)
        if bounds is None:
            run_end = index
            while run_end < len(source) and source[run_end] == "`":
                run_end += 1
            marker = source[index:run_end]
            visible.append(marker)
            index = run_end
            continue
        run_end, close = bounds
        marker = source[index:run_end]
        code = source[run_end:close].replace("\n", " ")
        if code.startswith(" ") and code.endswith(" ") and code.strip():
            code = code[1:-1]
        if contains_meaningful_rendered_character(code):
            return "", True
        index = close + len(marker)
    return "".join(visible), False


def escape_html_syntax(source: str) -> str:
    """Escape code text only for the internal HTML visibility pass."""
    return source.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def protect_inline_code_from_html(source: str) -> str:
    """Keep HTML-like code-span content visible to the Markdown pass."""
    protected: list[str] = []
    index = 0
    while index < len(source):
        if source[index] != "`":
            protected.append(source[index])
            index += 1
            continue
        bounds = markdown_code_span(source, index)
        if bounds is None:
            run_end = index
            while run_end < len(source) and source[run_end] == "`":
                run_end += 1
            marker = source[index:run_end]
            protected.append(marker)
            index = run_end
            continue
        run_end, close = bounds
        marker = source[index:run_end]
        protected.append(marker)
        protected.append(escape_html_syntax(source[run_end:close]))
        protected.append(marker)
        index = close + len(marker)
    return "".join(protected)


def protect_markdown_code_from_html(section: str) -> str:
    """Protect fenced and inline code before parsing surrounding raw HTML."""
    protected: list[str] = []
    prose: list[str] = []
    fence: tuple[str, int] | None = None

    def flush_prose() -> None:
        if prose:
            protected.append(protect_inline_code_from_html("".join(prose)))
            prose.clear()

    for line in section.splitlines(keepends=True):
        without_newline = line.rstrip("\r\n")
        if fence is not None:
            marker, minimum_length = fence
            close = re.fullmatch(
                rf" {{0,3}}({re.escape(marker)}+)[ \t]*", without_newline
            )
            if close is not None and len(close.group(1)) >= minimum_length:
                protected.append(line)
                fence = None
            else:
                protected.append(escape_html_syntax(line))
            continue
        opening = re.match(r"^ {0,3}(`{3,}|~{3,})", without_newline)
        if opening is not None:
            marker = opening.group(1)
            info = without_newline[opening.end() :]
            if marker[0] == "~" or "`" not in info:
                flush_prose()
                protected.append(line)
                fence = (marker[0], len(marker))
                continue
        prose.append(line)
    flush_prose()
    return "".join(protected)


def markdown_prose_and_code(section: str) -> tuple[str, bool]:
    """Remove non-rendered block syntax and report visible fenced code."""
    prose: list[str] = []
    fence: tuple[str, int] | None = None
    reference_continuation: str | None = None
    for line in section.splitlines(keepends=True):
        without_newline = line.rstrip("\r\n")
        if fence is not None:
            marker, minimum_length = fence
            close = re.fullmatch(
                rf" {{0,3}}({re.escape(marker)}+)[ \t]*", without_newline
            )
            if close is not None and len(close.group(1)) >= minimum_length:
                fence = None
            elif contains_meaningful_rendered_character(without_newline):
                return "", True
            continue
        if reference_continuation is not None:
            if not without_newline.strip():
                reference_continuation = None
                continue
            if reference_continuation == "destination":
                reference_continuation = "title"
                continue
            if COMMONMARK_REFERENCE_TITLE_RE.fullmatch(without_newline) is not None:
                reference_continuation = None
                continue
            reference_continuation = None
        opening = re.match(r"^ {0,3}(`{3,}|~{3,})", without_newline)
        if opening is not None:
            marker = opening.group(1)
            info = without_newline[opening.end() :]
            if marker[0] == "~" or "`" not in info:
                fence = (marker[0], len(marker))
                continue
        reference = COMMONMARK_REFERENCE_DEFINITION_RE.match(without_newline)
        if reference is not None:
            reference_continuation = (
                "destination"
                if not without_newline[reference.end() :].strip()
                else "title"
            )
            continue
        prose.append(line)
    return "".join(prose), False


def markdown_without_block_markers(source: str) -> str:
    """Remove list and quote markers that do not represent section text."""
    visible: list[str] = []
    for line in source.splitlines(keepends=True):
        while True:
            quote = re.match(r"^ {0,3}>[ \t]?", line)
            if quote is None:
                break
            line = line[quote.end() :]
        while True:
            marker = re.match(r"^ {0,3}(?:[-+*]|[0-9]{1,9}[.)])(?:[ \t]+|$)", line)
            if marker is None:
                break
            line = line[marker.end() :]
        task = re.match(r"^\[[ xX]\](?:[ \t]+|$)", line)
        if task is not None:
            line = line[task.end() :]
        visible.append(line)
    return "".join(visible)


def contains_meaningful_rendered_character(source: str) -> bool:
    """Accept alphanumeric text and pictographs, but not Markdown syntax alone."""
    return any(
        character.isalnum() or unicodedata.category(character) == "So"
        for character in source
    )


def has_meaningful_visible_text(section: str) -> bool:
    """Return whether one bounded section renders meaningful HTML or Markdown."""
    section = protect_markdown_escapes(section)
    section = protect_markdown_code_from_html(section)
    section = COMMONMARK_AUTOLINK_RE.sub(lambda match: match.group("label"), section)
    parser = MarkdownOutsideRawHtml()
    parser.feed(section)
    parser.close()
    prose, fenced_code = markdown_prose_and_code("".join(parser.text))
    if fenced_code:
        return True
    prose, inline_code = markdown_without_inline_code(prose)
    if inline_code:
        return True
    prose = markdown_without_link_destinations(prose)
    prose = markdown_without_block_markers(prose)
    return contains_meaningful_rendered_character(prose)


def render_release_notes(changelog: str, tag: str) -> str:
    """Validate and return one reviewed changelog section body."""
    if not RELEASE_TAG_RE.fullmatch(tag):
        raise ValueError(
            f"{tag!r} is not a release tag, expected vX.Y.Z, rpptx-vX.Y.Z, "
            "py-rdocx-vX.Y.Z, or py-rpptx-vX.Y.Z"
        )

    lines = changelog.splitlines(keepends=True)
    release_matches = [
        index
        for index, heading in markdown_heading_lines(lines, "## ")
        if heading == tag
    ]
    if len(release_matches) != 1:
        raise ValueError(
            f"CHANGELOG.md must contain exactly one `## {tag}` heading, "
            f"found {len(release_matches)}"
        )

    start = release_matches[0] + 1
    later_sections = [
        index
        for index, _ in markdown_heading_lines(lines[start:], "## ")
    ]
    end = start + later_sections[0] if later_sections else len(lines)
    body_lines = lines[start:end]
    while body_lines and not body_lines[0].strip():
        body_lines.pop(0)
    while body_lines and not body_lines[-1].strip():
        body_lines.pop()
    body = "".join(body_lines)
    if not body.strip():
        raise ValueError(f"`## {tag}` has no release-note body")
    if RELEASE_NOTE_PLACEHOLDER_RE.search(body):
        raise ValueError(f"`## {tag}` contains a placeholder token")

    headings = markdown_heading_lines(body_lines, "### ")
    found = [heading for _, heading in headings]
    for required in RELEASE_NOTE_HEADINGS:
        count = found.count(required)
        if count != 1:
            raise ValueError(
                f"`## {tag}` must contain exactly one `### {required}` heading, "
                f"found {count}"
            )
    required_positions = [found.index(required) for required in RELEASE_NOTE_HEADINGS]
    if required_positions != sorted(required_positions):
        raise ValueError(f"`## {tag}` required headings are out of order")

    for heading_index, (line_index, heading) in enumerate(headings):
        if heading not in RELEASE_NOTE_HEADINGS:
            continue
        next_line = (
            headings[heading_index + 1][0]
            if heading_index + 1 < len(headings)
            else len(body_lines)
        )
        section_text = "".join(body_lines[line_index + 1 : next_line])
        if not has_meaningful_visible_text(section_text):
            raise ValueError(f"`## {tag}` section `### {heading}` is empty")

    return body.rstrip() + "\n"


def cmd_release_notes(args: argparse.Namespace) -> int:
    try:
        notes = render_release_notes(CHANGELOG.read_text(encoding="utf-8"), args.tag)
    except (OSError, ValueError) as error:
        die(f"release notes: {error}")
    if args.render:
        sys.stdout.write(notes)
    else:
        print(f"release-notes {args.tag}: ok")
    return 0


def validate_python_project_metadata(
    metadata: object, distribution: str, artifact_name: str
) -> None:
    """Validate the user-facing metadata embedded in one Python artifact."""
    if not isinstance(metadata, Message):
        raise TypeError("metadata must be an email message")
    contract = PYTHON_RELEASE_METADATA[distribution]
    if metadata.get("Summary") != contract["summary"]:
        raise ValueError(f"{artifact_name} has incomplete project summary")
    description_content_type = metadata.get(
        "Description-Content-Type", ""
    ).partition(";")[0].strip().lower()
    if description_content_type != "text/markdown":
        raise ValueError(f"{artifact_name} lacks a Markdown long description")
    if metadata.get("Author") != "Atul Sharma":
        raise ValueError(f"{artifact_name} lacks the reviewed project author")
    keywords = tuple(
        keyword.strip()
        for keyword in metadata.get("Keywords", "").split(",")
        if keyword.strip()
    )
    if keywords != contract["keywords"]:
        raise ValueError(f"{artifact_name} has incomplete project keywords")
    if set(metadata.get_all("Classifier", [])) != PYTHON_RELEASE_CLASSIFIERS:
        raise ValueError(f"{artifact_name} has incomplete project classifiers")
    if set(metadata.get_all("Project-URL", [])) != PYTHON_RELEASE_PROJECT_URLS:
        raise ValueError(f"{artifact_name} has incomplete project URLs")
    description = metadata.get_payload()
    for required in (
        f"# {distribution}-py",
        "## Installation",
        f"python -m pip install {distribution}",
        "## Quick start",
        "## Type checking",
        "## Project links",
    ):
        if required not in description:
            raise ValueError(
                f"{artifact_name} long description lacks {required!r}"
            )
    reviewed_description = (REPO / contract["readme"]).read_text(encoding="utf-8")
    normalized_description = description.replace("\r\n", "\n")
    normalized_reviewed = reviewed_description.replace("\r\n", "\n")
    if normalized_description.rstrip("\n") != normalized_reviewed.rstrip("\n"):
        raise ValueError(
            f"{artifact_name} long description differs from the reviewed "
            f"{contract['readme']}"
        )


RELEASE_CRATES = {
    "rdocx": (
        "rdocx-opc", "rdocx-oxml", "rdocx-layout", "rdocx-html",
        "rdocx-pdf", "rdocx", "rdocx-cli",
    ),
    "rpptx": (
        "oxml-core", "oxml-opc", "oxml-media", "oxml-layout",
        "oxml-drawing", "oxml-pdf", "oxml-sml", "oxml-cli-support",
        "oxml-chart", "rpptx-oxml", "rpptx-chart", "rpptx-layout",
        "rpptx-render", "rpptx", "rpptx-cli",
    ),
}



# Strict SemVer policy is owned here, with RELEASE_CRATES as the sole membership owner.
# Pinned cargo-semver-checks 0.51.0 policy, literal insertion into existing workflow.py.
# Catalog SHA256: 4f1f5ae93d0c2e325eea77d11aa7db842adf6b506b7c3a17b444e95198243ce4
# Tuple fields: required_update, default_level. Built-in Warn/Allow stay intact.
SEMVER_QUERY_POLICY = {
    'attribute_proc_macro_missing': ('major', 'deny'),
    'auto_trait_impl_added': ('minor', 'allow'),
    'auto_trait_impl_removed': ('major', 'deny'),
    'constructible_struct_adds_field': ('major', 'deny'),
    'constructible_struct_adds_private_field': ('major', 'deny'),
    'constructible_struct_changed_type': ('major', 'deny'),
    'copy_impl_added': ('major', 'deny'),
    'declarative_macro_missing': ('major', 'deny'),
    'derive_helper_attr_removed': ('major', 'deny'),
    'derive_proc_macro_missing': ('major', 'deny'),
    'derive_trait_impl_removed': ('major', 'deny'),
    'enum_changed_kind': ('major', 'deny'),
    'enum_discriminants_undefined_non_exhaustive_variant': ('major', 'deny'),
    'enum_discriminants_undefined_non_unit_variant': ('major', 'deny'),
    'enum_marked_non_exhaustive': ('major', 'deny'),
    'enum_missing': ('major', 'deny'),
    'enum_must_use_added': ('minor', 'deny'),
    'enum_must_use_removed': ('minor', 'allow'),
    'enum_no_longer_non_exhaustive': ('minor', 'allow'),
    'enum_no_repr_variant_discriminant_changed': ('major', 'deny'),
    'enum_non_exhaustive_struct_variant_field_added': ('minor', 'allow'),
    'enum_non_exhaustive_tuple_variant_changed_kind': ('major', 'deny'),
    'enum_non_exhaustive_tuple_variant_field_added': ('minor', 'allow'),
    'enum_now_doc_hidden': ('major', 'deny'),
    'enum_repr_int_added': ('minor', 'allow'),
    'enum_repr_int_changed': ('major', 'deny'),
    'enum_repr_int_removed': ('major', 'deny'),
    'enum_repr_transparent_removed': ('major', 'deny'),
    'enum_repr_variant_discriminant_changed': ('major', 'deny'),
    'enum_struct_variant_changed_kind': ('major', 'deny'),
    'enum_struct_variant_field_added': ('major', 'deny'),
    'enum_struct_variant_field_marked_deprecated': ('minor', 'deny'),
    'enum_struct_variant_field_missing': ('major', 'deny'),
    'enum_struct_variant_field_now_doc_hidden': ('major', 'deny'),
    'enum_tuple_variant_changed_kind': ('major', 'deny'),
    'enum_tuple_variant_field_added': ('major', 'deny'),
    'enum_tuple_variant_field_marked_deprecated': ('minor', 'deny'),
    'enum_tuple_variant_field_missing': ('major', 'deny'),
    'enum_tuple_variant_field_now_doc_hidden': ('major', 'deny'),
    'enum_unit_variant_changed_kind': ('major', 'deny'),
    'enum_variant_added': ('major', 'deny'),
    'enum_variant_marked_deprecated': ('minor', 'deny'),
    'enum_variant_marked_non_exhaustive': ('major', 'deny'),
    'enum_variant_missing': ('major', 'deny'),
    'enum_variant_no_longer_non_exhaustive': ('minor', 'allow'),
    'exhaustive_enum_added': ('minor', 'allow'),
    'exhaustive_struct_added': ('minor', 'allow'),
    'exhaustive_struct_with_doc_hidden_fields_added': ('minor', 'allow'),
    'exhaustive_struct_with_private_fields_added': ('minor', 'allow'),
    'exported_function_abi_no_longer_unwind': ('major', 'deny'),
    'exported_function_abi_now_unwind': ('major', 'deny'),
    'exported_function_changed_abi': ('major', 'deny'),
    'exported_function_now_returns_unit': ('major', 'deny'),
    'exported_function_parameter_count_changed': ('major', 'deny'),
    'exported_function_requires_more_target_features': ('major', 'deny'),
    'exported_function_return_value_added': ('major', 'deny'),
    'exported_function_target_feature_added': ('major', 'deny'),
    'feature_missing': ('major', 'deny'),
    'feature_newly_enables_feature': ('minor', 'allow'),
    'feature_no_longer_enables_feature': ('major', 'deny'),
    'feature_not_enabled_by_default': ('major', 'deny'),
    'function_abi_no_longer_unwind': ('major', 'deny'),
    'function_abi_now_unwind': ('major', 'deny'),
    'function_changed_abi': ('major', 'deny'),
    'function_const_generic_reordered': ('major', 'deny'),
    'function_const_removed': ('major', 'deny'),
    'function_export_name_changed': ('major', 'deny'),
    'function_generic_type_reordered': ('major', 'deny'),
    'function_like_proc_macro_missing': ('major', 'deny'),
    'function_marked_deprecated': ('minor', 'deny'),
    'function_missing': ('major', 'deny'),
    'function_must_use_added': ('minor', 'deny'),
    'function_must_use_removed': ('minor', 'allow'),
    'function_no_longer_unsafe': ('minor', 'allow'),
    'function_no_longer_variadic': ('major', 'deny'),
    'function_now_const': ('minor', 'allow'),
    'function_now_doc_hidden': ('major', 'deny'),
    'function_now_returns_unit': ('major', 'deny'),
    'function_now_variadic': ('major', 'deny'),
    'function_parameter_count_changed': ('major', 'deny'),
    'function_requires_different_const_generic_params': ('major', 'deny'),
    'function_requires_different_generic_type_params': ('major', 'deny'),
    'function_unsafe_added': ('major', 'deny'),
    'global_value_marked_deprecated': ('minor', 'deny'),
    'inherent_associated_const_now_doc_hidden': ('major', 'deny'),
    'inherent_associated_pub_const_added': ('minor', 'allow'),
    'inherent_associated_pub_const_missing': ('major', 'deny'),
    'inherent_method_added': ('minor', 'allow'),
    'inherent_method_changed_abi': ('major', 'deny'),
    'inherent_method_const_generic_reordered': ('major', 'deny'),
    'inherent_method_const_removed': ('major', 'deny'),
    'inherent_method_generic_type_reordered': ('major', 'deny'),
    'inherent_method_missing': ('major', 'deny'),
    'inherent_method_must_use_added': ('minor', 'deny'),
    'inherent_method_must_use_removed': ('minor', 'allow'),
    'inherent_method_no_longer_unsafe': ('minor', 'allow'),
    'inherent_method_no_longer_unwind': ('major', 'deny'),
    'inherent_method_no_longer_variadic': ('major', 'deny'),
    'inherent_method_now_const': ('minor', 'allow'),
    'inherent_method_now_doc_hidden': ('major', 'deny'),
    'inherent_method_now_returns_unit': ('major', 'deny'),
    'inherent_method_now_unwind': ('major', 'deny'),
    'inherent_method_now_variadic': ('major', 'deny'),
    'inherent_method_unsafe_added': ('major', 'deny'),
    'macro_marked_deprecated': ('minor', 'deny'),
    'macro_no_longer_exported': ('major', 'deny'),
    'macro_now_doc_hidden': ('major', 'deny'),
    'method_export_name_changed': ('major', 'deny'),
    'method_no_longer_has_receiver': ('major', 'deny'),
    'method_parameter_count_changed': ('major', 'deny'),
    'method_receiver_mut_ref_became_owned': ('major', 'deny'),
    'method_receiver_ref_became_mut': ('major', 'deny'),
    'method_receiver_ref_became_owned': ('major', 'deny'),
    'method_receiver_type_changed': ('major', 'deny'),
    'method_requires_different_const_generic_params': ('major', 'deny'),
    'method_requires_different_generic_type_params': ('major', 'deny'),
    'module_missing': ('major', 'deny'),
    'non_exhaustive_enum_added': ('minor', 'allow'),
    'non_exhaustive_struct_added': ('minor', 'allow'),
    'non_exhaustive_struct_changed_type': ('major', 'deny'),
    'partial_ord_enum_struct_variant_fields_reordered': ('major', 'warn'),
    'partial_ord_enum_variants_reordered': ('major', 'warn'),
    'partial_ord_struct_fields_reordered': ('major', 'warn'),
    'proc_macro_marked_deprecated': ('minor', 'deny'),
    'proc_macro_now_doc_hidden': ('major', 'deny'),
    'pub_api_sealed_trait_became_unconditionally_sealed': ('minor', 'warn'),
    'pub_api_sealed_trait_became_unsealed': ('minor', 'warn'),
    'pub_api_sealed_trait_method_receiver_added': ('major', 'warn'),
    'pub_api_sealed_trait_method_receiver_mut_ref_became_ref': ('major', 'warn'),
    'pub_api_sealed_trait_method_return_value_added': ('major', 'warn'),
    'pub_api_sealed_trait_method_target_feature_removed': ('major', 'warn'),
    'pub_const_added': ('minor', 'allow'),
    'pub_module_level_const_missing': ('major', 'deny'),
    'pub_module_level_const_now_doc_hidden': ('major', 'deny'),
    'pub_static_added': ('minor', 'allow'),
    'pub_static_missing': ('major', 'deny'),
    'pub_static_mut_now_immutable': ('major', 'deny'),
    'pub_static_now_doc_hidden': ('major', 'deny'),
    'pub_static_now_mutable': ('major', 'deny'),
    'repr_align_added': ('major', 'deny'),
    'repr_align_changed': ('major', 'deny'),
    'repr_align_removed': ('major', 'deny'),
    'repr_c_added': ('minor', 'allow'),
    'repr_c_enum_struct_variant_fields_reordered': ('major', 'warn'),
    'repr_c_plain_struct_fields_reordered': ('major', 'warn'),
    'repr_c_removed': ('major', 'deny'),
    'repr_packed_added': ('major', 'deny'),
    'repr_packed_changed': ('major', 'deny'),
    'repr_packed_removed': ('major', 'deny'),
    'repr_transparent_added': ('minor', 'deny'),
    'safe_function_requires_more_target_features': ('major', 'deny'),
    'safe_function_target_feature_added': ('major', 'deny'),
    'safe_inherent_method_requires_more_target_features': ('major', 'deny'),
    'safe_inherent_method_target_feature_added': ('major', 'deny'),
    'sized_impl_removed': ('major', 'deny'),
    'static_became_unsafe': ('major', 'deny'),
    'struct_field_marked_deprecated': ('minor', 'deny'),
    'struct_marked_non_exhaustive': ('major', 'deny'),
    'struct_missing': ('major', 'deny'),
    'struct_must_use_added': ('minor', 'deny'),
    'struct_must_use_removed': ('minor', 'allow'),
    'struct_no_longer_has_non_pub_fields': ('minor', 'allow'),
    'struct_no_longer_non_exhaustive': ('minor', 'allow'),
    'struct_now_doc_hidden': ('major', 'deny'),
    'struct_pub_field_missing': ('major', 'deny'),
    'struct_pub_field_now_doc_hidden': ('major', 'deny'),
    'struct_repr_transparent_removed': ('major', 'deny'),
    'struct_with_no_pub_fields_changed_type': ('major', 'deny'),
    'struct_with_pub_fields_changed_type': ('major', 'deny'),
    'trait_added_supertrait': ('major', 'deny'),
    'trait_allows_fewer_const_generic_params': ('major', 'deny'),
    'trait_allows_fewer_generic_type_params': ('major', 'deny'),
    'trait_associated_const_added': ('major', 'deny'),
    'trait_associated_const_default_removed': ('major', 'deny'),
    'trait_associated_const_marked_deprecated': ('minor', 'deny'),
    'trait_associated_const_now_doc_hidden': ('major', 'deny'),
    'trait_associated_type_added': ('major', 'deny'),
    'trait_associated_type_default_removed': ('major', 'deny'),
    'trait_associated_type_marked_deprecated': ('minor', 'deny'),
    'trait_associated_type_now_doc_hidden': ('major', 'deny'),
    'trait_changed_kind': ('major', 'deny'),
    'trait_const_generic_reordered': ('major', 'deny'),
    'trait_generic_type_reordered': ('major', 'deny'),
    'trait_marked_deprecated': ('minor', 'deny'),
    'trait_method_added': ('major', 'deny'),
    'trait_method_changed_abi': ('major', 'deny'),
    'trait_method_const_generic_reordered': ('major', 'deny'),
    'trait_method_default_impl_removed': ('major', 'deny'),
    'trait_method_generic_type_reordered': ('major', 'deny'),
    'trait_method_marked_deprecated': ('minor', 'deny'),
    'trait_method_missing': ('major', 'deny'),
    'trait_method_no_longer_has_receiver': ('major', 'deny'),
    'trait_method_no_longer_unwind': ('major', 'deny'),
    'trait_method_no_longer_variadic': ('major', 'deny'),
    'trait_method_now_doc_hidden': ('major', 'deny'),
    'trait_method_now_returns_unit': ('major', 'deny'),
    'trait_method_now_unwind': ('major', 'deny'),
    'trait_method_now_variadic': ('major', 'deny'),
    'trait_method_parameter_count_changed': ('major', 'deny'),
    'trait_method_receiver_added': ('major', 'deny'),
    'trait_method_receiver_mut_ref_became_owned': ('major', 'deny'),
    'trait_method_receiver_mut_ref_became_ref': ('major', 'deny'),
    'trait_method_receiver_owned_became_mut_ref': ('major', 'deny'),
    'trait_method_receiver_owned_became_ref': ('major', 'deny'),
    'trait_method_receiver_ref_became_mut': ('major', 'deny'),
    'trait_method_receiver_ref_became_owned': ('major', 'deny'),
    'trait_method_receiver_type_changed': ('major', 'deny'),
    'trait_method_requires_different_const_generic_params': ('major', 'deny'),
    'trait_method_requires_different_generic_type_params': ('major', 'deny'),
    'trait_method_return_value_added': ('major', 'deny'),
    'trait_method_target_feature_removed': ('major', 'deny'),
    'trait_method_unsafe_added': ('major', 'deny'),
    'trait_method_unsafe_removed': ('major', 'deny'),
    'trait_mismatched_generic_lifetimes': ('major', 'deny'),
    'trait_missing': ('major', 'deny'),
    'trait_must_use_added': ('minor', 'deny'),
    'trait_must_use_removed': ('minor', 'allow'),
    'trait_newly_sealed': ('major', 'deny'),
    'trait_no_longer_dyn_compatible': ('major', 'deny'),
    'trait_now_doc_hidden': ('major', 'deny'),
    'trait_removed_associated_constant': ('major', 'deny'),
    'trait_removed_associated_type': ('major', 'deny'),
    'trait_removed_supertrait': ('major', 'deny'),
    'trait_requires_more_const_generic_params': ('major', 'deny'),
    'trait_requires_more_generic_type_params': ('major', 'deny'),
    'trait_unsafe_added': ('major', 'deny'),
    'trait_unsafe_removed': ('major', 'deny'),
    'tuple_struct_to_plain_struct': ('major', 'deny'),
    'type_allows_fewer_const_generic_params': ('major', 'deny'),
    'type_allows_fewer_generic_type_params': ('major', 'deny'),
    'type_associated_const_marked_deprecated': ('minor', 'deny'),
    'type_const_generic_reordered': ('major', 'deny'),
    'type_generic_type_reordered': ('major', 'deny'),
    'type_marked_deprecated': ('minor', 'deny'),
    'type_method_marked_deprecated': ('minor', 'deny'),
    'type_mismatched_generic_lifetimes': ('major', 'deny'),
    'type_requires_more_const_generic_params': ('major', 'deny'),
    'type_requires_more_generic_type_params': ('major', 'deny'),
    'unconditionally_sealed_trait_became_pub_api_sealed': ('minor', 'warn'),
    'unconditionally_sealed_trait_became_unsealed': ('minor', 'warn'),
    'union_added': ('minor', 'allow'),
    'union_changed_kind': ('major', 'deny'),
    'union_changed_to_incompatible_struct': ('major', 'deny'),
    'union_field_added_with_all_pub_fields': ('major', 'deny'),
    'union_field_added_with_non_pub_fields': ('major', 'warn'),
    'union_field_marked_deprecated': ('minor', 'deny'),
    'union_field_missing': ('major', 'deny'),
    'union_missing': ('major', 'deny'),
    'union_must_use_added': ('minor', 'deny'),
    'union_must_use_removed': ('minor', 'allow'),
    'union_now_doc_hidden': ('major', 'deny'),
    'union_pub_field_now_doc_hidden': ('major', 'deny'),
    'union_with_multiple_pub_fields_changed_to_struct': ('major', 'deny'),
    'unit_struct_changed_kind': ('major', 'deny'),
    'unsafe_function_requires_more_target_features': ('major', 'deny'),
    'unsafe_function_target_feature_added': ('major', 'deny'),
    'unsafe_inherent_method_requires_more_target_features': ('major', 'deny'),
    'unsafe_inherent_method_target_feature_added': ('major', 'deny'),
    'unsafe_trait_method_requires_more_target_features': ('major', 'deny'),
    'unsafe_trait_method_target_feature_added': ('major', 'deny'),
}

# Official native archive and executable pins, not native execution proof.
SEMVER_NATIVE_TOOLS = {
    'x86_64-unknown-linux-gnu': {
        'archive_url': 'https://github.com/obi1kenobi/cargo-semver-checks/releases/download/v0.51.0/cargo-semver-checks-x86_64-unknown-linux-gnu.tar.gz',
        'archive_member': 'cargo-semver-checks',
        'archive_sha256': 'cacadafaa5cd27ae6d05a74a103266917ebab4c5e43c898360e177d4fbf06a65',
        'executable_sha256': 'bc9eb06ce04e97d6160c875ffdee1f6e56793e423dea025f344826be03e66829',
    },
    'aarch64-apple-darwin': {
        'archive_url': 'https://github.com/obi1kenobi/cargo-semver-checks/releases/download/v0.51.0/cargo-semver-checks-aarch64-apple-darwin.tar.gz',
        'archive_member': 'cargo-semver-checks',
        'archive_sha256': '3cec5da41eee9a3336d75ad43d1f30a5c9680818f4ba209024786f452cabab5b',
        'executable_sha256': '30493e32d77a4a9f90ea6695cf67765a4c9106c46a7db4603f2026043d469bb8',
    },
    'x86_64-pc-windows-msvc': {
        'archive_url': 'https://github.com/obi1kenobi/cargo-semver-checks/releases/download/v0.51.0/cargo-semver-checks-x86_64-pc-windows-msvc.zip',
        'archive_member': 'cargo-semver-checks.exe',
        'archive_sha256': '7663ba2c938b59f0476d375ee1b5d8bece36a81819612d1541bc6ccf6316c4b7',
        'executable_sha256': '11a5efa48f9fce6bc4012c56576a6e862a287390afcf396e9faeca4e77f5aab2',
    },
}


# TEMP proposed constant for existing sprint_workflow.py. Not adopted or executed.
SEMVER_GH_VERIFY_TOOLS = {'x86_64-unknown-linux-gnu': {'version': '2.97.0', 'archive_url': 'https://github.com/cli/cli/releases/download/v2.97.0/gh_2.97.0_linux_amd64.tar.gz', 'archive_sha256': 'a2c9b8497e1f85b1ad0dfcb78b5a622e098801b8e461e459e88e1ee12f018112', 'archive_bytes': 14770812, 'executable_member': 'gh_2.97.0_linux_amd64/bin/gh', 'executable_sha256': '141507c337e8b202ad398550c3b73d72f5af92e86f71665214538a81efd4c409', 'executable_bytes': 40992930}, 'aarch64-apple-darwin': {'version': '2.97.0', 'archive_url': 'https://github.com/cli/cli/releases/download/v2.97.0/gh_2.97.0_macOS_arm64.zip', 'archive_sha256': 'a58b8fd77b417a38f47a0b54d1370c59b0fcdb324ccc9ca002b0998f7c4c999e', 'archive_bytes': 13845290, 'executable_member': 'gh_2.97.0_macOS_arm64/bin/gh', 'executable_sha256': '0d17dddf96bcc1dc50f3420a064d593d64016b0be16286a6c26121f2a5cb8316', 'executable_bytes': 38857376}, 'x86_64-pc-windows-msvc': {'version': '2.97.0', 'archive_url': 'https://github.com/cli/cli/releases/download/v2.97.0/gh_2.97.0_windows_amd64.zip', 'archive_sha256': '35d7fe05c4dd1411ffda1e73dfc7c6f44b75c936ca51fa6595c657fdc0350cec', 'archive_bytes': 14938517, 'executable_member': 'bin/gh.exe', 'executable_sha256': 'e2efa10a5d2ce93cac9bc4b676932b62947c0967c01c8f2c3a9cb4437ad358d3', 'executable_bytes': 41775416}}


SEMVER_TRUSTED_ROOT_JSONL = b'{"mediaType":"application/vnd.dev.sigstore.trustedroot+json;version=0.1","tlogs":[{"baseUrl":"https://rekor.sigstore.dev","hashAlgorithm":"SHA2_256","publicKey":{"rawBytes":"MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE2G2Y+2tabdTV5BcGiBIx0a9fAFwrkBbmLSGtks4L3qX6yYY0zufBnhC8Ur/iy55GhWP/9A/bY2LhC30M9+RYtw==","keyDetails":"PKIX_ECDSA_P256_SHA_256","validFor":{"start":"2021-01-12T11:53:27Z"}},"logId":{"keyId":"wNI9atQGlz+VWfO6LRygH4QUfY/8W4RFwiT5i5WRgB0="}},{"baseUrl":"https://log2025-1.rekor.sigstore.dev","hashAlgorithm":"SHA2_256","publicKey":{"rawBytes":"MCowBQYDK2VwAyEAt8rlp1knGwjfbcXAYPYAkn0XiLz1x8O4t0YkEhie244=","keyDetails":"PKIX_ED25519","validFor":{"start":"2025-09-23T00:00:00Z"}},"logId":{"keyId":"zxGZFVvd0FEmjR8WrFwMdcAJ9vtaY/QXf44Y1wUeP6A="}}],"certificateAuthorities":[{"subject":{"organization":"sigstore.dev","commonName":"sigstore"},"uri":"https://fulcio.sigstore.dev","certChain":{"certificates":[{"rawBytes":"MIIB+DCCAX6gAwIBAgITNVkDZoCiofPDsy7dfm6geLbuhzAKBggqhkjOPQQDAzAqMRUwEwYDVQQKEwxzaWdzdG9yZS5kZXYxETAPBgNVBAMTCHNpZ3N0b3JlMB4XDTIxMDMwNzAzMjAyOVoXDTMxMDIyMzAzMjAyOVowKjEVMBMGA1UEChMMc2lnc3RvcmUuZGV2MREwDwYDVQQDEwhzaWdzdG9yZTB2MBAGByqGSM49AgEGBSuBBAAiA2IABLSyA7Ii5k+pNO8ZEWY0ylemWDowOkNa3kL+GZE5Z5GWehL9/A9bRNA3RbrsZ5i0JcastaRL7Sp5fp/jD5dxqc/UdTVnlvS16an+2Yfswe/QuLolRUCrcOE2+2iA5+tzd6NmMGQwDgYDVR0PAQH/BAQDAgEGMBIGA1UdEwEB/wQIMAYBAf8CAQEwHQYDVR0OBBYEFMjFHQBBmiQpMlEk6w2uSu1KBtPsMB8GA1UdIwQYMBaAFMjFHQBBmiQpMlEk6w2uSu1KBtPsMAoGCCqGSM49BAMDA2gAMGUCMH8liWJfMui6vXXBhjDgY4MwslmN/TJxVe/83WrFomwmNf056y1X48F9c4m3a3ozXAIxAKjRay5/aj/jsKKGIkmQatjI8uupHr/+CxFvaJWmpYqNkLDGRU+9orzh5hI2RrcuaQ=="}]},"validFor":{"start":"2021-03-07T03:20:29Z","end":"2022-12-31T23:59:59.999Z"}},{"subject":{"organization":"sigstore.dev","commonName":"sigstore"},"uri":"https://fulcio.sigstore.dev","certChain":{"certificates":[{"rawBytes":"MIICGjCCAaGgAwIBAgIUALnViVfnU0brJasmRkHrn/UnfaQwCgYIKoZIzj0EAwMwKjEVMBMGA1UEChMMc2lnc3RvcmUuZGV2MREwDwYDVQQDEwhzaWdzdG9yZTAeFw0yMjA0MTMyMDA2MTVaFw0zMTEwMDUxMzU2NThaMDcxFTATBgNVBAoTDHNpZ3N0b3JlLmRldjEeMBwGA1UEAxMVc2lnc3RvcmUtaW50ZXJtZWRpYXRlMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE8RVS/ysH+NOvuDZyPIZtilgUF9NlarYpAd9HP1vBBH1U5CV77LSS7s0ZiH4nE7Hv7ptS6LvvR/STk798LVgMzLlJ4HeIfF3tHSaexLcYpSASr1kS0N/RgBJz/9jWCiXno3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU39Ppz1YkEZb5qNjpKFWixi4YZD8wHwYDVR0jBBgwFoAUWMAeX5FFpWapesyQoZMi0CrFxfowCgYIKoZIzj0EAwMDZwAwZAIwPCsQK4DYiZYDPIaDi5HFKnfxXx6ASSVmERfsynYBiX2X6SJRnZU84/9DZdnFvvxmAjBOt6QpBlc4J/0DxvkTCqpclvziL6BCCPnjdlIB3Pu3BxsPmygUY7Ii2zbdCdliiow="},{"rawBytes":"MIIB9zCCAXygAwIBAgIUALZNAPFdxHPwjeDloDwyYChAO/4wCgYIKoZIzj0EAwMwKjEVMBMGA1UEChMMc2lnc3RvcmUuZGV2MREwDwYDVQQDEwhzaWdzdG9yZTAeFw0yMTEwMDcxMzU2NTlaFw0zMTEwMDUxMzU2NThaMCoxFTATBgNVBAoTDHNpZ3N0b3JlLmRldjERMA8GA1UEAxMIc2lnc3RvcmUwdjAQBgcqhkjOPQIBBgUrgQQAIgNiAAT7XeFT4rb3PQGwS4IajtLk3/OlnpgangaBclYpsYBr5i+4ynB07ceb3LP0OIOZdxexX69c5iVuyJRQ+Hz05yi+UF3uBWAlHpiS5sh0+H2GHE7SXrk1EC5m1Tr19L9gg92jYzBhMA4GA1UdDwEB/wQEAwIBBjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBRYwB5fkUWlZql6zJChkyLQKsXF+jAfBgNVHSMEGDAWgBRYwB5fkUWlZql6zJChkyLQKsXF+jAKBggqhkjOPQQDAwNpADBmAjEAj1nHeXZp+13NWBNa+EDsDP8G1WWg1tCMWP/WHPqpaVo0jhsweNFZgSs0eE7wYI4qAjEA2WB9ot98sIkoF3vZYdd3/VtWB5b9TNMea7Ix/stJ5TfcLLeABLE4BNJOsQ4vnBHJ"}]},"validFor":{"start":"2022-04-13T20:06:15Z"}}],"ctlogs":[{"baseUrl":"https://ctfe.sigstore.dev/test","hashAlgorithm":"SHA2_256","publicKey":{"rawBytes":"MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAEbfwR+RJudXscgRBRpKX1XFDy3PyudDxz/SfnRi1fT8ekpfBd2O1uoz7jr3Z8nKzxA69EUQ+eFCFI3zeubPWU7w==","keyDetails":"PKIX_ECDSA_P256_SHA_256","validFor":{"start":"2021-03-14T00:00:00Z","end":"2022-10-31T23:59:59.999Z"}},"logId":{"keyId":"CGCS8ChS/2hF0dFrJ4ScRWcYrBY9wzjSbea8IgY2b3I="}},{"baseUrl":"https://ctfe.sigstore.dev/2022","hashAlgorithm":"SHA2_256","publicKey":{"rawBytes":"MFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAEiPSlFi0CmFTfEjCUqF9HuCEcYXNKAaYalIJmBZ8yyezPjTqhxrKBpMnaocVtLJBI1eM3uXnQzQGAJdJ4gs9Fyw==","keyDetails":"PKIX_ECDSA_P256_SHA_256","validFor":{"start":"2022-10-20T00:00:00Z"}},"logId":{"keyId":"3T0wasbHETJjGR4cmWc3AqJKXrjePK3/h4pygC8p7o4="}}],"timestampAuthorities":[{"subject":{"organization":"sigstore.dev","commonName":"sigstore-tsa-selfsigned"},"uri":"https://timestamp.sigstore.dev/api/v1/timestamp","certChain":{"certificates":[{"rawBytes":"MIICEDCCAZagAwIBAgIUOhNULwyQYe68wUMvy4qOiyojiwwwCgYIKoZIzj0EAwMwOTEVMBMGA1UEChMMc2lnc3RvcmUuZGV2MSAwHgYDVQQDExdzaWdzdG9yZS10c2Etc2VsZnNpZ25lZDAeFw0yNTA0MDgwNjU5NDNaFw0zNTA0MDYwNjU5NDNaMC4xFTATBgNVBAoTDHNpZ3N0b3JlLmRldjEVMBMGA1UEAxMMc2lnc3RvcmUtdHNhMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE4ra2Z8hKNig2T9kFjCAToGG30jky+WQv3BzL+mKvh1SKNR/UwuwsfNCg4sryoYAd8E6isovVA3M4aoNdm9QDi50Z8nTEyvqgfDPtTIwXItfiW/AFf1V7uwkbkAoj0xxco2owaDAOBgNVHQ8BAf8EBAMCB4AwHQYDVR0OBBYEFIn9eUOHz9BlRsMCRscsc1t9tOsDMB8GA1UdIwQYMBaAFJjsAe9/u1H/1JUeb4qImFMHic6/MBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMAoGCCqGSM49BAMDA2gAMGUCMDtpsV/6KaO0qyF/UMsX2aSUXKQFdoGTptQGc0ftq1csulHPGG6dsmyMNd3JB+G3EQIxAOajvBcjpJmKb4Nv+2Taoj8Uc5+b6ih6FXCCKraSqupe07zqswMcXJTe1cExvHvvlw=="},{"rawBytes":"MIIB9zCCAXygAwIBAgIUV7f0GLDOoEzIh8LXSW80OJiUp14wCgYIKoZIzj0EAwMwOTEVMBMGA1UEChMMc2lnc3RvcmUuZGV2MSAwHgYDVQQDExdzaWdzdG9yZS10c2Etc2VsZnNpZ25lZDAeFw0yNTA0MDgwNjU5NDNaFw0zNTA0MDYwNjU5NDNaMDkxFTATBgNVBAoTDHNpZ3N0b3JlLmRldjEgMB4GA1UEAxMXc2lnc3RvcmUtdHNhLXNlbGZzaWduZWQwdjAQBgcqhkjOPQIBBgUrgQQAIgNiAAQUQNtfRT/ou3YATa6wB/kKTe70cfJwyRIBovMnt8RcJph/COE82uyS6FmppLLL1VBPGcPfpQPYJNXzWwi8icwhKQ6W/Qe2h3oebBb2FHpwNJDqo+TMaC/tdfkv/ElJB72jRTBDMA4GA1UdDwEB/wQEAwIBBjASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBSY7AHvf7tR/9SVHm+KiJhTB4nOvzAKBggqhkjOPQQDAwNpADBmAjEAwGEGrfGZR1cen1R8/DTVMI943LssZmJRtDp/i7SfGHmGRP6gRbuj9vOK3b67Z0QQAjEAuT2H673LQEaHTcyQSZrkp4mX7WwkmF+sVbkYY5mXN+RMH13KUEHHOqASaemYWK/E"}]},"validFor":{"start":"2025-07-04T00:00:00Z"}}]}\n{"mediaType":"application/vnd.dev.sigstore.trustedroot+json;version=0.1","certificateAuthorities":[{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"fulcio.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICKjCCAbCgAwIBAgIUW3TJVeOvr+NSvJXdOw8nEEn7HhQwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMB4XDTIzMDkxMjE0MDY1NFoXDTI0MDkxMTE0MDY1NFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwyMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEsosodObhuHG6Pr5vp5y+pmnKawS1h2hwv3r3hBwqh3ZHJAw64mhDnDs9fw4jKkZEBYRSVyOHyZppz4day8hgpTIDwdj44Oan4RDb+wmj04jfhVLjLsQ4Q/X4K/ynRgNXo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUm0vkDkQZ29hutYdayJobIRmf/iMwHwYDVR0jBBgwFoAUwOG4UqRLTz7eejgRBs9JjqFFmzMwCgYIKoZIzj0EAwMDaAAwZQIwIBl93E7vkWTvdeIm1WSIM4qNsj0ApE8LCj3k1vrY5x6/7yhAZs7QlO3/FBCoEeaZAjEAlJcNr37uZq9BYHODHBeO/gP+6EfbzsNaLV22ASBlhF/a9y83ESLuqCNN7IxGxmWT"},{"rawBytes":"MIICFTCCAZugAwIBAgIUD3Jlqt4qhrcZI4UnGfPGrEq/pjQwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDkxMTEyMDAwMFoXDTI4MDkwOTEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE7X7nK0wC7uEmDjW+on0sXIX3FacL3hhcrhneA+M/kl1OtvQiPmFrH9lbUQqOj/AfspJ8uGY3jaq8WuSg6ghatzYfuuzLAJIK4nGpCBafncF8EynOssPq64/Dz+JUWXqlo2YwZDAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBATAdBgNVHQ4EFgQUwOG4UqRLTz7eejgRBs9JjqFFmzMwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAI8HWLrke7uzhOpwlD1cNixPmoX9XFKe7bEPozo0D+vKi0Gt6VlC7xPedFIw4/AypAIwQP+FGRWvfx0IAH5/n0aRiN7/LVpyFA5RkJASZOVOib2Y8pNuhXa9V3ZbWO6v6kW/"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2023-10-27T16:30:00Z","end":"2024-05-25T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"fulcio.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICKzCCAbCgAwIBAgIUOpyw2HaZefsj/4SPXutGof8E2CkwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMB4XDTI0MDUxMzAwMDAwMFoXDTI1MDUxMzAwMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwyMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEJNJWvW8fckfk/oQmh+qCeIlFXl9YLEkKSjZCgcVB92Fi1HQnvmpCiyqpvP91SmT1/G6QbrmTGV7MmIQlDnBWHNUT+jwZ3elGu/yfr/v8U0uhZTIli/BMj5Y4ICHK/j4do3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUD0fF3cs+ldPyiWohHJ3JmO91V7gwHwYDVR0jBBgwFoAUwOG4UqRLTz7eejgRBs9JjqFFmzMwCgYIKoZIzj0EAwMDaQAwZgIxAO7BRC9i7oGUHjjlcHU/bfqk2NLy7t6wm3K5W+jBLFbAj6sVjYcY+rrYhop/OjclbQIxALafBKLPIPjoCI29BUHwLBFP6e92ZlyaoFtoqccceXAevRaDjXFvb5+M7wnD6AuAJw=="},{"rawBytes":"MIICFTCCAZugAwIBAgIUD3Jlqt4qhrcZI4UnGfPGrEq/pjQwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDkxMTEyMDAwMFoXDTI4MDkwOTEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE7X7nK0wC7uEmDjW+on0sXIX3FacL3hhcrhneA+M/kl1OtvQiPmFrH9lbUQqOj/AfspJ8uGY3jaq8WuSg6ghatzYfuuzLAJIK4nGpCBafncF8EynOssPq64/Dz+JUWXqlo2YwZDAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBATAdBgNVHQ4EFgQUwOG4UqRLTz7eejgRBs9JjqFFmzMwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAI8HWLrke7uzhOpwlD1cNixPmoX9XFKe7bEPozo0D+vKi0Gt6VlC7xPedFIw4/AypAIwQP+FGRWvfx0IAH5/n0aRiN7/LVpyFA5RkJASZOVOib2Y8pNuhXa9V3ZbWO6v6kW/"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2024-05-13T00:00:00Z","end":"2024-10-25T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"fulcio.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICKzCCAbCgAwIBAgIUQeyd9UH06yZ63pDuqjgUZ58CnpMwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMB4XDTI0MTAwMzEyMDAwMFoXDTI1MTAwMzEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwyMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEwvbET2w+j9j9j50iTInH1gb9GSXkpsCvWz5orX1zgme+/Qh/5gMkpfmgfOSLV2ZRgT1hzujYmnKQvP2mCxYnbwQELAkAf+VhEY/7Uw3zZvguGQSdF1cxzRHiMTOha5eFo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUMib9z4ZYBcQANTVvVCa3KoTGbBUwHwYDVR0jBBgwFoAUwOG4UqRLTz7eejgRBs9JjqFFmzMwCgYIKoZIzj0EAwMDaQAwZgIxAPIU/zlJiJrxn6oTWNdEAD/YBSnhyxcvpq1D2DzFy8E8hbkEfMZPErYL7HyoL/BkdwIxAN9KDEKyktEUBrfHehfcLAzI2kERJx+8DSslXswOIbLaeqYfWsmrQAt5C0X/nOWxXA=="},{"rawBytes":"MIICFTCCAZugAwIBAgIUD3Jlqt4qhrcZI4UnGfPGrEq/pjQwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDkxMTEyMDAwMFoXDTI4MDkwOTEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE7X7nK0wC7uEmDjW+on0sXIX3FacL3hhcrhneA+M/kl1OtvQiPmFrH9lbUQqOj/AfspJ8uGY3jaq8WuSg6ghatzYfuuzLAJIK4nGpCBafncF8EynOssPq64/Dz+JUWXqlo2YwZDAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBATAdBgNVHQ4EFgQUwOG4UqRLTz7eejgRBs9JjqFFmzMwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAI8HWLrke7uzhOpwlD1cNixPmoX9XFKe7bEPozo0D+vKi0Gt6VlC7xPedFIw4/AypAIwQP+FGRWvfx0IAH5/n0aRiN7/LVpyFA5RkJASZOVOib2Y8pNuhXa9V3ZbWO6v6kW/"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2024-10-07T00:00:00Z","end":"2025-06-19T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"fulcio.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICKTCCAbCgAwIBAgIUU7Y2PESb8FK88XCRQvI2rKDB5skwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMB4XDTI1MDUyNzAwMDAwMFoXDTI2MDUyNzAwMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwyMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEevMuOTGccTTXLu65xinRJP4gxkOemyAPCmUXEvhXKtLd+aDgN87PI0QxutOmPuQD2dLZVfxQmt5lvwJpUDIVYkKpH3mDxYbF6MJPMbq+A/EhkMaOS2+mcJRJdY2uam0Vo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQU8tE1L3b7FV3qJKzSqy1w97vPt6AwHwYDVR0jBBgwFoAUwOG4UqRLTz7eejgRBs9JjqFFmzMwCgYIKoZIzj0EAwMDZwAwZAIwSIsnfcDIvp+9SGBe9Es4xLDgZ6/B3wtODu+jtNUznQFPw1/WHa3iBMp3vYPkIqZDAjBBD4AKu166Hg7tDd5E8ckSCNNDkXD0Q7Xdl4yujgqv0goVMeOgRA1pKDGszhunjAQ="},{"rawBytes":"MIICFTCCAZugAwIBAgIUD3Jlqt4qhrcZI4UnGfPGrEq/pjQwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDkxMTEyMDAwMFoXDTI4MDkwOTEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE7X7nK0wC7uEmDjW+on0sXIX3FacL3hhcrhneA+M/kl1OtvQiPmFrH9lbUQqOj/AfspJ8uGY3jaq8WuSg6ghatzYfuuzLAJIK4nGpCBafncF8EynOssPq64/Dz+JUWXqlo2YwZDAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBATAdBgNVHQ4EFgQUwOG4UqRLTz7eejgRBs9JjqFFmzMwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAI8HWLrke7uzhOpwlD1cNixPmoX9XFKe7bEPozo0D+vKi0Gt6VlC7xPedFIw4/AypAIwQP+FGRWvfx0IAH5/n0aRiN7/LVpyFA5RkJASZOVOib2Y8pNuhXa9V3ZbWO6v6kW/"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2025-05-27T00:00:00Z","end":"2025-12-09T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"fulcio.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICKjCCAbCgAwIBAgIUNansGfOzaol+jXgh/wsz0X36zTMwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMB4XDTI1MTExMzAwMDAwMFoXDTI2MTExMzAwMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwyMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE24Un0srydTqXkWtrVrNdgPRNa1cCZqBRpFImYqItuVs/mszkn+RSlwcMzg/LwphU84qwA8Xv2z8ZaIETHspoQgMO4dP5zhJ5W8M+9crL1kR4Co2vS3TJZF9Qpsg+HiTro3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUGO6KtZB8c9gbXS2uNF3tXCtFNhEwHwYDVR0jBBgwFoAUwOG4UqRLTz7eejgRBs9JjqFFmzMwCgYIKoZIzj0EAwMDaAAwZQIxALmOdTRy/SiOc+QSf+eAPz24xWGhMryvNpyNvBZDeyqthC9KaMKmX+kHwJ0A6/bTXQIwE1mPbkH6tBXuWxlR/2Bj7UZ3MADlHSYGQUjpHMdQqJk17UVBywad5U6uA/f0pFdU"},{"rawBytes":"MIICFTCCAZugAwIBAgIUD3Jlqt4qhrcZI4UnGfPGrEq/pjQwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDkxMTEyMDAwMFoXDTI4MDkwOTEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE7X7nK0wC7uEmDjW+on0sXIX3FacL3hhcrhneA+M/kl1OtvQiPmFrH9lbUQqOj/AfspJ8uGY3jaq8WuSg6ghatzYfuuzLAJIK4nGpCBafncF8EynOssPq64/Dz+JUWXqlo2YwZDAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBATAdBgNVHQ4EFgQUwOG4UqRLTz7eejgRBs9JjqFFmzMwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAI8HWLrke7uzhOpwlD1cNixPmoX9XFKe7bEPozo0D+vKi0Gt6VlC7xPedFIw4/AypAIwQP+FGRWvfx0IAH5/n0aRiN7/LVpyFA5RkJASZOVOib2Y8pNuhXa9V3ZbWO6v6kW/"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2025-11-13T00:00:00Z","end":"2026-07-13T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"fulcio.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICKjCCAbCgAwIBAgIUCZdc6761A2K2ny11d+8tvDIlBuAwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMB4XDTI2MDYxMjAwMDAwMFoXDTI3MDYxMjAwMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwyMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE8K+awztD1Q48OQsI2PzffEwHJxdHoKneYSSRgnG6VCk0RlZWeVHM7EwpsDMgMMz98IrRqVBAjxSt7fs39rz6NUKG3T26Khv0gmzuJ1xG0j0BmK0j0kkfRvzYpQKv108to3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwMwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUgeNOwb9e7ggupu+PZL6nr0Wgm88wHwYDVR0jBBgwFoAUwOG4UqRLTz7eejgRBs9JjqFFmzMwCgYIKoZIzj0EAwMDaAAwZQIweEGv0C3qXaA2+oPahLsBzRA2C5+oWNE5Jq1Qpqrn0QeDCSY5eCKzwFjVa5eICks5AjEAsifahLxtAiHc0I2SWHYX3s3sfOLPArgK6oBM0hKbIMCOokEUhss67N06KOuV2rqQ"},{"rawBytes":"MIICFTCCAZugAwIBAgIUD3Jlqt4qhrcZI4UnGfPGrEq/pjQwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDkxMTEyMDAwMFoXDTI4MDkwOTEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZGdWxjaW8gSW50ZXJtZWRpYXRlIGwxMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAE7X7nK0wC7uEmDjW+on0sXIX3FacL3hhcrhneA+M/kl1OtvQiPmFrH9lbUQqOj/AfspJ8uGY3jaq8WuSg6ghatzYfuuzLAJIK4nGpCBafncF8EynOssPq64/Dz+JUWXqlo2YwZDAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBATAdBgNVHQ4EFgQUwOG4UqRLTz7eejgRBs9JjqFFmzMwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAI8HWLrke7uzhOpwlD1cNixPmoX9XFKe7bEPozo0D+vKi0Gt6VlC7xPedFIw4/AypAIwQP+FGRWvfx0IAH5/n0aRiN7/LVpyFA5RkJASZOVOib2Y8pNuhXa9V3ZbWO6v6kW/"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2026-06-12T00:00:00Z"}}],"timestampAuthorities":[{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"timestamp.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICHDCCAaGgAwIBAgIUNDVlmtZuvoujn4KwiC/oxIr8hxAwCgYIKoZIzj0EAwMwMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMB4XDTIzMDgzMTEyMDAwMFoXDTI0MDgzMDEyMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgVGltZXN0YW1waW5nMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEV/zJhNTdu0Fa9hGCUih/JvqEoE81tEWrAVwUXXhdRgIY9hIFErLhNo6sSOpV9d7Zuy0KWMHhcimCUr41a1732ByVRy3f+Z4QhqpsgFMh5b5J90HJLK7HOyUZjehAnvSno3gwdjAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQUGwE6T5ZIh6lY9wP6vt42UHyVMewwHwYDVR0jBBgwFoAUdh+GTP65aetHLVLs9hdhGgDIKIwwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwCgYIKoZIzj0EAwMDaQAwZgIxAJo48LtrSsn0UmLtqGiUKg2EUvso+aDN5EyjpvMmobZ/Oq9zjnR7Of369hoABW4/1gIxANg5ZW4FqijhsXnA3md6jM9yLrLCI9QL+KnuZnXq6WgAcNQaAN7PNNjVDKV3iJEklw=="},{"rawBytes":"MIICJDCCAaqgAwIBAgIUckXVHpiw7iJY1V/jY8LYLj5TgqAwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTI4MDgwNTEyMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEISv96hTQ58QroEzzu4K+o9p8YkwDCBia2U7Y+VBNbOG/w1mLRibve9hSeUE1FSyLBMkiFSSm6MexcsbjyqOoNtRxuMinyYt6DSEox+/It2s/bTPyNAN0QP0DCQQOpnTZo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUdh+GTP65aetHLVLs9hdhGgDIKIwwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAIhf+2E5W2yOb/fCDAjhL/G/jerf74M0tG/zyo32U2keawxkzZosDdwnPaHaGLynAQIwa8nr3en4fZz1AdOZm6nK5hr1qK2F94nifgnAJ/WeT0fZnK/oHan0R28x363qYuYH"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2023-10-27T16:30:00Z","end":"2024-05-25T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"timestamp.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICGzCCAaGgAwIBAgIUPPgn6ner1PU/75CQ+62fdBuazaAwCgYIKoZIzj0EAwMwMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMB4XDTI0MDUxMzAwMDAwMFoXDTI1MDUxMzAwMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgVGltZXN0YW1waW5nMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEA0pG1mAC4qafk1JJuOoIvhnMME9XmBDxjGFreDLnyzaexIzRw+UHUFy8C2gE6Me+0tIGQt4Ftbu66NGmfvBkR6boPMYQSU2O5X5ykZBm/9LR/Aqz0lgmBy/OlXvTJjglo3gwdjAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQUDa4GVhd97Z2V8kiVl9DB0kC53CMwHwYDVR0jBBgwFoAUdh+GTP65aetHLVLs9hdhGgDIKIwwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwCgYIKoZIzj0EAwMDaAAwZQIwEQHd++b7IBAAuqT2/1i/wXf1WM2XrkFF6qd1c3kFcBVvdLQyJ5KoyNUHnfCCVJROAjEA+FoASOEcARlU6RqVcif9JthHwzh6nNwz0AfAHvO8xantN/7HjiLmFrFEGR/g0kN/"},{"rawBytes":"MIICJDCCAaqgAwIBAgIUckXVHpiw7iJY1V/jY8LYLj5TgqAwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTI4MDgwNTEyMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEISv96hTQ58QroEzzu4K+o9p8YkwDCBia2U7Y+VBNbOG/w1mLRibve9hSeUE1FSyLBMkiFSSm6MexcsbjyqOoNtRxuMinyYt6DSEox+/It2s/bTPyNAN0QP0DCQQOpnTZo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUdh+GTP65aetHLVLs9hdhGgDIKIwwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAIhf+2E5W2yOb/fCDAjhL/G/jerf74M0tG/zyo32U2keawxkzZosDdwnPaHaGLynAQIwa8nr3en4fZz1AdOZm6nK5hr1qK2F94nifgnAJ/WeT0fZnK/oHan0R28x363qYuYH"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2024-05-13T00:00:00Z","end":"2024-10-25T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"timestamp.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICGzCCAaGgAwIBAgIUH7swiMTn+svhcDh80OeZccDTj7AwCgYIKoZIzj0EAwMwMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMB4XDTI0MTAwNDEyMDAwMFoXDTI1MTAwNDEyMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgVGltZXN0YW1waW5nMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEM7jdYNBTeD6hjym2/y73b50u2AFQsf8305Sr1NleOqamH9aWt6obhJQH3NoNUw9iFzHcDvafYWQFMu7SmOxS5n3aqwwfR8oJxKnEl36uCmGB+8TXS3B76SVTHEhG5rzOo3gwdjAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQUQvz9YbWX3S6a+jruBkhRBiE2RCkwHwYDVR0jBBgwFoAUdh+GTP65aetHLVLs9hdhGgDIKIwwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwCgYIKoZIzj0EAwMDaAAwZQIxAI4dhu5iyx/g+z1vKAAWvHtebl1ZwsC+Vwgjm6Ttlq5yLNHHvYEnJ/h15Qv2IuXvdgIwZ8H/iy4lXsFJdFYSsB1/zavl24EgxSzxK/pCpihXMetYYDA/lX3xMyquisMx45rN"},{"rawBytes":"MIICJDCCAaqgAwIBAgIUckXVHpiw7iJY1V/jY8LYLj5TgqAwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTI4MDgwNTEyMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEISv96hTQ58QroEzzu4K+o9p8YkwDCBia2U7Y+VBNbOG/w1mLRibve9hSeUE1FSyLBMkiFSSm6MexcsbjyqOoNtRxuMinyYt6DSEox+/It2s/bTPyNAN0QP0DCQQOpnTZo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUdh+GTP65aetHLVLs9hdhGgDIKIwwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAIhf+2E5W2yOb/fCDAjhL/G/jerf74M0tG/zyo32U2keawxkzZosDdwnPaHaGLynAQIwa8nr3en4fZz1AdOZm6nK5hr1qK2F94nifgnAJ/WeT0fZnK/oHan0R28x363qYuYH"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2024-10-07T00:00:00Z","end":"2025-06-19T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"timestamp.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICGzCCAaGgAwIBAgIUaaAaeAO3NzYXR/muq5YMnGY34IUwCgYIKoZIzj0EAwMwMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMB4XDTI1MDUyNzAwMDAwMFoXDTI2MDUyNzAwMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgVGltZXN0YW1waW5nMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEAjyAb+jWITMjkgbjWC02N05Zf5/kPs73ICf5/8lHSyHnoqqItP2z3fkMpWiRpelX8eIjzCjZ0Zht+DpGxtJUNh5jMa+fFv2qoDLGbnQ17DXKXCr9nWHF4vCMflHZK5JHo3gwdjAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQUyvAcv8s1s37A/IWZAwdEptzjfWIwHwYDVR0jBBgwFoAUdh+GTP65aetHLVLs9hdhGgDIKIwwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwCgYIKoZIzj0EAwMDaAAwZQIwMnoHA4qDtR3eRohUIhvHGG70xT9XMfEnuj0kRvi9wu9VumL4pXqz1ktJgTBm6DrYAjEA5SYnRfUm2yAYmGEkJr9RvSCV9O7X+fL8KqfM4TOPrXJLMERQjhRXbWZ4DjxvifnF"},{"rawBytes":"MIICJDCCAaqgAwIBAgIUckXVHpiw7iJY1V/jY8LYLj5TgqAwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTI4MDgwNTEyMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEISv96hTQ58QroEzzu4K+o9p8YkwDCBia2U7Y+VBNbOG/w1mLRibve9hSeUE1FSyLBMkiFSSm6MexcsbjyqOoNtRxuMinyYt6DSEox+/It2s/bTPyNAN0QP0DCQQOpnTZo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUdh+GTP65aetHLVLs9hdhGgDIKIwwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAIhf+2E5W2yOb/fCDAjhL/G/jerf74M0tG/zyo32U2keawxkzZosDdwnPaHaGLynAQIwa8nr3en4fZz1AdOZm6nK5hr1qK2F94nifgnAJ/WeT0fZnK/oHan0R28x363qYuYH"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2025-05-27T00:00:00Z","end":"2025-12-09T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"timestamp.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICGzCCAaGgAwIBAgIUHiljSG9Nx/Qf/gPpl7V1kyRh/jYwCgYIKoZIzj0EAwMwMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMB4XDTI1MTExMzAwMDAwMFoXDTI2MTExMzAwMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgVGltZXN0YW1waW5nMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEG1ZUJgbtNpInC7jR9hi4OMWK5yF5VeRpmCs6ADS5lx23MEHo3ek32O123e3YwH+J3F9Vbpd7ePQHnpdc6tsBcQ/y/NPw2Vf+VwMI/E35C5/3Y9A5kRXq9CFXC1bU6Fj6o3gwdjAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQUXa9JaNrcN+1Je5aHEemVnH2XXtMwHwYDVR0jBBgwFoAUdh+GTP65aetHLVLs9hdhGgDIKIwwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwCgYIKoZIzj0EAwMDaAAwZQIwJyVOay8lriDg4Up0VFWc/o19mzuQeoEItzM2dG50eXF4KbPo9zfdviS3D6StPFRGAjEA5IHCXJ8qP8CsAsQmT3IPimUTAvDNJZI4GFIxs+od/R/EUHGu4NxuwPX3A1ZnJP7z"},{"rawBytes":"MIICJDCCAaqgAwIBAgIUckXVHpiw7iJY1V/jY8LYLj5TgqAwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTI4MDgwNTEyMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEISv96hTQ58QroEzzu4K+o9p8YkwDCBia2U7Y+VBNbOG/w1mLRibve9hSeUE1FSyLBMkiFSSm6MexcsbjyqOoNtRxuMinyYt6DSEox+/It2s/bTPyNAN0QP0DCQQOpnTZo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUdh+GTP65aetHLVLs9hdhGgDIKIwwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAIhf+2E5W2yOb/fCDAjhL/G/jerf74M0tG/zyo32U2keawxkzZosDdwnPaHaGLynAQIwa8nr3en4fZz1AdOZm6nK5hr1qK2F94nifgnAJ/WeT0fZnK/oHan0R28x363qYuYH"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2025-11-13T00:00:00Z","end":"2026-07-13T00:00:00Z"}},{"subject":{"organization":"GitHub, Inc.","commonName":"Internal Services Root"},"uri":"timestamp.githubapp.com","certChain":{"certificates":[{"rawBytes":"MIICGzCCAaGgAwIBAgIUFzPb1un+DvofqFPzAXiB5slHM1IwCgYIKoZIzj0EAwMwMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMB4XDTI2MDYxMjAwMDAwMFoXDTI3MDYxMjAwMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgVGltZXN0YW1waW5nMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAETpok7Tndv57l0YnFjKL+VpcusM8A2xgbTdV353wW8IynvsLTGAjF1BNS4x80TDKHbXxN3jSoSId1u1c3sCWpuEqHB1bbBqCfrlslNn6hyWzhwtiqRLE4LM0T+RPclKQCo3gwdjAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAdBgNVHQ4EFgQUu/GnSzNLJ4IjG9elW5L5a/W61JMwHwYDVR0jBBgwFoAUdh+GTP65aetHLVLs9hdhGgDIKIwwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwCgYIKoZIzj0EAwMDaAAwZQIxANGz5WZBgyZqGnuhr0gDhUW2s4ZXkJ1CMUaFzynjsQ0/8WZghN7Jh04vmT4Vl8y9zgIwJeBHIZNDZtjzDQB0UC+7im4snua3zsBrRbsEtWnqQhDzyT580H/uF4EVQN2kUUtZ"},{"rawBytes":"MIICJDCCAaqgAwIBAgIUckXVHpiw7iJY1V/jY8LYLj5TgqAwCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTI4MDgwNTEyMDAwMFowMjEVMBMGA1UEChMMR2l0SHViLCBJbmMuMRkwFwYDVQQDExBUU0EgaW50ZXJtZWRpYXRlMHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEISv96hTQ58QroEzzu4K+o9p8YkwDCBia2U7Y+VBNbOG/w1mLRibve9hSeUE1FSyLBMkiFSSm6MexcsbjyqOoNtRxuMinyYt6DSEox+/It2s/bTPyNAN0QP0DCQQOpnTZo3sweTAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUdh+GTP65aetHLVLs9hdhGgDIKIwwHwYDVR0jBBgwFoAUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaAAwZQIxAIhf+2E5W2yOb/fCDAjhL/G/jerf74M0tG/zyo32U2keawxkzZosDdwnPaHaGLynAQIwa8nr3en4fZz1AdOZm6nK5hr1qK2F94nifgnAJ/WeT0fZnK/oHan0R28x363qYuYH"},{"rawBytes":"MIIB9TCCAXqgAwIBAgIUNFryA06EHDIcd5EIbe8swbl9OY4wCgYIKoZIzj0EAwMwODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MB4XDTIzMDgwNzEyMDAwMFoXDTMzMDgwNDEyMDAwMFowODEVMBMGA1UEChMMR2l0SHViLCBJbmMuMR8wHQYDVQQDExZJbnRlcm5hbCBTZXJ2aWNlcyBSb290MHYwEAYHKoZIzj0CAQYFK4EEACIDYgAEXYaXx4H0oNuVP/2cfydA3oaafvvkkkgb5hbL8/j/BO25S7uTmDOCA5e4QLLWCKFuc+xp2j14tCH4WmHzMUDvf2tXtInVliY5wZgQMM9L6klo/IwA9x4omdcjnT+kKJAjo0UwQzAOBgNVHQ8BAf8EBAMCAQYwEgYDVR0TAQH/BAgwBgEB/wIBAjAdBgNVHQ4EFgQUfFJ5/6rhfHEZPnXAhrQLhGkJJMwwCgYIKoZIzj0EAwMDaQAwZgIxAPzXsV+eokrqOHSQZH/XhhHE1slOscKy3DQpYpYJ1AWmJ2lJu/XOmubBX5s7apllUwIxALw2Ts8CDACiK42UymC8fk6sbNfoXUAWqdyKTVt2Lst+wNdkRniGvx7jT65BKTkcsQ=="}]},"validFor":{"start":"2026-06-12T00:00:00Z"}}]}\n'


def parse_semver(value):
    """Return numeric core, prerelease, build. ASCII SemVer 2, no signed integers."""
    if not isinstance(value, str):
        raise ValueError('version must be a string')
    match = re.fullmatch(r'(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?', value)
    if match is None:
        raise ValueError(f'invalid SemVer: {value!r}')
    pre = tuple(match[4].split('.')) if match[4] else ()
    if any(part.isascii() and part.isdigit() and len(part) > 1 and part[0] == '0' for part in pre):
        raise ValueError('numeric prerelease identifier has leading zero')
    return tuple(int(match[i]) for i in (1, 2, 3)), pre, tuple(match[5].split('.')) if match[5] else ()


def semver_precedence(value):
    core, pre, _ = parse_semver(value)
    # Numeric identifiers sort before nonnumeric, stable after all prereleases.
    return core, not pre, tuple((0, int(p)) if p.isdigit() else (1, p) for p in pre)


def stable_candidate(value):
    core, pre, build = parse_semver(value)
    # Existing unified family TAG grammar has exactly numeric stable identities.
    if pre or build or core[0] != 0 or core[1] == 0:
        raise ValueError('candidate must be a stable numeric 0.y.z unified family identity with y > 0')
    return core


def select_semver_inventory(name, candidate, payload, *, release_candidate=False):
    """Consume complete unpaginated /crates/NAME/versions JSON, never /crates/NAME."""
    core = stable_candidate(candidate)
    if not isinstance(payload, dict) or 'errors' in payload:
        raise ValueError('inventory is not a successful object')
    rows, meta = payload.get('versions'), payload.get('meta')
    if not isinstance(rows, list) or not isinstance(meta, dict):
        raise ValueError('inventory lacks versions/meta')
    if type(meta.get('total')) is not int or meta['total'] != len(rows) or 'next_page' not in meta or meta['next_page'] is not None:
        raise ValueError('inventory incomplete or paginated')
    ids, versions, selected, same_line = set(), set(), [], []
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError('version row is not an object')
        number, identity = row.get('num'), row.get('id')
        if row.get('crate') != name or type(identity) is not int or identity <= 0 or type(row.get('yanked')) is not bool:
            raise ValueError('foreign/malformed version identity or yank state')
        numeric, pre, _ = parse_semver(number)
        if identity in ids or number in versions:
            raise ValueError('duplicate version or registry identity')
        ids.add(identity)
        versions.add(number)
        if not isinstance(row.get('checksum'), str) or re.fullmatch('[0-9a-f]{64}', row['checksum']) is None:
            raise ValueError('invalid compressed archive checksum')
        if row.get('dl_path') != f'/api/v1/crates/{name}/{number}/download':
            raise ValueError('invalid exact version download path')
        if not pre and numeric[:2] == core[:2]:
            same_line.append(row)
            if numeric <= core:
                selected.append(row)
    if release_candidate:
        if candidate in versions:
            raise ValueError('selected candidate is already published')
        if any(semver_precedence(row['num']) >= semver_precedence(candidate) for row in same_line):
            raise ValueError('candidate must exceed every stable same-line version, including yanked/build identities')
    selected.sort(key=lambda row: (semver_precedence(row['num']), row['num'], row['id']))
    return {'baselines': selected, 'applicability': 'compatible-baselines' if selected else 'empty-compatible-line',
            'candidate_occupied': candidate in versions, 'publication_checked': release_candidate}


def check_no_semver_overrides(package, workspace, query_policy, *, role):
    """Both actual parsed package and workspace tables required, even without inheritance."""
    if not isinstance(package, dict) or not isinstance(workspace, dict):
        raise ValueError(f'{role}: package/workspace facts must be tables (explicit empty workspace allowed)')
    for owner, table in (('package', package), ('workspace', workspace)):
        metadata = table.get('metadata', {})
        if not isinstance(metadata, dict):
            raise ValueError(f'{role}/{owner}: malformed metadata')
        config = metadata.get('cargo-semver-checks', {})
        if not isinstance(config, dict) or set(config) - {'lints'}:
            raise ValueError(f'{role}/{owner}: malformed/unknown SemVer configuration')
        lints = config.get('lints', {})
        if not isinstance(lints, dict):
            raise ValueError(f'{role}/{owner}: malformed SemVer lint table')
        for key in lints:
            if key == 'workspace':
                if lints[key] is not True:
                    raise ValueError(f'{role}/{owner}: explicit false/nonboolean inheritance')
            elif key not in query_policy:
                raise ValueError(f'{role}/{owner}: unknown query ID {key}')
            else:
                raise ValueError(f'{role}/{owner}: user override forbidden: {key}')
    # Standard Rust/Clippy lint tables are not SemVer overrides.
    standard = package.get('lints', {})
    if not isinstance(standard, dict) or ('workspace' in standard and standard['workspace'] is not True):
        raise ValueError(f'{role}/package: invalid standard lint inheritance')
    return {'role': role, 'no_user_suppression': True}


def select_semver_members(scope, validated_families, release_crates, packages):
    """Use existing family validator outputs and actual Cargo metadata target facts."""
    families = list(release_crates) if scope == 'all' else [
        'rpptx' if scope.startswith('rpptx-v') else 'rdocx']
    results = []
    for family in families:
        fact = validated_families[family]
        stable_candidate(fact['version'])
        expected_tag = ('v' if family == 'rdocx' else 'rpptx-v') + fact['version']
        if fact['distribution'] != family or fact['tag'] != expected_tag or scope not in ('all', expected_tag):
            raise ValueError('validated family does not match requested tag')
        for name in release_crates[family]:
            matches = [p for p in packages if p['name'] == name and p['version'] == fact['version']]
            if len(matches) != 1:
                raise ValueError(f'missing/ambiguous current metadata package: {name}')
            kinds = {kind for target in matches[0]['targets'] for kind in target['kind']}
            library = bool(kinds & {'lib', 'rlib', 'dylib', 'cdylib', 'staticlib', 'proc-macro'})
            if not library and 'bin' not in kinds:
                raise ValueError(f'no library or binary target: {name}')
            results.append({'family': family, 'name': name, 'version': fact['version'],
                            'api': 'library' if library else 'binary-only-cli'})
    if len({r['name'] for r in results}) != len(results):
        raise ValueError('conflicting release membership')
    return results


def release_semver_archive(crate: str, version: str, checksum: str, work: Path) -> Path:
    """Acquire one inventory-selected archive into an immutable owned bundle."""
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]*", crate):
        raise ValueError("invalid selected crate")
    if not re.fullmatch(r"(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)\.(?:0|[1-9][0-9]*)(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?", version):
        raise ValueError("invalid selected stable version")
    if not re.fullmatch(r"[0-9a-f]{64}", checksum):
        raise ValueError("invalid selected registry checksum")
    if work.is_symlink() or not work.is_dir():
        raise ValueError("acquisition requires an existing owned run directory")
    work = work.resolve()
    subject = f"{crate}-{version}"
    final = work / subject
    raw = work / f"{subject}.crate"
    observation = work / f"{subject}.acquisition.json"
    if any(os.path.lexists(path) for path in (final, raw, observation)):
        raise ValueError("acquisition subject already exists")
    stage = Path(tempfile.mkdtemp(prefix=f".{subject}-", dir=work))
    receipt = {
        "schema": 1, "crate": crate, "version": version, "expected_sha256": checksum,
        "url": f"https://static.crates.io/crates/{crate}/{subject}.crate",
        "utc": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "status": "INCOMPLETE", "raw_file": raw.name, "received_bytes": 0,
        "owned_stage": stage.name,
        "response_complete": False, "checksum_authenticated": False,
        "redirects": [],
        "redirect_provenance": "observed redirect responses and policy decisions before following",
    }
    class ArchiveRedirectPolicy(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, request, response, code, message, headers, new_url):
            allowed = urllib.parse.urlsplit(new_url).scheme.lower() == "https"
            receipt["redirects"].append({
                "from_url": request.full_url, "status": code, "location": new_url,
                "decision": "allow-https" if allowed else "refuse-non-https",
            })
            if not allowed:
                response.close()
                raise ValueError("archive redirect refused before non-HTTPS request")
            return super().redirect_request(request, response, code, message, headers, new_url)

    published = False
    try:
        os.chmod(stage, 0o700)
        opener = urllib.request.build_opener(ArchiveRedirectPolicy())
        digest = hashlib.sha256()
        with raw.open("xb") as output:
            try:
                response = opener.open(receipt["url"], timeout=60)
            except urllib.error.HTTPError as error:
                receipt["http_status"] = error.code
                receipt["final_url"] = error.geturl()
                error.close()
                raise
            with response:
                receipt["http_status"] = response.status
                receipt["final_url"] = response.geturl()
                if response.status != 200 or not response.geturl().startswith("https://"):
                    raise ValueError("archive response is not complete HTTPS success")
                while True:
                    chunk = response.read(min(1024 * 1024, SEMVER_ARCHIVE_COMPRESSED_LIMIT - receipt["received_bytes"] + 1))
                    if not chunk:
                        break
                    if len(chunk) > SEMVER_ARCHIVE_COMPRESSED_LIMIT - receipt["received_bytes"] + 1:
                        raise ValueError("response exceeded requested bounded read")
                    output.write(chunk)
                    digest.update(chunk)
                    receipt["received_bytes"] += len(chunk)
                    if receipt["received_bytes"] > SEMVER_ARCHIVE_COMPRESSED_LIMIT:
                        raise ValueError("compressed archive budget exceeded, retained prefix only")
        receipt["response_complete"] = True
        receipt["archive_sha256"] = digest.hexdigest()
        if digest.hexdigest() != checksum:
            raise ValueError("archive registry checksum mismatch")
        receipt["checksum_authenticated"] = True
        tar_path = stage / "bounded.tar"
        expanded = 0
        with gzip.open(raw, "rb") as source, tar_path.open("xb") as target:
            while True:
                chunk = source.read(min(1024 * 1024, SEMVER_ARCHIVE_TAR_LIMIT - expanded + 1))
                if not chunk:
                    break
                target.write(chunk)
                expanded += len(chunk)
                if expanded > SEMVER_ARCHIVE_TAR_LIMIT:
                    raise ValueError("decompressed tar budget exceeded")
        receipt["decompressed_tar_bytes"] = expanded
        members = []
        names = {}
        parent_names = set()
        end = 0
        with tarfile.open(tar_path, "r:") as archive:
            for member in archive:
                if len(members) >= SEMVER_ARCHIVE_MEMBER_LIMIT:
                    raise ValueError("logical member budget exceeded")
                directory = member.type == tarfile.DIRTYPE
                if member.type not in (tarfile.REGTYPE, tarfile.AREGTYPE, tarfile.DIRTYPE):
                    raise ValueError("unsupported archive member type")
                if member.sparse is not None or member.pax_headers.get("SCHILY.filetype") == "sparse" or any("sparse" in key.lower() or key == "SCHILY.realsize" for key in member.pax_headers):
                    raise ValueError("sparse archive member")
                # tarfile strips directory suffix slashes. Recover the checked
                # original effective spelling before applying our one-slash rule.
                with tar_path.open("rb") as physical:
                    physical.seek(member.offset)
                    long_name = None
                    while physical.tell() < member.offset_data - 512:
                        header = tarfile.TarInfo.frombuf(physical.read(512), archive.encoding, archive.errors)
                        if header.type not in (tarfile.XHDTYPE, tarfile.XGLTYPE, tarfile.SOLARIS_XHDTYPE,
                                               tarfile.GNUTYPE_LONGNAME, tarfile.GNUTYPE_LONGLINK):
                            raise ValueError("unexpected physical metadata header")
                        if header.size < 0 or header.size > expanded - physical.tell():
                            raise ValueError("metadata lies outside bounded tar")
                        payload_start = physical.tell()
                        if header.type == tarfile.GNUTYPE_LONGNAME:
                            long_name = physical.read(header.size).split(b"\0", 1)[0].decode(archive.encoding, archive.errors)
                        physical.seek(payload_start + ((header.size + 511) // 512) * 512)
                    if physical.tell() != member.offset_data - 512:
                        raise ValueError("physical metadata offset mismatch")
                    raw_header = physical.read(512)
                    checked_header = tarfile.TarInfo.frombuf(raw_header, archive.encoding, archive.errors)
                    raw_name = raw_header[:100].split(b"\0", 1)[0].decode(archive.encoding, archive.errors)
                    suffix_slashes = len(raw_name) - len(raw_name.rstrip("/"))
                    header_name = checked_header.name
                    if directory:
                        header_name += "/" * suffix_slashes
                    name = member.pax_headers.get("path", long_name if long_name is not None else header_name)
                if directory and name.endswith("/"):
                    name = name[:-1]
                parts = name.split("/")
                if any(char in name for char in ("\\", "\0", ":")) or any(part in ("", ".", "..") for part in parts):
                    raise ValueError("noncanonical archive path")
                if name != member.name.removesuffix("/"):
                    raise ValueError("effective metadata name mismatch")
                if parts[0] != subject or (len(parts) == 1 and not directory):
                    raise ValueError("wrong archive root")
                if name in names or (not directory and name in parent_names):
                    raise ValueError("duplicate or file-parent archive conflict")
                for index in range(1, len(parts)):
                    parent = "/".join(parts[:index])
                    if parent in names and not names[parent]:
                        raise ValueError("file-parent archive conflict")
                    parent_names.add(parent)
                if member.size < 0 or (directory and member.size != 0):
                    raise ValueError("invalid archive member length")
                names[name] = directory
                end = max(end, member.offset_data + ((member.size + 511) // 512) * 512)
                if end > expanded:
                    raise ValueError("member lies outside bounded tar")
                members.append((member, tuple(parts)))
            if f"{subject}/Cargo.toml" not in names or names[f"{subject}/Cargo.toml"]:
                raise ValueError("missing root Cargo.toml")
            with tar_path.open("rb") as physical:
                physical.seek(end)
                while chunk := physical.read(1024 * 1024):
                    if any(chunk):
                        raise ValueError("nonzero hidden tar tail")
            source_root = stage / "source"
            source_root.mkdir(mode=0o700)
            actual_root = source_root.resolve()
            directories = {()}
            inventory = []
            for member, parts in members:
                parent_parts = parts if member.isdir() else parts[:-1]
                for index in range(1, len(parent_parts) + 1):
                    key = parent_parts[:index]
                    path = source_root.joinpath(*key)
                    if not path.resolve().is_relative_to(actual_root):
                        raise ValueError("archive path escaped owned source")
                    if key not in directories:
                        path.mkdir(mode=0o700)
                        directories.add(key)
                    if path.is_symlink() or not path.resolve().is_relative_to(actual_root):
                        raise ValueError("owned parent escaped source")
                destination = source_root.joinpath(*parts)
                if member.isdir():
                    inventory.append({"path": "/".join(parts), "type": "directory", "size": 0})
                    continue
                if not destination.resolve().is_relative_to(actual_root):
                    raise ValueError("archive file escaped source")
                mode = 0o600 | (member.mode & 0o100)
                hasher = hashlib.sha256()
                count = 0
                with archive.extractfile(member) as source, os.fdopen(os.open(destination, os.O_WRONLY | os.O_CREAT | os.O_EXCL, mode), "wb") as target:
                    os.chmod(destination, mode)
                    while True:
                        chunk = source.read(min(1024 * 1024, member.size - count + 1))
                        if not chunk:
                            break
                        count += len(chunk)
                        if count > member.size:
                            raise ValueError("member exceeds declared size")
                        target.write(chunk)
                        hasher.update(chunk)
                if count != member.size or not stat.S_ISREG(destination.stat().st_mode):
                    raise ValueError("member length/type mismatch")
                inventory.append({"path": "/".join(parts), "type": "file", "size": count,
                                  "sha256": hasher.hexdigest(), "requested_mode": mode,
                                  "actual_mode": stat.S_IMODE(destination.stat().st_mode)})
        manifest = stage / "source" / subject / "Cargo.toml"
        package = tomllib.loads(manifest.read_text(encoding="utf-8")).get("package", {})
        if package.get("name") != crate or package.get("version") != version:
            raise ValueError("baseline manifest identity mismatch")
        receipt.update(status="READY", members=inventory,
                       original_lock_present=f"{subject}/Cargo.lock" in names,
                       source=f"{subject}/source/{subject}")
        tar_path.unlink()
        (stage / "receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
        if os.path.lexists(final):
            raise ValueError("final acquisition appeared before publication")
        os.rename(stage, final)
        published = True
        return final / "source" / subject
    except Exception as error:
        receipt.update(status="INCOMPLETE", error={"type": type(error).__name__, "message": str(error)})
        if raw.is_file():
            receipt["retained_raw_bytes"] = raw.stat().st_size
            receipt["retained_prefix_sha256"] = hashlib.sha256(raw.read_bytes()).hexdigest()
        cleanup_error = None
        if not published:
            try:
                shutil.rmtree(stage)
                receipt["owned_stage_removed"] = True
            except Exception as failure:
                cleanup_error = failure
                receipt["cleanup_error"] = {"type": type(failure).__name__, "message": str(failure)}
        with observation.open("x", encoding="utf-8") as target:
            target.write(json.dumps(receipt, indent=2) + "\n")
        if cleanup_error is not None:
            raise cleanup_error from error
        raise


def parse_semver_query_protocol(crate, baseline, current, policy, stderr, stdout, exit_code, *, timed_out=False):
    """Only unoverridden stable0.y equal/patch or increasing-minor migration calls."""
    def need(condition, message):
        if not condition:
            raise ValueError(message)

    def core(value):
        match = re.fullmatch(r'0\.([1-9][0-9]*)\.(0|[1-9][0-9]*)(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?', value)
        need(match is not None, 'unsupported nonstable/non0.y version')
        return int(match[1]), int(match[2])

    old, new = core(baseline), core(current)
    need(new >= old, 'backward version movement unsupported')
    need(re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9_-]*', crate) is not None, 'invalid crate identity')
    change = 'major' if new[0] != old[0] else 'minor'
    phrase = 'no change; assume minor' if old == new else change + ' change'
    need(len(policy) == 260 and all(isinstance(v, tuple) and len(v) == 2 and v[0] in ('major', 'minor') and v[1] in ('deny', 'warn', 'allow') for v in policy.values()), 'exact pinned catalog required')
    need(type(exit_code) is int and exit_code in (0, 100) and timed_out is False, 'setup/timeout/unexplained tool exit')
    need(isinstance(stderr, str) and isinstance(stdout, str) and stderr.endswith('\n') and (not stdout or stdout.endswith('\n')), 'truncated/invalid raw stream')
    need('\x1b' not in stderr + stdout, 'unexpected color escape with approved color-never argv')
    selected, skipped = {}, {}
    for query, (required, level) in policy.items():
        reasons = []
        if level == 'allow':
            reasons.append('built-in-allow')
        if change == 'major' or required == 'minor':
            reasons.append('release-level-satisfied')
        if reasons:
            skipped[query] = reasons
        else:
            selected[query] = (required, level)
    stamp = r'\[[ \t]*[0-9]+\.[0-9]{3}s\]'
    phases, results, nested, unclassified, warnings = [], [], [], [], []
    query_pattern = r'(PASS|WARN|FAIL)[ \t]+' + stamp + r'[ \t]+(major|minor)[ \t]+([A-Za-z0-9_]+)'
    for line_number, raw in enumerate(stderr.splitlines()):
        line = raw.strip()
        phase = re.match(r'(Checking|Starting|Checked|Summary|Finished)[ \t]+(.*)$', line)
        if phase:
            kind, detail = phase.groups()
            if kind == 'Checking' and '->' not in detail or kind == 'Finished' and not detail.startswith('['):
                nested.append({'line': line_number, 'raw': raw})
                continue
            phases.append((kind, detail, line_number))
        elif re.match(r'(PASS|WARN|FAIL)(?:\s|$)', line):
            match = re.fullmatch(query_pattern, line)
            need(match is not None, 'malformed verbose query result')
            results.append((*match.groups(), line_number))
        elif re.match(r'Warning(?:\s|$)', line):
            warnings.append((line, line_number))
        else:
            need(re.match(r'(?:warning|error)(?::|\[|\s|$)', line) is None, 'unqualified tool/Cargo warning or error diagnostic')
            unclassified.append({'line': line_number, 'raw': raw})
    need([p[0] for p in phases] == ['Checking', 'Starting', 'Checked', 'Summary', 'Finished'], 'unique ordered five tool phases required')
    details = {name: detail for name, detail, _ in phases}
    need(details['Checking'] == f'{crate} v{baseline} -> v{current} ({phrase})', 'foreign/overridden version classification')
    start = re.fullmatch(r'([0-9]+) checks, ([0-9]+) unnecessary(?: on ([0-9]+) threads)?', details['Starting'])
    need(start is not None and (int(start[1]), int(start[2])) == (len(selected), len(skipped)) and (start[3] is None or int(start[3]) > 1), 'Starting selected/skip/thread counts')
    checked = re.fullmatch(stamp + r' ([0-9]+) checks: ([0-9]+) pass,(?: ([0-9]+) fail, ([0-9]+) warn,)? ([0-9]+) skip', details['Checked'])
    need(checked is not None, 'malformed Checked counts')
    counts = {'selected': int(checked[1]), 'pass': int(checked[2]), 'fail': int(checked[3] or 0), 'warn': int(checked[4] or 0), 'skip': int(checked[5])}
    need(counts['selected'] == len(selected) == counts['pass'] + counts['fail'] + counts['warn'] and counts['skip'] == len(skipped), 'Checked count inconsistency')
    need(len(results) == len({r[2] for r in results}) == len(selected) and {r[2] for r in results} == set(selected), 'missing/duplicate/foreign query IDs')
    for status, required, query, line in results:
        need(phases[1][2] < line < phases[2][2], 'query result outside Starting/Checked interval')
        need(required == selected[query][0], 'query category mismatch')
        need(status == 'PASS' or selected[query][1] == {'WARN': 'warn', 'FAIL': 'deny'}[status], 'query status conflicts with pinned default')
    for status in ('PASS', 'FAIL', 'WARN'):
        need(counts[status.lower()] == sum(r[0] == status for r in results), 'verbose/Checked status disagreement')
    need(bool(checked[3] is not None) == bool(counts['fail'] or counts['warn']), 'unexplained extended Checked form')
    summary = f"semver requires new major version: {counts['fail']} major and 0 minor checks failed" if counts['fail'] else 'no semver update required'
    need(details['Summary'] == summary, 'unexplained Summary')
    need(re.fullmatch(stamp + ' ' + re.escape(crate), details['Finished']) is not None, 'foreign/malformed Finished')
    need(exit_code == (100 if counts['fail'] else 0), 'exit does not match completed query verdict')
    expected_warning = f"Warning produced {counts['warn']} major and 0 minor level warnings"
    need(len(warnings) == bool(counts['warn']) and all(line == expected_warning and phases[3][2] < n < phases[4][2] for line, n in warnings), 'unexplained/missing tool Warning')
    recommendation = 'produced warnings suggest new major version'
    recommendations = [r['line'] for r in unclassified if r['raw'].strip() == recommendation]
    need(len(recommendations) == bool(counts['warn'] and not counts['fail']) and all(warnings[0][1] < n < phases[4][2] for n in recommendations), 'missing/unexplained warning recommendation')
    expected_details = {query: 'failure' if status == 'FAIL' else 'warning' for status, _, query, _ in results if status != 'PASS'}
    sections = []
    lines = stdout.splitlines()
    for number, line in enumerate(lines):
        if line.startswith('--- failure') or line.startswith('--- warning'):
            match = re.fullmatch(r'--- (failure|warning) ([A-Za-z0-9_]+): (.+) ---', line)
            need(match is not None, 'malformed stdout detail heading')
            sections.append((*match.groups(), number))
    need(len(sections) == len(expected_details) and {s[1] for s in sections} == set(expected_details), 'missing/duplicate/foreign stdout detail sections')
    need([s[0] for s in sections] == sorted((s[0] for s in sections)), 'failure details must precede warning details')
    for index, (kind, query, _, number) in enumerate(sections):
        need(kind == expected_details[query], 'stdout detail status disagreement')
        end = sections[index + 1][3] if index + 1 < len(sections) else len(lines)
        body = '\n'.join(lines[number + 1:end])
        need('Description:\n' in body and 'Failed in:\n' in body and body.split('Failed in:\n', 1)[1].strip(), 'incomplete stdout detail body')
        need(f'https://github.com/obi1kenobi/cargo-semver-checks/tree/v0.51.0/src/lints/{query}.ron' in body, 'missing exact stdout detail implementation link')
    if not sections:
        need(not stdout.strip(), 'unexpected stdout without triggered-query details')
    return {'protocol_complete': True, 'clean_query_results': not counts['fail'] and not counts['warn'] and bool(selected),
            'qualification': 'MIGRATION_ALL_SKIPPED_NOT_API_COVERAGE' if not selected else 'QUERY_PROTOCOL_ONLY_NOT_COMPATIBILITY_ACCEPTANCE',
            'classification': change, 'counts': counts,
            'queries': [{'id': query, 'status': status, 'required_update': required} for status, required, query, _ in results],
            'skipped': skipped, 'nested_cargo_statuses': nested, 'unclassified_stderr_lines': unclassified,
            'stdout_details': [{'kind': kind, 'id': query} for kind, query, _, _ in sections]}


def select_semver_attempt(attempts: list[dict]) -> dict:
    """Choose one maximal attempt in one authenticated row/input scope or refuse."""
    if not attempts:
        raise ValueError("required row has no attempts")
    identities = {(a["row_key"], a["input_sha256"]) for a in attempts}
    if len(identities) != 1:
        raise ValueError("ordering requires one exact applicable row/input identity")
    by_id = {}
    timestamps = {}
    for attempt in attempts:
        ident = attempt["id"]
        if not isinstance(ident, str) or not ident or ident in by_id:
            raise ValueError("unique concrete attempt identities required")
        by_id[ident] = attempt
        if attempt["status"] not in {"success", "failed", "canceled", "timed_out", "incomplete", "in_progress", "required_skipped"}:
            raise ValueError("unknown attempt disposition")
        run = (attempt["repository"], attempt["workflow_id"], attempt["run_id"])
        if any(not isinstance(v, str) or not v for v in run):
            raise ValueError("concrete producer scope required")
        if type(attempt["run_attempt"]) is not int or attempt["run_attempt"] < 1:
            raise ValueError("positive actual run_attempt required")
        if not isinstance(attempt["predecessors"], list) or len(set(attempt["predecessors"])) != len(attempt["predecessors"]):
            raise ValueError("unique authorized predecessor references required")
        facts = []
        for field in ("started", "completed"):
            value = attempt[field]
            if value is None:
                facts.append(None)
                continue
            if not isinstance(value, str) or not value.endswith("Z"):
                raise ValueError("authoritative timestamps must be UTC Z strings")
            facts.append(datetime.fromisoformat(value[:-1] + "+00:00"))
        start, end = facts
        if end is not None and (start is None or end < start):
            raise ValueError("invalid actual interval")
        timestamps[ident] = (start, end)
    before = {ident: set() for ident in by_id}
    for ident, attempt in by_id.items():
        for prior in attempt["predecessors"]:
            if prior not in by_id or prior == ident:
                raise ValueError("dangling or self predecessor")
            before[prior].add(ident)
    values = list(by_id.values())
    for left in values:
        for right in values:
            if left is right:
                continue
            same_run = all(left[k] == right[k] for k in ("repository", "workflow_id", "run_id"))
            if same_run:
                if left["run_attempt"] < right["run_attempt"]:
                    before[left["id"]].add(right["id"])
            else:
                _, end = timestamps[left["id"]]
                start, _ = timestamps[right["id"]]
                # Equality does not establish order across concrete runs.
                if end is not None and start is not None and end < start:
                    before[left["id"]].add(right["id"])
    # The finite catalog must be consistent before selecting any green row.
    for _ in by_id:
        for ident in by_id:
            expanded = set(before[ident])
            for following in before[ident]:
                expanded.update(before[following])
            before[ident] = expanded
    if any(ident in following for ident, following in before.items()):
        raise ValueError("contradictory or cyclic attempt order")
    latest = [ident for ident, following in before.items() if not following]
    if len(latest) != 1:
        raise ValueError("incomparable deciding attempts require a real successor")
    selected = by_id[latest[0]]
    if selected["status"] != "success":
        raise ValueError("latest applicable attempt is " + selected["status"])
    return selected


def authenticate_semver_files(root: Path, inventory: dict) -> dict[str, Path]:
    """Check a caller-selected payload subtree. No producer/authority claim follows."""
    if not isinstance(inventory, dict) or not inventory:
        raise ValueError("explicit nonempty retained payload inventory required")
    if root.is_symlink() or not root.is_dir():
        raise ValueError("retained payload root must be a real directory")
    root = root.resolve()
    paths = {}
    native_names = set()
    for name, fact in inventory.items():
        if not isinstance(name, str) or any(c in name for c in ("\\", ":", "\0")):
            raise ValueError("noncanonical retained path")
        parts = name.split("/")
        if any(p in ("", ".", "..") for p in parts):
            raise ValueError("noncanonical retained path")
        folded = tuple(p.casefold() for p in parts)
        if folded in native_names:
            raise ValueError("retained file spelling aliases another payload")
        native_names.add(folded)
        if not isinstance(fact, dict) or set(fact) != {"sha256", "bytes"}:
            raise ValueError("exact retained hash/size fact required")
        digest = fact["sha256"]
        if not isinstance(digest, str) or len(digest) != 64 or any(c not in "0123456789abcdef" for c in digest):
            raise ValueError("invalid retained SHA256")
        if type(fact["bytes"]) is not int or fact["bytes"] < 0:
            raise ValueError("nonnegative retained byte count required")
        path = root
        for index, part in enumerate(parts):
            path = path / part
            mode = path.lstat().st_mode
            if stat.S_ISLNK(mode) or not path.resolve().is_relative_to(root):
                raise ValueError("retained path escaped through symlink")
            if index < len(parts) - 1 and not stat.S_ISDIR(mode):
                raise ValueError("retained parent is not a directory")
        if not stat.S_ISREG(mode):
            raise ValueError("retained payload is not a regular file")
        hasher = hashlib.sha256()
        count = 0
        with path.open("rb") as stream:
            if not stat.S_ISREG(os.fstat(stream.fileno()).st_mode):
                raise ValueError("opened retained payload is not regular")
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                count += len(block)
                if count > fact["bytes"]:
                    raise ValueError("retained payload exceeds declared byte count")
                hasher.update(block)
        if count != fact["bytes"] or hasher.hexdigest() != digest:
            raise ValueError("retained payload hash/size mismatch")
        paths[name] = path
    # Every received payload must be accounted for; the enclosing receipt/seal
    # must live outside this explicitly selected subtree to avoid circular hashes.
    actual = set()
    def refuse_enumeration(error):
        raise error

    for parent, directories, files in os.walk(root, followlinks=False, onerror=refuse_enumeration):
        for name in directories:
            child = Path(parent) / name
            if child.is_symlink():
                raise ValueError("unlisted symlink directory in retained payload")
        for name in files:
            child = Path(parent) / name
            if not stat.S_ISREG(child.lstat().st_mode):
                raise ValueError("nonregular file in retained payload")
            actual.add(child.relative_to(root).as_posix())
    if actual != set(inventory):
        raise ValueError("retained payload inventory is incomplete or has extra entries")
    return paths


SEMVER_CANARY_SOURCES = {
    'legacy-map': b'use std::collections::HashMap;\nuse oxml_opc::OpcPackage;\npub fn old_api(package: &mut OpcPackage) {\n    let parts: &mut HashMap<String, Vec<u8>> = &mut package.parts;\n    parts.insert("/custom/data.bin".into(), vec![1, 2, 3]);\n    let setter: fn(&mut OpcPackage, &str, Vec<u8>) = OpcPackage::set_part;\n    setter(package, "/custom/data.bin", vec![4, 5]);\n    let _removed: Option<Vec<u8>> = package.remove_part("/custom/data.bin");\n}\n',
    'legacy-into': b'use oxml_opc::OpcPackage;\npub fn setter_api(package: &mut OpcPackage) {\n    package.set_part("/custom/data.bin", vec![1, 2, 3].into());\n}\n',
    'published-arc': b'use std::{collections::HashMap, sync::Arc};\nuse oxml_opc::OpcPackage;\npub fn cow_api(package: &mut OpcPackage) {\n    let parts: &mut HashMap<String, Arc<Vec<u8>>> = &mut package.parts;\n    parts.insert("/custom/data.bin".into(), Arc::new(vec![1, 2, 3]));\n    package.set_part("/custom/data.bin", Arc::new(vec![4, 5]));\n    let _removed: Option<Arc<Vec<u8>>> = package.remove_part("/custom/data.bin");\n}\n',
    'declared-arc': b'use std::{collections::HashMap, sync::Arc};\nuse oxml_opc::OpcPackage;\npub fn cow_api(package: &mut OpcPackage) {\n    let parts: &mut HashMap<String, Arc<Vec<u8>>> = &mut package.parts;\n    parts.insert("/custom/data.bin".into(), Arc::new(vec![1, 2, 3]));\n    package.set_part("/custom/data.bin", Arc::new(vec![4, 5]));\n    let _slice: Option<&[u8]> = package.get_part("/custom/data.bin");\n    let _removed: Option<Arc<Vec<u8>>> = package.remove_part("/custom/data.bin");\n}\n',
    'word-nominal': b'pub fn word_nominal(bytes: &[u8], limits: oxml_opc::PackageReadLimits) {\n    let limits: rdocx::PackageReadLimits = limits;\n    let _document = rdocx::Document::from_bytes_with_limits(bytes, limits);\n}\n',
}


SEMVER_CANARY_OUTCOMES = {
    ("legacy-map", "published-0.14.3"): "compile-pass",
    ("legacy-map", "published-0.14.4"): "api-negative",
    ("legacy-into", "published-0.14.3"): "compile-pass",
    ("legacy-into", "published-0.14.4"): "api-negative",
    ("published-arc", "published-0.14.4"): "compile-pass",
    ("declared-arc", "current-candidate"): "compile-pass",
    ("word-nominal", "current-candidate"): "compile-pass",
}


def release_semver_canary_diagnostics(canary: str, client: dict, raw_stdout: bytes, process: dict, *, retained=None, producer_target=None) -> dict:
    """Qualify producer-side Cargo JSON outcome, not graph/native authority."""
    if canary not in SEMVER_CANARY_SOURCES:
        raise ValueError("unknown canary source")
    source = SEMVER_CANARY_SOURCES[canary]
    def absolute(value):
        if retained is None:
            return Path(value).resolve()
        path = _semver_path(value, producer_target)
        return path
    def physical(path):
        if retained is None:
            return Path(path)
        if path not in retained:
            raise ValueError("missing exact retained canary producer-file association")
        return retained[path]
    root = absolute(client["root"])
    source_path = root / "src" / "lib.rs"
    manifest = root / "Cargo.toml"
    if physical(source_path).is_symlink() or physical(manifest).is_symlink() or not physical(manifest).is_file() or physical(source_path).read_bytes() != source:
        raise ValueError("canary source identity mismatch")
    if client["source_sha256"] != hashlib.sha256(source).hexdigest():
        raise ValueError("declared canary source hash mismatch")
    if absolute(process["cwd"]) != root or process["timed_out"] is not False or process["launch_error"] is not None:
        raise ValueError("not a completed canary process at its actual root")
    if type(process["exit"]) is not int:
        raise ValueError("missing actual child exit")
    events = []
    for line in raw_stdout.decode("utf-8").splitlines():
        if not line.strip():
            continue
        try:
            event = json.loads(line)
        except (ValueError, TypeError) as error:
            raise ValueError("malformed Cargo JSON stream") from error
        if not isinstance(event, dict) or not isinstance(event.get("reason"), str):
            raise ValueError("malformed Cargo event")
        events.append(event)
    if not events or events[-1]["reason"] != "build-finished":
        raise ValueError("Cargo stream has no terminal build-finished")
    finished = [event for event in events if event["reason"] == "build-finished"]
    if len(finished) != 1 or type(finished[0].get("success")) is not bool:
        raise ValueError("missing/duplicate Cargo completion")
    allowed_reasons = {"compiler-message", "compiler-artifact", "build-script-executed", "build-finished"}
    if any(event["reason"] not in allowed_reasons for event in events):
        raise ValueError("unqualified Cargo event kind")

    if SEMVER_CANARY_OUTCOMES.get((canary, client["subject"])) != client["expected"]:
        raise ValueError("canary expectation differs from fixed witness policy")
    negative = client["expected"] == "api-negative"
    if client["expected"] not in ("api-negative", "compile-pass") or negative and canary not in ("legacy-map", "legacy-into"):
        raise ValueError("invalid canary outcome policy")
    expected_sites = {
        "legacy-map": {"map": ("E0308", b"&mut package.parts"), "remove": ("E0308", b'package.remove_part("/custom/data.bin")')},
        "legacy-into": {"inference": ("E0283", b"vec![1, 2, 3].into()")},
    }.get(canary, {})
    sites = {}
    for name, (code, marker) in expected_sites.items():
        start = source.index(marker)
        if source.count(marker) != 1:
            raise ValueError("ambiguous fixed canary site")
        sites[name] = (code, start, start + len(marker))
    errors = []
    artifacts = []
    observed = set()
    for event in events:
        if event["reason"] == "compiler-message":
            message = event.get("message")
            if not isinstance(message, dict) or not isinstance(message.get("level"), str):
                raise ValueError("malformed compiler diagnostic")
            if message["level"] != "error":
                continue
            if not negative or event.get("package_id") != client["package_id"]:
                raise ValueError("unexpected canary/dependency compiler error")
            target = event.get("target", {})
            if target.get("name") != client["target_name"] or target.get("kind") != ["lib"] or absolute(target.get("src_path", "")) != source_path:
                raise ValueError("compiler error is not the expected client target")
            code = message.get("code")
            code = code.get("code") if isinstance(code, dict) else None
            matched = set()
            for span in message.get("spans", []):
                if span.get("is_primary") is not True:
                    continue
                file = span.get("file_name")
                if not isinstance(file, str):
                    raise ValueError("missing diagnostic source")
                span_path = (PureWindowsPath(file) if producer_target and producer_target.endswith("msvc") else PurePosixPath(file)) if retained is not None else Path(file)
                span_path = span_path if span_path.is_absolute() else root / span_path
                if absolute(str(span_path)) != source_path:
                    continue
                start, end = span.get("byte_start"), span.get("byte_end")
                if type(start) is not int or type(end) is not int or not (0 <= start < end <= len(source)):
                    raise ValueError("invalid diagnostic byte span")
                line = source[:start].count(b"\n") + 1
                if span.get("line_start") != line or span.get("line_end") != source[:end - 1].count(b"\n") + 1:
                    raise ValueError("diagnostic line/byte disagreement")
                for name, (expected_code, lo, hi) in sites.items():
                    if code == expected_code and lo <= start < end <= hi:
                        matched.add(name)
            if len(matched) != 1 or matched & observed:
                raise ValueError("compiler error does not prove one new expected API site")
            observed.update(matched)
            errors.append({"code": code, "site": next(iter(matched))})
        elif event["reason"] == "compiler-artifact" and event.get("package_id") == client["package_id"]:
            target = event.get("target", {})
            if absolute(event.get("manifest_path", "")) != manifest or target.get("name") != client["target_name"] or target.get("kind") != ["lib"] or target.get("crate_types") != ["lib"] or absolute(target.get("src_path", "")) != source_path:
                raise ValueError("artifact client/manifest/target mismatch")
            if event.get("fresh") is not False or event.get("profile", {}).get("test") is not False:
                raise ValueError("positive canary lacks an actual non-test compilation artifact")
            filenames = event.get("filenames")
            if not isinstance(filenames, list) or not filenames:
                raise ValueError("missing compiler artifact files")
            artifact_root = absolute(client["artifact_root"])
            retained_artifacts = []
            for name in filenames:
                if not isinstance(name, str):
                    raise ValueError("invalid compiler artifact filename")
                path = absolute(name)
                actual_file = physical(path)
                if actual_file.is_symlink() or not actual_file.is_file() or not path.is_relative_to(artifact_root):
                    raise ValueError("compiler artifact is absent/outside owned output")
                raw = actual_file.read_bytes()
                if not raw:
                    raise ValueError("empty compiler artifact")
                retained_artifacts.append({"path": str(path), "sha256": hashlib.sha256(raw).hexdigest(), "bytes": len(raw)})
            if not any(Path(item["path"]).suffix in (".rmeta", ".rlib") for item in retained_artifacts):
                raise ValueError("no Rust library artifact")
            artifacts.append(retained_artifacts)
    if negative:
        if process["exit"] != 101 or finished[0]["success"] is not False or observed != set(sites) or artifacts:
            raise ValueError("expected API negative is incomplete or has unrelated outcome")
    elif process["exit"] != 0 or finished[0]["success"] is not True or errors or len(artifacts) != 1:
        raise ValueError("compile-pass lacks genuine successful build/artifact facts")
    return {"canary": canary, "qualification": "producer JSON outcome only, graph/native/artifact provenance separately required",
            "source_sha256": hashlib.sha256(source).hexdigest(), "outcome": client["expected"],
            "errors": errors, "artifacts": artifacts, "stdout_sha256": hashlib.sha256(raw_stdout).hexdigest(),
            "exit": process["exit"], "build_finished": finished[0]["success"]}


SEMVER_TRUSTED_ROOT_SHA256 = '65ca537f6ed8a47fd0e560c421baa1f6c1efb8b25fc200d8c5c02c0e92eb2b9c'


def qualify_semver_attestation(raw_stdout, process, expected, *, verifier_policy):
    """Check direct output of an authenticated trusted-gh process, never saved JSON alone."""
    def need(condition, message):
        if not condition:
            raise ValueError(message)

    def strict_object(pairs):
        value = {}
        for key, item in pairs:
            need(key not in value, 'duplicate JSON key')
            value[key] = item
        return value

    def zoned_time(value):
        need(isinstance(value, str) and re.fullmatch(r'[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]{1,9})?(?:Z|[+-][0-9]{2}:[0-9]{2})', value) is not None, 'invalid verified RFC3339 timestamp')
        time = datetime.fromisoformat(value.replace('Z', '+00:00'))
        need(time.year > 1 and time.utcoffset() is not None, 'zero/naive verified timestamp')
        return value

    for key in ('packet_sha256', 'bundle_sha256'):
        need(isinstance(expected[key], str) and re.fullmatch('[0-9a-f]{64}', expected[key]) is not None, 'invalid expected packet/bundle hash')
    for key in ('source_digest', 'signer_digest', 'build_config_digest'):
        need(isinstance(expected[key], str) and re.fullmatch('[0-9a-f]{40}', expected[key]) is not None, 'invalid independently expected commit identity')
    for key in ('run_id', 'run_attempt'):
        need(type(expected[key]) is int and expected[key] > 0, 'positive actual run/attempt required')
    for key in ('source_ref', 'signer_ref'):
        need(isinstance(expected[key], str) and expected[key].startswith('refs/') and not any(c.isspace() for c in expected[key]), 'invalid expected Git ref')
    for key in ('signer_workflow', 'source_workflow'):
        need(isinstance(expected[key], str) and re.fullmatch(r'\.github/workflows/[A-Za-z0-9_-]+\.ya?ml', expected[key]) is not None, 'reviewed workflow path required')
    need(isinstance(expected['event_name'], str) and expected['event_name'] and isinstance(expected['packet_subject_name'], str) and expected['packet_subject_name'], 'concrete expected event/subject required')
    need(type(process['exit']) is int and process['exit'] == 0 and process['timed_out'] is False and process['launch_error'] is None, 'verifier process failed/incomplete')
    # Policy comes from the deciding caller, never the evidence packet or expected claims.
    need(isinstance(verifier_policy, dict) and verifier_policy.get('implementation') == 'gh' and verifier_policy.get('version') == '2.97.0', 'unsupported/missing independently selected verifier policy')
    verifier_sha = verifier_policy.get('executable_sha256')
    need(isinstance(verifier_sha, str) and re.fullmatch('[0-9a-f]{64}', verifier_sha) is not None and isinstance(verifier_policy.get('consumer_host'), str) and verifier_policy['consumer_host'], 'invalid independent verifier integrity/host policy')
    need(process['executable_sha256'] == verifier_sha and process['verifier_implementation'] == verifier_policy['implementation'] and process['verifier_version'] == verifier_policy['version'] and process['consumer_host'] == verifier_policy['consumer_host'], 'actual verifier identity differs from independent consumer policy')
    need(process['trusted_root_sha256'] == SEMVER_TRUSTED_ROOT_SHA256, 'unqualified independent root snapshot')
    need(process['packet_sha256'] == expected['packet_sha256'] and process['bundle_sha256'] == expected['bundle_sha256'], 'actual verifier inputs differ from expected packet/bundle')
    need(isinstance(raw_stdout, bytes) and process['stdout_sha256'] == hashlib.sha256(raw_stdout).hexdigest(), 'raw verifier output binding mismatch')
    for key in ('executable', 'packet_path', 'bundle_path', 'trusted_root_path'):
        need(isinstance(process[key], str) and process[key], 'actual verifier input paths required')
    need(len({process[k] for k in ('packet_path','bundle_path','trusted_root_path')}) == 3, 'verifier input paths alias')
    repo = 'tensorbee/rdocx'
    repo_uri = 'https://github.com/' + repo
    signer_uri = repo_uri + '/' + expected['signer_workflow'] + '@' + expected['signer_ref']
    config_uri = repo_uri + '/' + expected['source_workflow'] + '@' + expected['source_ref']
    invocation = repo_uri + f"/actions/runs/{expected['run_id']}/attempts/{expected['run_attempt']}"
    argv = [process['executable'], 'attestation', 'verify', process['packet_path'], '-R', repo,
            '--bundle', process['bundle_path'], '--custom-trusted-root', process['trusted_root_path'],
            '--signer-workflow', repo + '/' + expected['signer_workflow'], '--signer-digest', expected['signer_digest'],
            '--source-digest', expected['source_digest'], '--source-ref', expected['source_ref'],
            '--cert-oidc-issuer', 'https://token.actions.githubusercontent.com', '--deny-self-hosted-runners', '--hostname', 'github.com', '--format', 'json']
    need(process['argv'] == argv, 'actual verifier argv differs from approved local-bundle/root policy')
    data = json.loads(raw_stdout.decode('utf-8'), object_pairs_hook=strict_object)
    need(isinstance(data, list) and data, 'expected actual gh verified-result array')
    certificate_policy = {'issuer': 'https://token.actions.githubusercontent.com',
        'subjectAlternativeName': signer_uri, 'buildSignerURI': signer_uri,
        'buildSignerDigest': expected['signer_digest'], 'runnerEnvironment': 'github-hosted',
        'sourceRepositoryURI': repo_uri, 'sourceRepositoryDigest': expected['source_digest'],
        'sourceRepositoryRef': expected['source_ref'], 'sourceRepositoryIdentifier': '1164240333',
        'sourceRepositoryOwnerURI': 'https://github.com/tensorbee', 'sourceRepositoryOwnerIdentifier': '263075749',
        'buildConfigURI': config_uri, 'buildConfigDigest': expected['build_config_digest'],
        'buildTrigger': expected['event_name'], 'runInvocationURI': invocation}
    matches = []
    for index, result in enumerate(data):
        need(isinstance(result, dict) and isinstance(result.get('verificationResult'), dict), 'malformed verifier result')
        verified = result['verificationResult']
        statement = verified.get('statement')
        need(isinstance(statement, dict) and isinstance(statement.get('subject'), list), 'missing verified statement subjects')
        subjects = statement['subject']
        packet_subjects = [s for s in subjects if isinstance(s, dict) and s.get('digest') == {'sha256': expected['packet_sha256']} and s.get('name') == expected['packet_subject_name']]
        if not packet_subjects:
            continue
        need(len(subjects) == len(packet_subjects) == 1, 'packet must be the sole signed subject')
        signature = verified.get('signature')
        certificate = signature.get('certificate') if isinstance(signature, dict) else None
        need(isinstance(certificate, dict) and all(isinstance(certificate.get(k), str) and certificate[k] for k in certificate_policy), 'missing verified certificate identity fields')
        if any(certificate[k] != value for k, value in certificate_policy.items()):
            continue
        matches.append((index, verified, statement, certificate))
    need(len(matches) == 1, 'expected exactly one full signed-subject/certificate-context match')
    index, verified, statement, certificate = matches[0]
    need(verified.get('mediaType') == 'application/vnd.dev.sigstore.verificationresult+json;version=0.1', 'wrong verifier result schema')
    need(statement.get('_type') == 'https://in-toto.io/Statement/v1' and statement.get('predicateType') == 'https://slsa.dev/provenance/v1', 'wrong verified statement/predicate type')
    timestamps = verified.get('verifiedTimestamps')
    need(isinstance(timestamps, list) and timestamps, 'missing verified signed timestamps')
    for stamp in timestamps:
        need(isinstance(stamp, dict) and stamp.get('type') in ('Tlog','TimestampAuthority') and isinstance(stamp.get('uri'), str) and stamp['uri'], 'unqualified timestamp observation')
        zoned_time(stamp.get('timestamp'))
    # Workflow-controlled predicate only corroborates already selected certificate facts.
    predicate = statement.get('predicate')
    need(isinstance(predicate, dict), 'missing corroborating predicate')
    definition, details = predicate.get('buildDefinition'), predicate.get('runDetails')
    need(isinstance(definition, dict) and isinstance(details, dict), 'missing predicate definition/run details')
    need(definition.get('buildType') == 'https://actions.github.io/buildtypes/workflow/v1', 'wrong predicate build type')
    external, internal, metadata = definition.get('externalParameters'), definition.get('internalParameters'), details.get('metadata')
    need(isinstance(external, dict) and isinstance(internal, dict) and isinstance(metadata, dict), 'malformed corroborating predicate tables')
    need(external.get('workflow') == {'ref': expected['source_ref'], 'repository': repo_uri, 'path': expected['source_workflow']}, 'predicate workflow contradicts certificate/policy')
    need(internal.get('github') == {'event_name': expected['event_name'], 'repository_id':'1164240333','repository_owner_id':'263075749','runner_environment':'github-hosted'}, 'predicate github claims contradict certificate/policy')
    need(definition.get('resolvedDependencies') == [{'uri': 'git+' + repo_uri + '@' + expected['source_ref'], 'digest': {'gitCommit': expected['source_digest']}}], 'predicate source contradicts certificate/policy')
    need(details.get('builder') == {'id': signer_uri} and metadata.get('invocationId') == invocation, 'predicate signer/invocation contradicts certificate')
    return {'qualification':'TRUSTED_VERIFIER_OUTPUT_QUALIFIER_ONLY_NOT_JOB_NATIVE_OR_LATEST_AUTHORITY',
            'matching_result_index':index,'verified_certificate':certificate,'verified_timestamps':timestamps,
            'packet_sha256':expected['packet_sha256'],'verifier_stdout_sha256':process['stdout_sha256'],
            'run_invocation_uri':invocation,'trusted_root_sha256':SEMVER_TRUSTED_ROOT_SHA256,
            'verifier_sha256':verifier_sha,'verifier_implementation':verifier_policy['implementation'],
            'verifier_version':verifier_policy['version'],'verifier_consumer_host':verifier_policy['consumer_host'],
            'certificate_binding_qualified':True}


SEMVER_ARCHIVE_COMPRESSED_LIMIT = 16 * 1024 * 1024
SEMVER_ARCHIVE_TAR_LIMIT = 128 * 1024 * 1024
SEMVER_ARCHIVE_MEMBER_LIMIT = 4096




def _semver_hash(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def _semver_write(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x', encoding='utf-8') as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write('\n')


def _semver_read(path):
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError('duplicate SemVer JSON key')
            result[key] = value
        return result
    return json.loads(Path(path).read_text(encoding='utf-8'), object_pairs_hook=unique)


def _semver_identity(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(',', ':')).encode()).hexdigest()


def _semver_host():
    pair = (sys.platform, platform.machine().lower())
    choices = {('linux', 'x86_64'): 'x86_64-unknown-linux-gnu',
               ('darwin', 'arm64'): 'aarch64-apple-darwin',
               ('win32', 'amd64'): 'x86_64-pc-windows-msvc'}
    if pair not in choices:
        raise ValueError(f'unsupported genuine native SemVer consumer: {pair}')
    return choices[pair]


def _semver_mode(args):
    for name in ('target', 'tool', 'output', 'check_evidence'):
        if getattr(args, name) is not None and len(getattr(args, name)) != 1:
            raise ValueError(f'--{name.replace("_", "-")} must occur once')
    if args.refresh_registry > 1 or args.release_candidate > 1:
        raise ValueError('SemVer mode flag repeated')
    if args.scope != 'all':
        match = re.fullmatch(r'(v|rpptx-v)(.+)', args.scope)
        if match is None:
            raise ValueError('SemVer scope must be all or one unified family tag')
        stable_candidate(match[2])
    execute = args.check_evidence is None
    if execute:
        if not all((args.target, args.tool, args.output)) or args.refresh_registry or args.release_candidate:
            raise ValueError('execute requires target/tool/output and forbids evidence-only flags')
    elif args.target or args.tool or args.output or args.reuse_evidence:
        raise ValueError('evidence mode forbids execute-only arguments')
    if args.release_candidate and args.scope == 'all':
        raise ValueError('release-candidate requires one explicit family tag')
    directories = args.reuse_evidence if execute else args.check_evidence[0]
    if any(not directory.strip() for directory in directories):
        raise ValueError('empty SemVer evidence directory')
    if execute and any(not getattr(args, key)[0].strip() for key in ('target', 'tool', 'output')):
        raise ValueError('empty SemVer execute argument')
    return {'execute': execute, 'directories': [Path(d).resolve() for d in directories],
            'fresh_registry': execute or bool(args.refresh_registry or args.release_candidate)}


def _semver_scope(scope):
    workspace = tomllib.loads((REPO / 'Cargo.toml').read_text())
    versions = {'rdocx': workspace['workspace']['package']['version'],
                'rpptx': tomllib.loads((REPO / 'crates/rpptx/Cargo.toml').read_text())['package']['version']}
    families = {}
    for family, version in versions.items():
        tag = ('v' if family == 'rdocx' else 'rpptx-v') + version
        if scope == 'all' or scope == tag:
            families[family] = validate_unified_release_family(tag)
    if not families:
        raise ValueError('SemVer tag does not match current validated family carriers')
    for family in families:
        for name in RELEASE_CRATES[family]:
            manifest=tomllib.loads((REPO/'crates'/name/'Cargo.toml').read_text())
            check_no_semver_overrides(manifest['package'],workspace['workspace'],SEMVER_QUERY_POLICY,role='current/'+name)
    return {'scope': scope, 'families': families,
            'packages': {name: {'family': family, 'version': fact['version']}
                         for family, fact in families.items() for name in RELEASE_CRATES[family]}}


def _semver_source():
    names = subprocess.run(['git', 'ls-files', '-z'], cwd=REPO, check=True, capture_output=True).stdout.decode().split('\0')
    relevant = {}
    for name in names:
        if not name:
            continue
        if (name in ('Cargo.toml', 'Cargo.lock', 'rust-toolchain.toml', 'scripts/sprint_workflow.py',
                     '.github/workflows/ci.yml', '.github/workflows/wheels.yml', '.claude/plans/F-X275-design.md') or name.startswith('.claude/reviews/F-X275-') or name.startswith('.cargo/')
                or name.startswith('crates/')):
            path = REPO / name
            if path.is_symlink() or not path.is_file():
                raise ValueError(f'SemVer source is absent/nonregular: {name}')
            relevant[name] = _semver_hash(path)
    if not relevant:
        raise ValueError('SemVer requires actual source identity')
    return {'files': relevant, 'sha256': _semver_identity(relevant), 'git_blobs': _semver_git_blobs(relevant)}


def _semver_policy():
    return _semver_identity({'tool': SEMVER_NATIVE_TOOLS, 'queries': SEMVER_QUERY_POLICY,
        'verifiers': SEMVER_GH_VERIFY_TOOLS, 'root': SEMVER_TRUSTED_ROOT_SHA256,
        'canaries': {k: hashlib.sha256(v).hexdigest() for k,v in SEMVER_CANARY_SOURCES.items()},
        'script': _semver_hash(Path(__file__)), 'schema': 1})


def _semver_run(argv, cwd, env, destination, name, timeout=3600):
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    stdout, stderr = destination / (name + '.stdout'), destination / (name + '.stderr')
    row = {'argv': [str(a) for a in argv], 'cwd': str(Path(cwd).resolve()),
           'started': datetime.now(timezone.utc).isoformat().replace('+00:00','Z'), 'exit': None,
           'timed_out': False, 'launch_error': None}
    start = time.monotonic()
    with stdout.open('xb') as out, stderr.open('xb') as err:
        try:
            kwargs = {'start_new_session': True} if os.name != 'nt' else {'creationflags': subprocess.CREATE_NEW_PROCESS_GROUP}
            child = subprocess.Popen(row['argv'], cwd=cwd, env=env, stdout=out, stderr=err, **kwargs)
            row['pid'] = child.pid
            try:
                row['exit'] = child.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                row['timed_out'] = True
                if os.name == 'nt':
                    subprocess.run(['taskkill', '/PID', str(child.pid), '/T', '/F'], capture_output=True, timeout=30)
                else:
                    try:
                        os.killpg(child.pid, signal.SIGTERM)
                    except ProcessLookupError:
                        pass
                    try:
                        child.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        os.killpg(child.pid, signal.SIGKILL)
                row['exit'] = child.wait(timeout=30)
        except (OSError, subprocess.SubprocessError) as error:
            row['launch_error'] = type(error).__name__ + ': ' + str(error)
    row.update(completed=datetime.now(timezone.utc).isoformat().replace('+00:00','Z'), seconds=time.monotonic()-start,
               stdout_sha256=_semver_hash(stdout), stderr_sha256=_semver_hash(stderr),
               stdout_file=stdout.name, stderr_file=stderr.name)
    _semver_write(destination / (name + '.process.json'), row)
    return row


def _semver_success(row):
    if row['exit'] != 0 or row['timed_out'] or row['launch_error'] is not None:
        raise ValueError(f"actual SemVer command failed/incomplete: {row['argv']} exit={row['exit']}")


def _semver_configs(cwd):
    cargo_home = Path(os.environ.get('CARGO_HOME', str(Path.home()/'.cargo'))).resolve()
    paths = {p/'.cargo'/n for p in [Path(cwd).resolve(), *Path(cwd).resolve().parents, Path.home()] for n in ('config','config.toml')}
    paths.update(cargo_home/n for n in ('config','config.toml'))
    facts = {}
    for path in sorted(paths):
        if path.is_file():
            data = tomllib.loads(path.read_text())
            if any(k in data.get('build', {}) for k in ('target-dir','build-dir','rustc','rustdoc','rustc-wrapper','rustc-workspace-wrapper')) or any(k in data.get('env', {}) for k in ('CARGO_TARGET_DIR','CARGO_BUILD_TARGET_DIR','CARGO_BUILD_BUILD_DIR','RUSTC','RUSTDOC','PATH','RUSTC_WRAPPER','RUSTC_WORKSPACE_WRAPPER')):
                raise ValueError('shared Cargo target/build directory configuration forbids two-side retention')
            facts[str(path)] = {'present': True, 'sha256': _semver_hash(path)}
        else:
            facts[str(path)] = {'present': False}
    return facts


def _semver_environment():
    for key in ('CARGO_TARGET_DIR','CARGO_BUILD_TARGET_DIR','CARGO_BUILD_BUILD_DIR','RUSTC_WRAPPER','RUSTC_WORKSPACE_WRAPPER','CARGO_BUILD_RUSTC','CARGO_BUILD_RUSTDOC','CARGO_BUILD_RUSTC_WRAPPER','CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER'):
        if os.environ.get(key):
            raise ValueError('unqualified shared target/compiler wrapper: ' + key)
    env = dict(os.environ)
    env['CARGO_TERM_COLOR'] = 'never'
    keys = ('RUSTFLAGS','RUSTDOCFLAGS','CARGO_ENCODED_RUSTFLAGS','CARGO_ENCODED_RUSTDOCFLAGS',
            'RUSTUP_TOOLCHAIN','RUSTC','RUSTDOC','CARGO_BUILD_TARGET','CARGO_NET_OFFLINE')
    facts = {key: env.get(key) for key in keys}
    facts['other_cargo_rust_env_sha256'] = {key: hashlib.sha256(value.encode()).hexdigest()
        for key,value in env.items() if key.startswith(('CARGO_','RUST')) and key not in keys}
    facts['configs'] = _semver_configs(REPO)
    return env, facts


def _semver_keep(source, root, relative):
    source, destination = Path(source), Path(root)/relative
    if source.is_symlink() or not source.is_file() or not destination.resolve().is_relative_to(Path(root).resolve()):
        raise ValueError('invalid SemVer retained file association')
    destination.parent.mkdir(parents=True, exist_ok=True)
    with source.open('rb') as stream, destination.open('xb') as output:
        shutil.copyfileobj(stream, output)
    if _semver_hash(source) != _semver_hash(destination):
        raise ValueError('SemVer copy changed during retention')
    return {'path': relative, 'sha256': _semver_hash(destination), 'bytes': destination.stat().st_size,
            'producer_path': str(source.resolve())}


def _semver_inventory_files(root):
    root = Path(root)
    inventory = {}
    def refuse(error):
        raise error
    for parent, dirs, files in os.walk(root, followlinks=False, onerror=refuse):
        if any((Path(parent)/name).is_symlink() for name in dirs):
            raise ValueError('symlink directory in SemVer payload')
        for name in files:
            path = Path(parent)/name
            if path.is_symlink() or not stat.S_ISREG(path.lstat().st_mode):
                raise ValueError('nonregular SemVer payload')
            inventory[path.relative_to(root).as_posix()] = {'sha256': _semver_hash(path), 'bytes': path.stat().st_size}
    return inventory


def _semver_http(url, destination):
    """Retain actual HTTPS response, including errors, before JSON interpretation."""
    destination = Path(destination)
    class HTTPSOnly(urllib.request.HTTPRedirectHandler):
        def redirect_request(self, request, response, code, message, headers, new_url):
            if urllib.parse.urlsplit(new_url).scheme != 'https':
                response.close()
                raise ValueError('SemVer refused non-HTTPS redirect')
            return super().redirect_request(request,response,code,message,headers,new_url)
    if urllib.parse.urlsplit(url).scheme != 'https':
        raise ValueError('SemVer acquisition requires HTTPS')
    row = {'url': url, 'utc': datetime.now(timezone.utc).isoformat().replace('+00:00','Z'), 'complete': False}
    try:
        request = urllib.request.Request(url, headers={'User-Agent':'rdocx-semver-policy (tensorbee/rdocx)','Accept':'application/json'})
        response = urllib.request.build_opener(HTTPSOnly()).open(request, timeout=60)
        with response, destination.open('xb') as output:
            row.update(status=response.status, final_url=response.geturl())
            shutil.copyfileobj(response, output)
        row.update(complete=True, sha256=_semver_hash(destination), bytes=destination.stat().st_size)
    except urllib.error.HTTPError as error:
        row.update(status=error.code, final_url=error.geturl(), error=str(error))
        with error, destination.open('xb') as output:
            shutil.copyfileobj(error, output)
        row.update(sha256=_semver_hash(destination), bytes=destination.stat().st_size)
    except (OSError, ValueError) as error:
        row['error'] = type(error).__name__ + ': ' + str(error)
        if destination.exists():
            row.update(sha256=_semver_hash(destination), bytes=destination.stat().st_size)
    _semver_write(destination.with_suffix(destination.suffix+'.http.json'), row)
    if not row['complete'] or row.get('status') != 200:
        raise ValueError('required SemVer HTTP acquisition failed/incomplete: ' + url)
    return row


def _semver_registry(scope, destination, publication=False):
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    rows = {}
    for name, package in scope['packages'].items():
        raw = destination/(name+'.json')
        _semver_http(f'https://crates.io/api/v1/crates/{name}/versions', raw)
        payload = _semver_read(raw)
        rows[name] = {'raw_sha256': _semver_hash(raw), 'raw_file': raw.name,
                      **select_semver_inventory(name, package['version'], payload, release_candidate=publication)}
    if publication:
        for family, fact in scope['families'].items():
            raw = destination/(family+'-pypi.json')
            _semver_http(f'https://pypi.org/pypi/{family}/json', raw)
            info = _semver_read(raw)
            if fact['version'] in info.get('releases', {}):
                raise ValueError('selected Python candidate is already published')
    return rows


def _prepare_release_semver_tools(target, directory):
    """Authorized online caller setup. Frozen evidence consumption never calls it."""
    if target != _semver_host():
        raise ValueError('tool preparation requires its genuine native host')
    directory = Path(directory)
    directory.mkdir(mode=0o700)
    result = {}
    for kind, policy in (('tool',SEMVER_NATIVE_TOOLS[target]),('gh',SEMVER_GH_VERIFY_TOOLS[target])):
        url = policy['archive_url']
        archive = directory/urllib.parse.urlsplit(url).path.rsplit('/',1)[1]
        _semver_http(url, archive)
        if _semver_hash(archive) != policy['archive_sha256']:
            raise ValueError('official native tool archive checksum mismatch')
        member = policy.get('archive_member', policy.get('executable_member'))
        if archive.suffix == '.zip':
            with zipfile.ZipFile(archive) as bundle:
                names = bundle.namelist()
                if names.count(member) != 1:
                    raise ValueError('missing/duplicate pinned executable archive member')
                data = bundle.read(member)
        else:
            with tarfile.open(archive,'r:gz') as bundle:
                entries = [m for m in bundle if m.name == member]
                if len(entries) != 1 or not entries[0].isfile() or entries[0].issparse():
                    raise ValueError('missing/invalid pinned executable archive member')
                with bundle.extractfile(entries[0]) as stream:
                    data = stream.read()
        if hashlib.sha256(data).hexdigest() != policy['executable_sha256']:
            raise ValueError('pinned native executable checksum mismatch')
        executable = directory/('gh.exe' if target.endswith('msvc') else 'gh') if kind == 'gh' else directory/('cargo-semver-checks.exe' if target.endswith('msvc') else 'cargo-semver-checks')
        with executable.open('xb') as stream:
            stream.write(data)
        executable.chmod(0o700)
        result[kind] = str(executable.resolve())
    return result


def _semver_tool_identity(executable, target, policy):
    executable = Path(executable).resolve()
    archive = executable.parent/urllib.parse.urlsplit(policy['archive_url']).path.rsplit('/',1)[1]
    if executable.is_symlink() or not executable.is_file() or _semver_hash(executable) != policy['executable_sha256'] or not archive.is_file() or archive.is_symlink() or _semver_hash(archive) != policy['archive_sha256']:
        raise ValueError('native executable/archive is not independently pinned')
    return {'target': target,'executable':str(executable),'executable_sha256':_semver_hash(executable),
            'archive':str(archive),'archive_sha256':_semver_hash(archive),'archive_url':policy['archive_url']}



def _semver_authority_location(directories):
    sprint, _features = parse_current_sprint()
    path = state_path(sprint)
    for candidate in (SCRATCH, path):
        if candidate.is_symlink():
            raise ValueError('SemVer authority state cannot be a symlink')
    resolved = path.resolve()
    for directory in directories:
        directory = Path(directory).resolve()
        if resolved.is_relative_to(directory) or SCRATCH.resolve().is_relative_to(directory):
            raise ValueError('SemVer authority state aliases supplied evidence')
    return sprint, path


def _load_release_semver_authority(sprint, scope, directories=()):
    resolved_sprint, path = _semver_authority_location(directories)
    if sprint != resolved_sprint or not path.is_file():
        raise ValueError('independent SemVer authority context is absent')
    data = load(sprint)
    context = data.get('semver_authority')
    if not isinstance(context, dict) or context.get('status') != 'complete':
        raise ValueError('independent SemVer authority context is absent/incomplete')
    source = _semver_source()
    if context.get('source_identity') != source['sha256'] or context.get('policy_identity') != _semver_policy() or context.get('scope') != scope:
        raise ValueError('SemVer authority source/policy/scope context differs')
    if context.get('purpose') == 'sign':
        raise ValueError('signing-only authority cannot decide SemVer evidence')
    catalog_path = Path(context['catalog_path'])
    if catalog_path.is_symlink() or not catalog_path.is_file():
        raise ValueError('independent SemVer authority catalog is absent/aliased')
    for directory in directories:
        if catalog_path.resolve().is_relative_to(Path(directory).resolve()):
            raise ValueError('SemVer authority catalog is inside supplied evidence')
    if _semver_hash(catalog_path) != context['catalog_sha256']:
        raise ValueError('independently expected SemVer authority catalog changed')
    catalog = _semver_read(catalog_path)
    if catalog.get('context') != {key:context[key] for key in ('source_identity','policy_identity','scope','cutoff')}:
        raise ValueError('SemVer catalog does not match independent state scope/cutoff')
    authenticate_semver_files(catalog_path.parent/'raw', catalog['raw_inventory'])
    if catalog.get('repository_id') != 1164240333 or catalog.get('owner_id') != 263075749 or not isinstance(catalog.get('attempts'), list):
        raise ValueError('SemVer authority repository/catalog schema differs')
    return context, catalog


def cmd_release_semver(args):
    try:
        mode = _semver_mode(args)
    except ValueError as error:
        print('release-semver usage: ' + str(error), file=sys.stderr)
        return 2
    try:
        scope = _semver_scope(args.scope)
        if mode['execute']:
            return _execute_release_semver(args, scope, mode)
        sprint, _path = _semver_authority_location(mode['directories'])
        context, catalog = _load_release_semver_authority(sprint, scope, mode['directories'])
        result = _consume_release_semver(scope, mode['directories'], context, catalog,
            refresh_registry=mode['fresh_registry'], release_candidate=bool(args.release_candidate))
        print(json.dumps(result, sort_keys=True))
        return 0
    except (ValueError, OSError, KeyError, TypeError, subprocess.SubprocessError) as error:
        print('release-semver refused: ' + str(error), file=sys.stderr)
        return 1



def _semver_native_setup(tool, target, payload, env):
    setup = Path(payload)/'setup'
    setup.mkdir()
    cargo_home=Path(env.get('CARGO_HOME',str(Path.home()/'.cargo'))).absolute()
    canonical_bin=cargo_home/'bin'
    def identity(path):
        if path is None:return {'path':None}
        path=Path(path)
        try:
            return {'path':str(path.absolute()),'resolved_path':str(path.resolve()),'sha256':_semver_hash(path),'bytes':path.stat().st_size}
        except OSError as error:
            return {'path':str(path.absolute()),'error':type(error).__name__+': '+str(error)}
    inherited={name:shutil.which(name,path=env.get('PATH')) for name in ('rustc','cargo','rustdoc','rustup')}
    _semver_write(setup/'tool-resolution-input.json',{'cargo_home':str(cargo_home),'inherited':{name:identity(path) for name,path in inherited.items()}})
    sysroot = _semver_run(['rustc','--print','sysroot'], REPO, env, setup, 'sysroot', 120)
    _semver_success(sysroot)
    rust = Path((setup/'sysroot.stdout').read_text().strip())/'bin'
    binaries = {}
    for name, options in (('rustc',['-Vv']),('cargo',['-V']),('rustdoc',['-V'])):
        executable = rust/(name+'.exe' if os.name == 'nt' else name)
        if executable.is_symlink() or not executable.is_file():
            raise ValueError('missing concrete native Rust executable')
        if name.upper() in env and env[name.upper()] and Path(env[name.upper()]).resolve() != executable.resolve():
            raise ValueError('Rust executable override differs from actual sysroot')
        row = _semver_run([executable,*options], REPO, env, setup, name, 120)
        _semver_success(row)
        text = (setup/(name+'.stdout')).read_text()
        if not text.startswith(name+' 1.97.1'):
            raise ValueError('SemVer compiler toolchain is not pinned Rust1.97.1')
        if name == 'rustc' and 'host: '+target not in text:
            raise ValueError('SemVer target is not the actual native compiler host')
        binaries[name] = {'path':str(executable.resolve()),'sha256':_semver_hash(executable),'version':text}
    resolutions={}
    suffix='.exe' if os.name=='nt' else ''
    for name,fact in binaries.items():
        path=inherited[name]
        if path is None:
            raise ValueError('missing inherited inner tool resolution: '+name)
        actual=Path(path).resolve()
        resolution={'inherited':identity(path),'qualified_path':fact['path'],'sha256':fact['sha256'],'kind':'direct'}
        if 'sha256' not in resolution['inherited']:
            raise ValueError('inherited tool identity could not be retained: '+name)
        if actual!=Path(fact['path']):
            rustup=inherited['rustup']
            if Path(path).absolute()!=canonical_bin/(name+suffix) or rustup is None or Path(rustup).absolute()!=canonical_bin/('rustup'+suffix):
                raise ValueError('bare inner tool differs from canonical Rustup/qualified toolchain: '+name)
            for mode,arguments in (('current',['which',name]),('pinned',['which','--toolchain','1.97.1',name])):
                process=_semver_run([rustup,*arguments],REPO,env,setup,'resolve-'+mode+'-'+name,120)
                _semver_success(process)
                selected=Path((setup/('resolve-'+mode+'-'+name+'.stdout')).read_text().strip()).resolve()
                if selected!=Path(fact['path']):
                    raise ValueError('rustup '+mode+' selection differs from qualified compiler: '+name)
            rustup_identity=identity(rustup)
            if 'sha256' not in rustup_identity:
                raise ValueError('canonical Rustup identity could not be retained')
            resolution.update(kind='canonical-rustup-proxy',rustup=rustup_identity)
        options=['-Vv'] if name=='rustc' else ['-V']
        bare=_semver_run([path,*options],REPO,env,setup,'inherited-'+name,120)
        _semver_success(bare)
        if (setup/('inherited-'+name+'.stdout')).read_text()!=fact['version']:
            raise ValueError('actual inherited tool output differs from qualified compiler: '+name)
        resolutions[name]=resolution
    env['PATH']=str(rust)+os.pathsep+env.get('PATH','')
    env['RUSTC']=binaries['rustc']['path'];env['RUSTDOC']=binaries['rustdoc']['path']
    for name in binaries:
        if Path(shutil.which(name,path=env['PATH'])).resolve()!=Path(binaries[name]['path']):
            raise ValueError('prepared bare tool resolution is not qualified')
    _semver_write(setup/'inner-tools.json',resolutions)
    version = _semver_run([tool,'--version'],REPO,env,setup,'tool',120)
    _semver_success(version)
    if not (setup/'tool.stdout').read_text().startswith('cargo-semver-checks 0.51.0'):
        raise ValueError('wrong pinned API tool version')
    return binaries


def _semver_metadata(manifest, cwd, cargo, env, destination, name, *, no_deps=False):
    command = [cargo,'metadata','--locked','--format-version','1','--manifest-path',manifest]
    if no_deps:
        command.append('--no-deps')
    else:
        command.append('--offline')
    row = _semver_run(command,cwd,env,destination,name,300)
    _semver_success(row)
    return _semver_read(Path(destination)/(name+'.stdout'))


def _semver_external_work(prefix):
    root = Path(tempfile.mkdtemp(prefix=prefix)).resolve()
    try:
        if root.is_relative_to(REPO.resolve()) or any((p/'Cargo.toml').exists() for p in (root,*root.parents)):
            raise ValueError('SemVer working root has Cargo workspace ancestry')
    except BaseException:
        shutil.rmtree(root)
        raise
    return root


def _semver_path(value, target):
    path = PureWindowsPath(value) if target.endswith('msvc') else PurePosixPath(value)
    if not path.is_absolute() or '..' in path.parts:
        raise ValueError('invalid absolute producer path binding')
    return path



def _semver_row_identity(kind, crate, version, baseline, view, common, extra=None):
    facts = {'kind':kind,'crate':crate,'version':version,'baseline':baseline,'view':view,
             'target':common['target'],'policy':common['policy'],
             'compiler':{k:{'sha256':v['sha256'],'version':v['version']} for k,v in common['rust'].items()},
             'environment':common['environment_identity'],'extra':extra}
    if kind == 'api' or baseline == 'current-candidate':
        facts['current_source'] = common['source']['sha256']
    return facts, _semver_identity(facts)


def _semver_generated_side(source_root, target_root, crate, version, view, common, row_directory, role, before_projects):
    source_root, target_root = Path(source_root).resolve(), Path(target_root).resolve()
    candidates = []
    for manifest in (target_root/'semver-checks').glob('*/Cargo.toml'):
        state=_semver_project_state(manifest)
        if before_projects.get(str(manifest.resolve()))==state:
            continue
        data = tomllib.loads(manifest.read_text())
        dependency = data.get('dependencies',{}).get(crate)
        if data.get('package',{}).get('name') != 'placeholder' or not isinstance(dependency,dict) or not dependency.get('path'):
            continue
        if Path(dependency['path']).resolve() != source_root:
            continue
        if dependency.get('default-features',True) is not True:
            continue
        extra = dependency.get('features',[])
        if not isinstance(extra,list) or any(not isinstance(f,str) for f in extra):
            raise ValueError('malformed generated feature request')
        if view == 'default' and extra:
            continue
        candidates.append((manifest,data,dependency))
    qualified = []
    for manifest,data,dependency in candidates:
        # Match original tool cwd/config semantics, not the generated project cwd.
        before = _semver_hash(manifest.parent/'Cargo.lock')
        observed = _semver_metadata(manifest,REPO,common['rust']['cargo']['path'],common['env'],row_directory,role+'-metadata') if len(candidates)==1 else None
        if observed is None:
            # Multiple previous feature slugs must be selected by actual requested features.
            package_manifest = tomllib.loads((source_root/'Cargo.toml').read_text())
            declared = set(package_manifest.get('features',{}))
            declared.update(name for name,dep in package_manifest.get('dependencies',{}).items() if isinstance(dep,dict) and dep.get('optional') and not any('dep:'+name in values for values in package_manifest.get('features',{}).values()))
            if view == 'all' and set(dependency.get('features',[])) == declared:
                observed = _semver_metadata(manifest,REPO,common['rust']['cargo']['path'],common['env'],row_directory,role+'-metadata')
            else:
                continue
        if _semver_hash(manifest.parent/'Cargo.lock') != before:
            raise ValueError('auxiliary generated metadata changed actual tool lock')
        packages = [p for p in observed['packages'] if p['name']==crate and Path(p['manifest_path']).resolve()==source_root/'Cargo.toml']
        if len(packages)!=1 or packages[0]['version'] != version or packages[0]['source'] is not None:
            raise ValueError('generated metadata does not bind exact source/version')
        package = packages[0]
        declared = set(package['features'])
        if view == 'all' and set(dependency.get('features',[])) != declared:
            continue
        libraries = [t for t in package['targets'] if 'lib' in t['kind']]
        if len(libraries)!=1:
            raise ValueError('subject lacks a unique actual library target')
        node = next(n for n in observed['resolve']['nodes'] if n['id']==package['id'])
        if not set(dependency.get('features',[])).issubset(node['features']):
            raise ValueError('explicit generated features missing from effective side')
        target = Path(observed['target_directory']).resolve()
        if not target.is_relative_to(manifest.parent.resolve()):
            raise ValueError('generated target is shared/outside its actual project')
        doc = target/common['target']/'doc'/(libraries[0]['name'].replace('-','_')+'.json')
        document = _semver_read(doc)
        if document['crate_version'] != version or document['index'][str(document['root'])]['name'] != libraries[0]['name']:
            raise ValueError('actual rustdoc subject/version identity differs')
        local = row_directory/role
        local.mkdir()
        refs = {name:_semver_keep(path,row_directory,role+'/'+name) for name,path in (
            ('Cargo.toml',manifest),('Cargo.lock',manifest.parent/'Cargo.lock'),('lib.rs',manifest.parent/'lib.rs'),('subject.json',doc))}
        qualified.append({'source_root':str(source_root),'project':str(manifest.parent.resolve()),'target':str(target),'json_path':str(doc),
            'manifest':refs['Cargo.toml'],'lock':refs['Cargo.lock'],'lib':refs['lib.rs'],'json':refs['subject.json'],
            'metadata':_semver_existing_ref(row_directory,role+'-metadata.stdout'),
            'metadata_process':_semver_existing_ref(row_directory,role+'-metadata.process.json'),
            'effective_features':node['features'],'available_features':sorted(declared),'requested_features':dependency.get('features',[]),
            'subject_package_id':package['id'],'library':libraries[0]['name'],'json_format':document['format_version'],
            'generation_before':before_projects.get(str(manifest.resolve())),'generation_after':_semver_project_state(manifest),
            'configuration_cwd':str(REPO.resolve()),'metadata_qualification':'post-tool locked/offline observation, original cwd/config and unchanged generated lock'})
    if len(qualified)!=1:
        raise ValueError('no unique fully qualified actual generated side')
    return qualified[0]


def _semver_api_row(crate, version, baseline, archive_source, view, common, row_directory):
    work = _semver_external_work('rdocx-semver-row-')
    metadata=None;baseline_metadata=None;retain_work=False
    ancestry={str(parent/'Cargo.toml'):(parent/'Cargo.toml').exists() for parent in (work,*work.parents)}
    try:
        baseline_root = work/'baseline'
        shutil.copytree(archive_source,baseline_root)
        original = _semver_inventory_files(baseline_root)
        source_root = REPO/'crates'/crate
        manifest = source_root/'Cargo.toml'
        current_manifest = tomllib.loads(manifest.read_text())
        workspace = tomllib.loads((REPO/'Cargo.toml').read_text())['workspace']
        baseline_manifest = tomllib.loads((baseline_root/'Cargo.toml').read_text())
        check_no_semver_overrides(current_manifest['package'],workspace,SEMVER_QUERY_POLICY,role='current')
        check_no_semver_overrides(baseline_manifest['package'],baseline_manifest.get('workspace',{}),SEMVER_QUERY_POLICY,role='baseline')
        metadata = _semver_metadata(manifest,REPO,common['rust']['cargo']['path'],common['env'],row_directory,'current-source',no_deps=True)
        baseline_metadata = _semver_metadata(baseline_root/'Cargo.toml',REPO,common['rust']['cargo']['path'],common['env'],row_directory,'baseline-source',no_deps=True)
        projects={role:_semver_snapshot_projects(observed['target_directory']) for role,observed in [('current',metadata),('baseline',baseline_metadata)]}
        argv = [common['tool']['executable'],'check-release','--manifest-path',str(manifest.relative_to(REPO)),
                '--package',crate,'--baseline-root',str(baseline_root),'--target',common['target'],
                '--default-features' if view=='default' else '--all-features','--color','never','--verbose']
        process = _semver_run(argv,REPO,common['env'],row_directory,'check-release')
        stdout = (row_directory/'check-release.stdout').read_text()
        stderr = (row_directory/'check-release.stderr').read_text()
        builds = re.findall(r'^\s*Building '+re.escape(crate)+r' v([^ ]+) \((current|baseline)\)',stderr,re.M)
        if builds != [(version,'current'),(baseline,'baseline')]:
            raise ValueError('actual current-before-baseline build identity/order differs')
        sides = {'current':_semver_generated_side(source_root,metadata['target_directory'],crate,version,view,common,row_directory,'current',projects['current']),
                 'baseline':_semver_generated_side(baseline_root,baseline_metadata['target_directory'],crate,baseline,view,common,row_directory,'baseline',projects['baseline'])}
        if any(sides['current'][key] == sides['baseline'][key] for key in ('project','target','json_path')):
            raise ValueError('current/baseline generated artifacts alias')
        protocol = parse_semver_query_protocol(crate,baseline,version,SEMVER_QUERY_POLICY,stderr,stdout,process['exit'],timed_out=process['timed_out'])
        if process['launch_error'] or not protocol['clean_query_results']:
            raise ValueError('API row has failed/warned/incomplete or all-skipped query coverage')
        after = {name:{'sha256':_semver_hash(baseline_root/name),'bytes':(baseline_root/name).stat().st_size} for name in original}
        if {k:v for k,v in original.items() if k!='Cargo.lock'} != {k:v for k,v in after.items() if k!='Cargo.lock'}:
            raise ValueError('published non-lock source changed during tool operation')
        lock_disposition = {'original':original.get('Cargo.lock'),'working_after':after.get('Cargo.lock'),'changed':original.get('Cargo.lock')!=after.get('Cargo.lock')}
        if lock_disposition['changed']:
            # Retain the real mutation, but do not pretend it is the published lock or a qualified graph.
            if (baseline_root/'Cargo.lock').is_file():
                lock_disposition['retained_after'] = _semver_keep(baseline_root/'Cargo.lock',row_directory,'baseline-working-Cargo.lock')
            raise ValueError('working baseline original lock changed; graph qualification is incomplete')
        return {'process':process,'protocol':protocol,'sides':sides,'baseline_lock':lock_disposition,
                'baseline_work_root':str(work),'baseline_ancestor_cargo':ancestry,
                'path_roles':[{'role':role+'-'+kind,'producer_prefix':side['source_root' if kind=='source' else 'project']}
                     for role,side in sides.items() for kind in ('source','generated')]}
    except BaseException as error:
        roots=[('baseline',work/'baseline',Path(baseline_metadata['target_directory']) if baseline_metadata else work/'baseline'/'target'),
               ('current',REPO/'crates'/crate,Path(metadata['target_directory']) if metadata else REPO/'target')]
        retain_work=True
        _semver_preserve_failed_generated(roots,row_directory,common['target'],error)
        retain_work=False
        raise
    finally:
        if not retain_work:
            shutil.rmtree(work)


def _semver_canary_row(canary, subject, source_root, common, row_directory):
    work = _semver_external_work('rdocx-semver-canary-')
    retain_work=False
    try:
        (work/'src').mkdir()
        (work/'src/lib.rs').write_bytes(SEMVER_CANARY_SOURCES[canary])
        crate = 'rdocx' if canary=='word-nominal' else 'oxml-opc'
        dependencies = {crate: {'path':str(source_root),'version':'='+common['scope']['packages'][crate]['version']}} if subject=='current-candidate' else {'oxml-opc':{'path':str(source_root),'version':'='+subject.removeprefix('published-')}}
        if canary=='word-nominal':
            dependencies['oxml-opc'] = {'path':str(REPO/'crates/oxml-opc'),'version':'='+common['scope']['packages'].get('oxml-opc',{'version':common['shared_version']})['version']}
        text = '[package]\nname="rdocx-semver-canary"\nversion="0.0.0"\nedition="2024"\n[workspace]\n[dependencies]\n'
        for name,fact in dependencies.items():
            text += name+' = '+json.dumps(fact).replace(': ', ' = ')+'\n'
        # JSON object syntax differs from TOML inline tables only in separators.
        (work/'Cargo.toml').write_text(text)
        client_configs = _semver_configs(work)
        before_source = _semver_inventory_files(source_root) if subject!='current-candidate' else None
        client = {'root':str(work),'source_sha256':hashlib.sha256(SEMVER_CANARY_SOURCES[canary]).hexdigest(),
                  'target_name':'rdocx_semver_canary','artifact_root':str(work/'target'),'subject':subject,
                  'expected':SEMVER_CANARY_OUTCOMES[(canary,subject)]}
        process = _semver_run([common['rust']['cargo']['path'],'check','--manifest-path',work/'Cargo.toml',
            '--target',common['target'],'--message-format=json'],work,common['env'],row_directory,'cargo-check')
        client_lock_before=_semver_hash(work/'Cargo.lock')
        lock_before_fetch=_semver_keep(work/'Cargo.lock',row_directory,'client-before-fetch/Cargo.lock')
        fetch=_semver_run([common['rust']['cargo']['path'],'fetch','--locked','--manifest-path',work/'Cargo.toml'],work,common['env'],row_directory,'client-fetch',300)
        _semver_success(fetch)
        if _semver_hash(work/'Cargo.lock')!=client_lock_before or _semver_configs(work)!=client_configs:
            raise ValueError('all-target cache acquisition changed client lock/configuration')
        metadata = _semver_metadata(work/'Cargo.toml',work,common['rust']['cargo']['path'],common['env'],row_directory,'client-metadata')
        if _semver_hash(work/'Cargo.lock')!=client_lock_before:
            raise ValueError('post-check metadata changed external client lock')
        packages = [p for p in metadata['packages'] if Path(p['manifest_path']).resolve()==work/'Cargo.toml']
        if len(packages)!=1:
            raise ValueError('missing genuine external canary package')
        client['package_id'] = packages[0]['id']
        libraries = [t for t in packages[0]['targets'] if 'lib' in t['kind']]
        if len(libraries)!=1:
            raise ValueError('canary library target is ambiguous')
        client['target_name'] = libraries[0]['name']
        opc = [p for p in metadata['packages'] if p['name']=='oxml-opc']
        if len(opc)!=1 or opc[0]['source'] is not None or Path(opc[0]['manifest_path']).resolve() != (REPO/'crates/oxml-opc/Cargo.toml' if subject=='current-candidate' else Path(source_root)/'Cargo.toml').resolve():
            raise ValueError('canary dependency has a different/duplicate nominal OPC identity')
        actual = release_semver_canary_diagnostics(canary,client,(row_directory/'cargo-check.stdout').read_bytes(),process)
        if _semver_configs(work) != client_configs:
            raise ValueError('external client Cargo configuration changed')
        if before_source is not None and _semver_inventory_files(source_root) != before_source:
            raise ValueError('published canary source changed')
        files = {'source':_semver_keep(work/'src/lib.rs',row_directory,'client/src/lib.rs'),
                 'manifest':_semver_keep(work/'Cargo.toml',row_directory,'client/Cargo.toml'),
                 'lock':_semver_keep(work/'Cargo.lock',row_directory,'client/Cargo.lock')}
        artifact_refs = []
        for group in actual['artifacts']:
            for artifact in group:
                artifact_refs.append(_semver_keep(artifact['path'],row_directory,'client/artifacts/'+Path(artifact['path']).name))
        return {'process':process,'diagnostics':actual,'client':client,'files':files,'artifacts':artifact_refs,
                'metadata_file':'client-metadata.stdout','metadata':_semver_existing_ref(row_directory,'client-metadata.stdout'),
                'metadata_process':_semver_existing_ref(row_directory,'client-metadata.process.json'),
                'fetch':{'process':_semver_existing_ref(row_directory,'client-fetch.process.json'),'stdout':_semver_existing_ref(row_directory,'client-fetch.stdout'),'stderr':_semver_existing_ref(row_directory,'client-fetch.stderr'),'lock_before':lock_before_fetch},'feature_mode':'default','configs':client_configs,
                'source_root':str(source_root),'client_root':str(work)}
    except BaseException as error:
        retain_work=True
        _semver_preserve_failed_client(work,row_directory,error)
        retain_work=False
        raise
    finally:
        if not retain_work:
            shutil.rmtree(work)


def _semver_freeze_packet(output, receipt):
    output = Path(output)
    receipt['payload_inventory'] = _semver_inventory_files(output/'payload')
    _semver_write(output/'receipt.json',receipt)
    packet = output/('semver-'+receipt['target']+'.tar')
    with tarfile.open(packet,'x:') as bundle:
        for name in ['receipt.json',*['payload/'+n for n in sorted(receipt['payload_inventory'])]]:
            bundle.add(output/name,arcname=name,recursive=False)
    _semver_write(output/'packet.json',{'file':packet.name,'sha256':_semver_hash(packet),'bytes':packet.stat().st_size})
    return packet


def _execute_release_semver(args, scope, mode):
    target, tool, output = args.target[0], args.tool[0], Path(args.output[0])
    if target!=_semver_host():
        raise ValueError('execute target requires genuine matching native host')
    tool_identity = _semver_tool_identity(tool,target,SEMVER_NATIVE_TOOLS[target])
    if os.path.lexists(output):
        raise ValueError('SemVer output must be a new immutable directory')
    env, environment = _semver_environment()
    source = _semver_source()
    output.mkdir(mode=0o700,parents=True)
    payload = output/'payload';payload.mkdir()
    receipt = {'schema':1,'scope':scope,'source':source,'policy':_semver_policy(),'target':target,
        'tool':tool_identity,'environment':environment,'rows':{},'applicability':{},'complete':False,
        'qualification':'NATIVE_SEMVER_PRODUCER_NOT_PUBLICATION_AUTHORITY','errors':[],
        'producer':{'repository':os.environ.get('GITHUB_REPOSITORY'),'run_id':os.environ.get('GITHUB_RUN_ID'),
                    'run_attempt':os.environ.get('GITHUB_RUN_ATTEMPT'),'source_head':git_head(),
                    'ref':os.environ.get('GITHUB_REF'),'event':os.environ.get('GITHUB_EVENT_NAME')}}
    external = None
    archives = {}
    keep_external = False
    reused = {}
    try:
        external = _semver_external_work('rdocx-semver-archives-')
        rust = _semver_native_setup(tool_identity['executable'],target,payload,env)
        metadata = _semver_metadata(REPO/'Cargo.toml',REPO,rust['cargo']['path'],env,payload,'workspace',no_deps=True)
        receipt['members'] = select_semver_members(scope['scope'],scope['families'],RELEASE_CRATES,metadata['packages'])
        receipt['rust'] = rust
        registry = _semver_registry(scope,payload/'registry')
        receipt['registry'] = registry
        environment_identity = {key:value for key,value in environment.items() if key!='configs'}
        environment_identity['present_configs'] = {path:fact for path,fact in environment['configs'].items() if fact['present']}
        common = {'target':target,'tool':tool_identity,'rust':rust,'env':env,'environment_identity':environment_identity,
                  'source':source,'scope':scope,'policy':receipt['policy'],
                  'shared_version':tomllib.loads((REPO/'crates/rpptx/Cargo.toml').read_text())['package']['version']}
        reused = _semver_execute_reuse(scope,mode['directories']) if mode['directories'] else {}
        def baseline(crate, version, checksum):
            nonlocal keep_external
            key = crate+'-'+version
            if key not in archives:
                try:
                    root = release_semver_archive(crate,version,checksum,external)
                except BaseException:
                    keep_external = True
                    for filename in (key+'.crate',key+'.acquisition.json'):
                        if (external/filename).is_file():
                            _semver_keep(external/filename,payload,'failed-acquisitions/'+filename)
                    keep_external = False
                    raise
                bundle = root.parent.parent
                proof = payload/'archives'/key
                shutil.copytree(bundle,proof)
                _semver_keep(external/(key+'.crate'),payload,'archives/'+key+'/published.crate')
                archives[key] = root
            return archives[key]
        for member in receipt['members']:
            name, version = member['name'], member['version']
            if member['api']=='binary-only-cli':
                receipt['applicability'][name] = {'kind':'binary-only-cli','api_tool':'not-applicable','existing_cli_behavior_and_version_gates':'required separately'}
                continue
            selected = registry[name]['baselines']
            receipt['applicability'][name] = {'kind':'library','api_tool':registry[name]['applicability'],'baselines':[r['num'] for r in selected]}
            for published in selected:
                baseline_root=baseline(name,published['num'],published['checksum'])
                for view in ('default','all'):
                    facts,key = _semver_row_identity('api',name,version,published['num'],view,common,{'archive_sha256':published['checksum']})
                    row_directory=payload/'rows'/key;row_directory.mkdir(parents=True)
                    row={'key':key,'kind':'api','input':facts,'status':'incomplete'}
                    try:
                        if key in reused:
                            row = _semver_copy_reused_row(reused[key],row_directory,facts)
                        else:
                            result=_semver_api_row(name,version,published['num'],baseline_root,view,common,row_directory)
                            row.update(status='success',result=result)
                    except (ValueError,OSError,KeyError,TypeError,subprocess.SubprocessError) as error:
                        row['error']=type(error).__name__+': '+str(error)
                    _semver_write(row_directory/'row.json',row)
                    receipt['rows'][key]={'path':'rows/'+key+'/row.json','input':facts,'status':row['status']}
        # Historical opposing contracts are real fixed witnesses, not same-line API baselines.
        historical = None
        canary_roles = [(canary,subject) for canary,subject in SEMVER_CANARY_OUTCOMES
                        if subject!='current-candidate' or canary=='declared-arc' and 'rpptx' in scope['families'] or canary=='word-nominal' and 'rdocx' in scope['families']]
        for canary,subject in canary_roles:
            if subject=='current-candidate':
                name='rdocx' if canary=='word-nominal' else 'oxml-opc';version=scope['packages'][name]['version'];checksum=None
                canary_root=REPO/'crates'/name
            else:
                name='oxml-opc';version=subject.removeprefix('published-')
                if historical is None:
                    raw=payload/'registry/historical-opc.json';_semver_http('https://crates.io/api/v1/crates/oxml-opc/versions',raw)
                    historical=_semver_read(raw);select_semver_inventory('oxml-opc',common['shared_version'],historical)
                matches=[r for r in historical['versions'] if r['num']==version]
                if len(matches)!=1:raise ValueError('required historical canary archive absent')
                checksum=matches[0]['checksum'];canary_root=baseline(name,version,checksum)
            facts,key=_semver_row_identity('canary',name,version,subject,'default',common,{'canary':canary,'source_sha256':hashlib.sha256(SEMVER_CANARY_SOURCES[canary]).hexdigest(),'archive_sha256':checksum})
            row_directory=payload/'rows'/key;row_directory.mkdir(parents=True)
            row={'key':key,'kind':'canary','input':facts,'status':'incomplete'}
            try:
                if key in reused:
                    row=_semver_copy_reused_row(reused[key],row_directory,facts)
                else:
                    row.update(status='success',result=_semver_canary_row(canary,subject,canary_root,common,row_directory))
            except (ValueError,OSError,KeyError,TypeError,subprocess.SubprocessError) as error:
                row['error']=type(error).__name__+': '+str(error)
            _semver_write(row_directory/'row.json',row)
            receipt['rows'][key]={'path':'rows/'+key+'/row.json','input':facts,'status':row['status']}
        if _semver_source()!=source or _semver_environment()[1]!=environment:
            raise ValueError('source/config/environment changed during SemVer acquisition')
        receipt['complete']=all(row['status']=='success' for row in receipt['rows'].values())
        if not receipt['complete']:
            receipt['errors'].append('required API/compiler row failed or incomplete')
    except (ValueError,OSError,KeyError,TypeError,subprocess.SubprocessError) as error:
        receipt['errors'].append(type(error).__name__+': '+str(error))
    finally:
        packet=_semver_freeze_packet(output,receipt)
        if external is not None and not keep_external:
            shutil.rmtree(external)
        for parent in {str(ref['packet']['owned_parent']) for ref in reused.values()}:
            shutil.rmtree(parent)
    print(json.dumps({'complete':receipt['complete'],'packet':str(packet),'sha256':_semver_hash(packet),'errors':receipt['errors']}))
    return 0 if receipt['complete'] else 1



SEMVER_NATIVE_JOBS = {
    'x86_64-unknown-linux-gnu': {'key':'semver-linux','name':'SemVer native (linux-x86_64)','runner':'ubuntu-24.04'},
    'aarch64-apple-darwin': {'key':'semver-macos','name':'SemVer native (macos-arm64)','runner':'macos-26'},
    'x86_64-pc-windows-msvc': {'key':'semver-windows','name':'SemVer native (windows-x86_64)','runner':'windows-2025'},
}


def _semver_git_blobs(files):
    return {name:hashlib.sha1(b'blob '+str((REPO/name).stat().st_size).encode()+b'\0'+(REPO/name).read_bytes()).hexdigest() for name in files}


def _semver_gh():
    target = _semver_host()
    executable = shutil.which('gh')
    if executable is None:
        raise ValueError('independently pinned local gh is required')
    return _semver_tool_identity(executable,target,SEMVER_GH_VERIFY_TOOLS[target])


def _semver_api(endpoint, directory, sequence, *, paginated=False):
    if not endpoint.startswith('/repos/tensorbee/rdocx') or any(c in endpoint for c in ('\n','\r','\\')):
        raise ValueError('unexpected SemVer control-plane endpoint')
    gh = _semver_gh()
    name = 'api-'+str(sequence)
    argv = [gh['executable'],'api','--hostname','github.com','--method','GET']
    if paginated:
        argv.extend(['--paginate','--slurp'])
    argv.append(endpoint)
    process = _semver_run(argv,REPO,dict(os.environ),directory,name,300)
    _semver_success(process)
    data = _semver_read(Path(directory)/(name+'.stdout'))
    if paginated:
        if not isinstance(data,list) or not data:
            raise ValueError('complete paginated API response is absent')
        pages = data
    else:
        pages = [data]
    return pages, {'endpoint':endpoint,'process_file':name+'.process.json','stdout_file':name+'.stdout',
                   'stdout_sha256':process['stdout_sha256'],'pages':len(pages),'http_success':'actual pinned gh api exit0'}


def _semver_api_union(pages, field):
    rows=[];totals=set()
    for page in pages:
        if not isinstance(page,dict) or type(page.get('total_count')) is not int or not isinstance(page.get(field),list):
            raise ValueError('malformed paginated SemVer authority response')
        totals.add(page['total_count']);rows.extend(page[field])
    if len(totals)!=1 or next(iter(totals))!=len(rows) or any(type(r.get('id')) is not int for r in rows) or len({r['id'] for r in rows})!=len(rows):
        raise ValueError('incomplete/duplicated/changing SemVer authority pagination')
    return rows


def _semver_initialize_authority_state(sprint):
    if state_path(sprint).exists():
        return load(sprint)
    cmd_init(argparse.Namespace(sprint=sprint,force=False,resume=False,max_review_passes=MAX_REVIEW_PASSES,max_workers=None,reviewer=None))
    return load(sprint)


def _collect_release_semver_authority(sprint, scope):
    """Actual deciding-caller collection. No evidence packet supplies this context."""
    resolved,_path=_semver_authority_location(())
    if sprint!=resolved:
        raise ValueError('authority sprint differs from canonical CURRENT_SPRINT')
    data=_semver_initialize_authority_state(sprint)
    context={'status':'incomplete','scope':scope,'collector_started':datetime.now(timezone.utc).isoformat().replace('+00:00','Z')}
    data['semver_authority']=context;save(sprint,data)
    owned=None
    try:
        source=_semver_source();policy=_semver_policy();blobs=_semver_git_blobs(source['files'])
        parent=SCRATCH/'semver-authority';parent.mkdir(parents=True,exist_ok=True)
        owned=Path(tempfile.mkdtemp(prefix='catalog-',dir=parent)).resolve();raw=owned/'raw';raw.mkdir()
        source_roles=('scripts/sprint_workflow.py','.github/workflows/ci.yml','.github/workflows/wheels.yml','rust-toolchain.toml')
        context.update(source_identity=source['sha256'],policy_identity=policy,catalog_path=str(owned/'catalog.json'),head=git_head())
        observations=[];sequence=0
        def api(endpoint,paginated=False):
            nonlocal sequence
            sequence+=1
            pages,observation=_semver_api(endpoint,raw,sequence,paginated=paginated)
            observations.append(observation)
            return pages
        repository=api('/repos/tensorbee/rdocx')[0]
        if repository.get('id')!=1164240333 or repository.get('owner',{}).get('id')!=263075749 or repository.get('full_name')!='tensorbee/rdocx' or repository.get('private') is not False:
            raise ValueError('independent current repository identity differs')
        workflow=api('/repos/tensorbee/rdocx/actions/workflows/ci.yml')[0]
        if type(workflow.get('id')) is not int or workflow.get('path')!='.github/workflows/ci.yml':
            raise ValueError('actual CI workflow identity differs')
        endpoint=f"/repos/tensorbee/rdocx/actions/workflows/{workflow['id']}/runs?per_page=100"
        runs=_semver_api_union(api(endpoint,True),'workflow_runs')
        current_run=os.environ.get('GITHUB_RUN_ID');current_attempt=os.environ.get('GITHUB_RUN_ATTEMPT')
        job_key=os.environ.get('GITHUB_JOB','')
        purpose='sign' if job_key=='semver-sign' else 'current-gate' if job_key=='ci-gate' else 'producer' if job_key in {v['key'] for v in SEMVER_NATIVE_JOBS.values()} else 'external'
        context['purpose']=purpose
        attempts=[];trees={};admitted=[];excluded=[]
        for run in runs:
            if run.get('event') not in ('push','workflow_dispatch') or run.get('event')=='push' and run.get('head_branch')!='main':
                continue
            if run.get('workflow_id')!=workflow['id'] or not isinstance(run.get('repository'),dict) or not isinstance(run.get('head_repository'),dict) or run['repository'].get('id')!=1164240333 or run['head_repository'].get('id')!=1164240333:
                continue
            head=run.get('head_sha')
            if not isinstance(head,str) or re.fullmatch('[0-9a-f]{40}',head) is None:
                raise ValueError('invalid actual authority source commit')
            if head not in trees:
                tree=api(f'/repos/tensorbee/rdocx/git/trees/{head}?recursive=1')[0]
                if tree.get('truncated') is not False or not isinstance(tree.get('tree'),list):
                    raise ValueError('source-provider tree is incomplete')
                trees[head]={entry['path']:entry['sha'] for entry in tree['tree'] if entry.get('type')=='blob'}
            tree=trees[head]
            if any(tree.get(name)!=blobs[name] for name in source_roles):
                continue
            if type(run.get('run_attempt')) is not int or run['run_attempt']<1:
                raise ValueError('missing actual run retry count')
            artifact_rows=_semver_api_union(api(f"/repos/tensorbee/rdocx/actions/runs/{run['id']}/artifacts?per_page=100",True),'artifacts')
            admitted.append({'id':run['id'],'run_attempt':run['run_attempt'],'head_sha':head,'event':run['event'],'head_branch':run['head_branch']})
            for attempt_number in range(1,run['run_attempt']+1):
                attempt=api(f"/repos/tensorbee/rdocx/actions/runs/{run['id']}/attempts/{attempt_number}")[0]
                if attempt.get('id')!=run['id'] or attempt.get('run_attempt')!=attempt_number or attempt.get('head_sha')!=head:
                    raise ValueError('attempt-qualified run identity changed')
                jobs=_semver_api_union(api(f"/repos/tensorbee/rdocx/actions/runs/{run['id']}/attempts/{attempt_number}/jobs?per_page=100",True),'jobs')
                owner_jobs=[j for j in jobs if j.get('name')=='Owner unchanged compatibility corpus']
                if len(owner_jobs)>1:
                    raise ValueError('ambiguous source-qualified owner-mode job')
                if run['event']=='workflow_dispatch' and len(owner_jobs)==1 and owner_jobs[0].get('started_at') and owner_jobs[0].get('conclusion')!='skipped':
                    # This reviewed job can only start after the actual owner-only conditional.
                    excluded.append({'run_id':run['id'],'run_attempt':attempt_number,'reason':'actual source-qualified owner-only job started','job_id':owner_jobs[0]['id']})
                    continue
                same_current=str(run['id'])==current_run and str(attempt_number)==current_attempt
                for target,job_policy in SEMVER_NATIVE_JOBS.items():
                    native=[j for j in jobs if j.get('name')==job_policy['name']]
                    sign=[j for j in jobs if j.get('name')=='SemVer sign ('+target+')']
                    gate=[j for j in jobs if j.get('name')=='CI gate']
                    required=native+(sign if not same_current or purpose not in ('sign','producer') else [])+(gate if not same_current or purpose=='external' else [])
                    expected_count=1+(0 if same_current and purpose in ('sign','producer') else 1)+(0 if same_current and purpose!='external' else 1)
                    status='success'
                    if len(native)!=1 or len(sign)>1 or len(gate)>1 or len(required)!=expected_count:
                        status='incomplete'
                    elif any(j.get('status')!='completed' for j in required):
                        status='in_progress'
                    elif any(j.get('conclusion')=='skipped' for j in required):
                        status='required_skipped'
                    elif any(j.get('conclusion')!='success' for j in required):
                        status='failed'
                    elif (not same_current or purpose=='external') and (attempt.get('status')!='completed' or attempt.get('conclusion')!='success'):
                        status='failed' if attempt.get('status')=='completed' else 'in_progress'
                    if native and ('self-hosted' in native[0].get('labels',[]) or job_policy['runner'] not in native[0].get('labels',[])):
                        status='failed'
                    native_name=f"semver-native-{target}-{run['id']}-{attempt_number}"
                    native_artifacts=[a for a in artifact_rows if a.get('name')==native_name]
                    if len(native_artifacts)>1:
                        raise ValueError('ambiguous actual native producer artifact')
                    proof_name=f"semver-proof-{target}-{run['id']}-{attempt_number}"
                    proof=[a for a in artifact_rows if a.get('name')==proof_name]
                    if len(proof)>1:
                        raise ValueError('ambiguous attempt-qualified SemVer artifact')
                    if status=='success' and not (same_current and purpose in ('sign','producer')) and (len(proof)!=1 or proof[0].get('expired') is not False):
                        status='incomplete'
                    started=min((j['started_at'] for j in required if j.get('started_at')),default=attempt.get('run_started_at'))
                    completed=max((j['completed_at'] for j in required if j.get('completed_at')),default=None) if status!='in_progress' else None
                    attempts.append({'id':f"{run['id']}:{attempt_number}:{target}",'repository':'tensorbee/rdocx','workflow_id':str(workflow['id']),
                        'run_id':str(run['id']),'run_attempt':attempt_number,'target':target,'status':status,'started':started,'completed':completed,
                        'predecessors':[],'source_head':head,'source_ref':'refs/heads/'+run['head_branch'],'event':run['event'],
                        'source_blobs':tree,'current_source_equivalent':all(tree.get(name)==digest for name,digest in blobs.items()),
                        'native_job':native[0] if len(native)==1 else None,'sign_job':sign[0] if len(sign)==1 else None,
                        'ci_gate':gate[0] if len(gate)==1 else None,'artifact':proof[0] if len(proof)==1 else None,
                        'native_artifact':native_artifacts[0] if len(native_artifacts)==1 else None,
                        'active_current_producer':same_current and purpose=='producer'})
        # Detect acquisition drift instead of blessing a partial pagination window.
        again=_semver_api_union(api(endpoint,True),'workflow_runs')
        before={r['id']:(r.get('run_attempt'),r.get('updated_at'),r.get('status'),r.get('conclusion')) for r in runs}
        after={r['id']:(r.get('run_attempt'),r.get('updated_at'),r.get('status'),r.get('conclusion')) for r in again}
        if before!=after:
            raise ValueError('workflow inventory changed during authority acquisition; new attempt required')
        context.update(status='complete',cutoff=datetime.now(timezone.utc).isoformat().replace('+00:00','Z'))
        catalog={'schema':1,'repository_id':1164240333,'owner_id':263075749,'workflow_id':workflow['id'],
                 'context':{key:context[key] for key in ('source_identity','policy_identity','scope','cutoff')},
                 'observations':observations,'admitted_runs':admitted,'excluded_owner_attempts':excluded,'attempts':attempts,'raw_inventory':_semver_inventory_files(raw)}
        _semver_write(owned/'catalog.json',catalog)
        context['catalog_sha256']=_semver_hash(owned/'catalog.json')
        data=load(sprint);data['semver_authority']=context;save(sprint,data)
        return context,catalog
    except (ValueError,OSError,KeyError,TypeError,subprocess.SubprocessError) as error:
        context.update(status='incomplete',error=type(error).__name__+': '+str(error))
        try:
            data=load(sprint);data['semver_authority']=context;save(sprint,data)
            if owned is not None:_semver_write(owned/'failure.json',context)
        except (OSError,ValueError,KeyError,TypeError) as retention_error:
            error.add_note('failure observation retention also failed: '+str(retention_error))
        raise


def _semver_attempt_for_row(catalog, row_input, *, for_reuse=False):
    candidates=[]
    key=_semver_identity(row_input)
    for attempt in catalog['attempts']:
        if attempt['target']!=row_input['target']:
            continue
        if 'current_source' in row_input and not attempt['current_source_equivalent']:
            continue
        if for_reuse and attempt.get('active_current_producer'):
            continue
        candidate={**attempt,'row_key':key,'input_sha256':key}
        candidates.append(candidate)
    return select_semver_attempt(candidates)


def _semver_safe_zip(archive, destination, expected_names):
    destination=Path(destination);destination.mkdir(mode=0o700)
    with zipfile.ZipFile(archive) as bundle:
        members=bundle.infolist()
        if len(members)!=len(expected_names) or {m.filename for m in members}!=set(expected_names):
            raise ValueError('unexpected/missing/duplicate fixed SemVer artifact payload')
        for member in members:
            if member.is_dir() or stat.S_ISLNK(member.external_attr>>16) or '/' in member.filename or '\\' in member.filename:
                raise ValueError('unsafe SemVer diagnostic ZIP member')
            target=destination/member.filename
            with bundle.open(member) as stream,target.open('xb') as output:
                shutil.copyfileobj(stream,output)
            if target.stat().st_size!=member.file_size:
                raise ValueError('SemVer diagnostic ZIP size differs')
    return destination


def _semver_download_artifact(artifact, destination):
    if type(artifact.get('id')) is not int or artifact.get('expired') is not False or not isinstance(artifact.get('digest'),str) or re.fullmatch('sha256:[0-9a-f]{64}',artifact['digest']) is None:
        raise ValueError('artifact lacks actual ID/wrapper digest or is expired')
    destination=Path(destination);destination.mkdir(parents=True)
    gh=_semver_gh()
    row=_semver_run([gh['executable'],'api','--hostname','github.com',f"/repos/tensorbee/rdocx/actions/artifacts/{artifact['id']}/zip"],REPO,dict(os.environ),destination,'artifact-download',300)
    _semver_success(row)
    archive=destination/'artifact-download.stdout'
    if _semver_hash(archive)!=artifact['digest'].removeprefix('sha256:'):
        raise ValueError('actual artifact wrapper digest differs from API authority')
    return archive



def _semver_unpack_packet(packet, destination):
    destination=Path(destination);destination.mkdir(mode=0o700)
    seen=set()
    with tarfile.open(packet,'r:') as bundle:
        for member in bundle:
            name=member.name;parts=name.split('/')
            if member.type not in (tarfile.REGTYPE,tarfile.AREGTYPE) or member.issparse() or any('sparse' in k.lower() for k in member.pax_headers) or any(c in name for c in ('\\',':','\0')) or any(p in ('','.','..') for p in parts) or name in seen or not (name=='receipt.json' or len(parts)>1 and parts[0]=='payload'):
                raise ValueError('unsafe/duplicate/unexpected signed packet member')
            seen.add(name);path=destination.joinpath(*parts)
            if not path.resolve().is_relative_to(destination.resolve()):
                raise ValueError('signed packet path escaped owned destination')
            path.parent.mkdir(parents=True,exist_ok=True)
            count=0
            with bundle.extractfile(member) as stream,path.open('xb') as output:
                while block:=stream.read(1024*1024):
                    count+=len(block)
                    if count>member.size:raise ValueError('signed packet member exceeds declared size')
                    output.write(block)
            if count!=member.size:raise ValueError('truncated signed packet member')
    if 'receipt.json' not in seen:
        raise ValueError('signed packet receipt absent')
    receipt=_semver_read(destination/'receipt.json')
    authenticate_semver_files(destination/'payload',receipt['payload_inventory'])
    return receipt


def _semver_bound_file(root, reference):
    if not isinstance(reference,dict) or not isinstance(reference.get('path'),str):
        raise ValueError('missing retained file reference')
    relative=reference['path']
    parts=relative.split('/')
    if any(p in ('','.','..') for p in parts) or any(c in relative for c in ('\\',':','\0')):
        raise ValueError('noncanonical retained role path')
    path=Path(root)
    for part in parts:
        path=path/part
        if path.is_symlink():raise ValueError('symlink in retained role path')
    if path.is_symlink() or not path.is_file() or not path.resolve().is_relative_to(Path(root).resolve()) or _semver_hash(path)!=reference['sha256'] or path.stat().st_size!=reference['bytes']:
        raise ValueError('retained role-file hash/size/path differs')
    return path


def _semver_validate_archive(payload, crate, version, checksum):
    folder=Path(payload)/'archives'/(crate+'-'+version)
    archive=folder/'published.crate';receipt=_semver_read(folder/'receipt.json')
    if _semver_hash(archive)!=checksum or receipt.get('status')!='READY' or receipt.get('expected_sha256')!=checksum or receipt.get('archive_sha256')!=checksum or receipt.get('checksum_authenticated') is not True or receipt.get('response_complete') is not True or receipt.get('crate')!=crate or receipt.get('version')!=version:
        raise ValueError('published baseline archive/provenance identity differs')
    if receipt['received_bytes']>SEMVER_ARCHIVE_COMPRESSED_LIMIT or receipt['decompressed_tar_bytes']>SEMVER_ARCHIVE_TAR_LIMIT or len(receipt['members'])>SEMVER_ARCHIVE_MEMBER_LIMIT:
        raise ValueError('baseline acquisition exceeded approved private budget')
    expected={m['path']:{'sha256':m['sha256'],'bytes':m['size']} for m in receipt['members'] if m['type']=='file'}
    authenticate_semver_files(folder/'source',expected)
    manifest=tomllib.loads((folder/'source'/(crate+'-'+version)/'Cargo.toml').read_text())
    if manifest['package']['name']!=crate or manifest['package']['version']!=version:
        raise ValueError('published manifest identity differs')
    check_no_semver_overrides(manifest['package'],manifest.get('workspace',{}),SEMVER_QUERY_POLICY,role='baseline')
    return receipt


def _semver_validate_side(directory, side, crate, version, target, view, original_cwd):
    manifest_path=_semver_bound_file(directory,side['manifest'])
    if side['generation_before']==side['generation_after'] or side['generation_after']['sha256']!=_semver_hash(manifest_path):
        raise ValueError('selected generated project was not actually rewritten by this invocation')
    manifest=tomllib.loads(manifest_path.read_text())
    lock=tomllib.loads(_semver_bound_file(directory,side['lock']).read_text())
    _semver_bound_file(directory,side['lib'])
    document=_semver_read(_semver_bound_file(directory,side['json']))
    for name in ('manifest','lock','lib','json'):
        expected=(_semver_path(side['project'],target)/{'manifest':'Cargo.toml','lock':'Cargo.lock','lib':'lib.rs'}[name]) if name!='json' else _semver_path(side['json_path'],target)
        if _semver_path(side[name]['producer_path'],target)!=expected:
            raise ValueError('retained generated file producer path differs from exact role join')
    metadata_path=_semver_bound_file(directory,side['metadata'])
    if _semver_hash(metadata_path)!=side['metadata']['sha256']:
        raise ValueError('generated metadata observation hash differs')
    metadata=_semver_read(metadata_path)
    source=_semver_path(side['source_root'],target);project=_semver_path(side['project'],target)
    dependency=manifest.get('dependencies',{}).get(crate)
    if manifest.get('package',{}).get('name')!='placeholder' or 'workspace' not in manifest or not isinstance(dependency,dict) or _semver_path(dependency['path'],target)!=source or dependency.get('default-features',True) is not True:
        raise ValueError('generated manifest does not bind its actual source role')
    if side['configuration_cwd']!=original_cwd:
        raise ValueError('auxiliary metadata changed original tool Cargo cwd semantics')
    observed_process=_semver_read(_semver_bound_file(directory,side['metadata_process']))
    _semver_success(observed_process)
    metadata_argv=observed_process['argv']
    if observed_process['cwd']!=original_cwd or observed_process['stdout_sha256']!=side['metadata']['sha256'] or '--locked' not in metadata_argv or '--offline' not in metadata_argv or '--manifest-path' not in metadata_argv or _semver_path(metadata_argv[metadata_argv.index('--manifest-path')+1],target)!=_semver_path(side['project'],target)/'Cargo.toml':
        raise ValueError('metadata graph was not frozen under original tool context')
    target_root=_semver_path(metadata['target_directory'],target)
    if target_root!=_semver_path(side['target'],target) or not target_root.is_relative_to(project):
        raise ValueError('generated target role is shared/outside project')
    packages=[p for p in metadata['packages'] if p['id']==side['subject_package_id']]
    if len(packages)!=1:
        raise ValueError('generated subject package is ambiguous')
    package=packages[0]
    if package['name']!=crate or package['version']!=version or package['source'] is not None or _semver_path(package['manifest_path'],target)!=source/'Cargo.toml':
        raise ValueError('generated subject source/version differs')
    libraries=[t for t in package['targets'] if 'lib' in t['kind']]
    if len(libraries)!=1 or libraries[0]['name']!=side['library']:
        raise ValueError('actual library target differs from retained JSON role')
    node=next(n for n in metadata['resolve']['nodes'] if n['id']==package['id'])
    requested=dependency.get('features',[])
    if node['features']!=side['effective_features'] or requested!=side['requested_features'] or not set(requested).issubset(node['features']) or sorted(package['features'])!=side['available_features']:
        raise ValueError('effective/generated feature sides differ')
    if view=='default' and requested or view=='all' and set(requested)!=set(package['features']):
        raise ValueError('explicit default/all feature request was silently reduced')
    expected_json=target_root/target/'doc'/(side['library'].replace('-','_')+'.json')
    if _semver_path(side['json_path'],target)!=expected_json or document['crate_version']!=version or document['index'][str(document['root'])]['name']!=side['library'] or document['format_version']!=side['json_format']:
        raise ValueError('actual rustdoc library/version/target binding differs')
    subjects=[p for p in lock['package'] if p['name']==crate]
    if len(subjects)!=1 or subjects[0]['version']!=version or 'source' in subjects[0]:
        raise ValueError('generated lock does not bind unique local subject')
    return side


def _semver_validate_row(row, directory, receipt):
    facts=row['input'];key=_semver_identity(facts)
    if row['key']!=key or row['status']!='success' or facts['target']!=receipt['target'] or facts['policy']!=receipt['policy']:
        raise ValueError('SemVer row identity/status/policy differs')
    if facts.get('current_source') is not None and facts['current_source']!=receipt['source']['sha256']:
        raise ValueError('SemVer row current source identity differs')
    target=receipt['target'];result=row['result'];process=result['process']
    origin=row.get('origin_context',receipt)
    if origin['target']!=target or origin['policy']!=receipt['policy'] or origin['source']['sha256']!=receipt['source']['sha256'] and 'current_source' in facts:
        raise ValueError('reused original context differs from applicable row inputs')
    expected_compiler={name:{'sha256':fact['sha256'],'version':fact['version']} for name,fact in origin['rust'].items()}
    expected_environment={k:v for k,v in origin['environment'].items() if k!='configs'}
    expected_environment['present_configs']={path:fact for path,fact in origin['environment']['configs'].items() if fact['present']}
    if facts['compiler']!=expected_compiler or facts['environment']!=expected_environment:
        raise ValueError('row compiler/configuration facts differ from original observation')
    process_file='check-release' if row['kind']=='api' else 'cargo-check'
    recorded=_semver_read(Path(directory)/(process_file+'.process.json'))
    if recorded!=process or _semver_hash(Path(directory)/(process_file+'.stdout'))!=process['stdout_sha256'] or _semver_hash(Path(directory)/(process_file+'.stderr'))!=process['stderr_sha256']:
        raise ValueError('raw row process/output bytes differ')
    if process['launch_error'] is not None:
        raise ValueError('row process did not launch completely')
    if row['kind']=='api':
        crate,version,baseline,view=facts['crate'],facts['version'],facts['baseline'],facts['view']
        protocol=parse_semver_query_protocol(crate,baseline,version,SEMVER_QUERY_POLICY,
            (Path(directory)/'check-release.stderr').read_text(),(Path(directory)/'check-release.stdout').read_text(),process['exit'],timed_out=process['timed_out'])
        if not protocol['clean_query_results'] or protocol!=result['protocol']:
            raise ValueError('API row is failed/warned/all-skipped or protocol changed')
        sides=result['sides']
        expected_roles={role+'-'+kind:side['source_root' if kind=='source' else 'project'] for role,side in sides.items() for kind in ('source','generated')}
        actual_roles=result.get('path_roles')
        if not isinstance(actual_roles,list) or len(actual_roles)!=4 or len({r.get('role') for r in actual_roles})!=4 or {r.get('role'):r.get('producer_prefix') for r in actual_roles}!=expected_roles:
            raise ValueError('actual supplied source/generated role map differs')
        work=_semver_path(result['baseline_work_root'],target)
        checkout=_semver_path(process['cwd'],target)
        baseline_source=_semver_path(sides['baseline']['source_root'],target)
        if baseline_source!=work/'baseline' or work.is_relative_to(checkout) or checkout.is_relative_to(work):
            raise ValueError('baseline work/source aliases current Cargo workspace')
        checks={str(parent/'Cargo.toml'):False for parent in (work,*work.parents)}
        if result['baseline_ancestor_cargo']!=checks:
            raise ValueError('baseline external Cargo ancestry was not qualified')
        for role,side in sides.items():
            for field in ('manifest','lock','lib','json'):
                if _semver_resolve_role(side[field]['producer_path'],actual_roles,target)!=role+'-generated':
                    raise ValueError('retained file selects another source/generated role')
            _semver_validate_side(directory,side,crate,version if role=='current' else baseline,target,view,process['cwd'])
        if set(sides)!={'current','baseline'} or any(_semver_path(sides['current'][name],target)==_semver_path(sides['baseline'][name],target) for name in ('project','target','json_path')):
            raise ValueError('both-side generated artifacts are missing/aliased')
        argv=[origin['tool']['executable'],'check-release','--manifest-path',f'crates/{crate}/Cargo.toml','--package',crate,
              '--baseline-root',sides['baseline']['source_root'],'--target',target,'--default-features' if view=='default' else '--all-features','--color','never','--verbose']
        if process['argv']!=argv or result['baseline_lock']['changed'] is not False:
            raise ValueError('API argv or original/generated lock disposition differs')
        current_source=_semver_path(sides['current']['source_root'],target)
        if current_source!=_semver_path(process['cwd'],target)/'crates'/crate:
            raise ValueError('current source path is not its actual checkout role')
    elif row['kind']=='canary':
        canary=facts['extra']['canary']
        if result['client']['subject']!=facts['baseline']:
            raise ValueError('compiler canary subject role differs')
        retained={}
        for reference in [*result['files'].values(),*result['artifacts']]:
            path=_semver_bound_file(directory,reference)
            original=_semver_path(reference['producer_path'],target)
            if original in retained:
                raise ValueError('ambiguous retained canary source/artifact role')
            retained[original]=path
        diagnostics=release_semver_canary_diagnostics(canary,result['client'],(Path(directory)/'cargo-check.stdout').read_bytes(),process,retained=retained,producer_target=target)
        if diagnostics!=result['diagnostics']:
            raise ValueError('portable canary diagnostics differ')
        if result['metadata']['path']!=result['metadata_file']:
            raise ValueError('canary metadata role path differs')
        metadata=_semver_read(_semver_bound_file(directory,result['metadata']))
        observed_metadata=_semver_read(_semver_bound_file(directory,result['metadata_process']))
        _semver_success(observed_metadata)
        expected_metadata_argv=[origin['rust']['cargo']['path'],'metadata','--locked','--format-version','1','--manifest-path',str(_semver_path(result['client_root'],target)/'Cargo.toml'),'--offline']
        if observed_metadata['argv']!=expected_metadata_argv or observed_metadata['cwd']!=result['client_root'] or observed_metadata['stdout_sha256']!=result['metadata']['sha256']:
            raise ValueError('canary graph was not frozen under exact client command/context')
        fetch=result['fetch']
        fetch_process=_semver_read(_semver_bound_file(directory,fetch['process']))
        _semver_success(fetch_process)
        expected_fetch_argv=[origin['rust']['cargo']['path'],'fetch','--locked','--manifest-path',str(_semver_path(result['client_root'],target)/'Cargo.toml')]
        if fetch_process['argv']!=expected_fetch_argv or fetch_process['cwd']!=result['client_root'] or _semver_hash(_semver_bound_file(directory,fetch['stdout']))!=fetch_process['stdout_sha256'] or _semver_hash(_semver_bound_file(directory,fetch['stderr']))!=fetch_process['stderr_sha256']:
            raise ValueError('canary all-target cache acquisition command/context/raw output differs')
        expected_client_lock=_semver_path(result['client_root'],target)/'Cargo.lock'
        if _semver_path(fetch['lock_before']['producer_path'],target)!=expected_client_lock or _semver_path(result['files']['lock']['producer_path'],target)!=expected_client_lock or _semver_hash(_semver_bound_file(directory,fetch['lock_before']))!=result['files']['lock']['sha256']:
            raise ValueError('canary cache acquisition changed original check lock or its source role')
        source=_semver_path(result['source_root'],target)
        opc=[p for p in metadata['packages'] if p['name']=='oxml-opc']
        if len(opc)!=1 or opc[0]['source'] is not None:
            raise ValueError('compiler graph has duplicate/registry nominal OPC identity')
        if canary!='word-nominal' and (opc[0]['version']!=facts['version'] or _semver_path(opc[0]['manifest_path'],target)!=source/'Cargo.toml'):
            raise ValueError('actual compiler dependency source differs')
        if canary=='word-nominal':
            word=[p for p in metadata['packages'] if p['name']=='rdocx']
            if len(word)!=1 or word[0]['source'] is not None or _semver_path(word[0]['manifest_path'],target)!=source/'Cargo.toml':
                raise ValueError('Word nominal client does not use current Word source')
            if _semver_path(opc[0]['manifest_path'],target)!=source.parent/'oxml-opc'/'Cargo.toml':
                raise ValueError('Word nominal OPC source is not the same candidate closure')
        cargo_argv=[origin['rust']['cargo']['path'],'check','--manifest-path',str(_semver_path(result['client_root'],target)/'Cargo.toml'),'--target',target,'--message-format=json']
        if process['argv']!=cargo_argv:
            raise ValueError('actual external compiler command differs')
    else:
        raise ValueError('unqualified SemVer row kind')
    return row


def _semver_validate_producer(receipt, directory, attempt):
    if receipt.get('schema')!=1 or receipt.get('complete') is not True or receipt.get('policy')!=_semver_policy() or receipt.get('target')!=attempt['target']:
        raise ValueError('native packet is incomplete/wrong policy/target')
    source=receipt['source']
    if _semver_identity(source['files'])!=source['sha256'] or set(source['files'])!=set(source['git_blobs']) or any(attempt['source_blobs'].get(name)!=digest for name,digest in source['git_blobs'].items()):
        raise ValueError('actual source-provider tree differs from signed measured source')
    producer=receipt['producer']
    if producer['source_head']!=attempt['source_head'] or producer['run_id']!=attempt['run_id'] or int(producer['run_attempt'])!=attempt['run_attempt'] or producer['repository']!='tensorbee/rdocx' or producer['ref']!=attempt['source_ref'] or producer['event']!=attempt['event']:
        raise ValueError('signed native producer run/source/event association differs')
    target=receipt['target'];pin=SEMVER_NATIVE_TOOLS[target]
    if receipt['tool']['executable_sha256']!=pin['executable_sha256'] or receipt['tool']['archive_sha256']!=pin['archive_sha256']:
        raise ValueError('native API-tool integrity differs from policy')
    setup=Path(directory)/'payload/setup'
    for name in ('rustc','cargo','rustdoc'):
        version=(setup/(name+'.stdout')).read_text()
        if version!=receipt['rust'][name]['version'] or not version.startswith(name+' 1.97.1'):
            raise ValueError('actual native compiler version observation differs')
    if 'host: '+target not in receipt['rust']['rustc']['version'] or not (setup/'tool.stdout').read_text().startswith('cargo-semver-checks 0.51.0'):
        raise ValueError('genuine native host/tool identity is absent')
    observed_members=select_semver_members(receipt['scope']['scope'],receipt['scope']['families'],RELEASE_CRATES,_semver_read(Path(directory)/'payload/workspace.stdout')['packages'])
    if receipt['members']!=observed_members:
        raise ValueError('actual library/CLI membership inventory differs')
    for key,fact in receipt['rows'].items():
        if re.fullmatch('[0-9a-f]{64}',key) is None or fact['path']!='rows/'+key+'/row.json':
            raise ValueError('noncanonical retained row membership path')
        row_path=Path(directory)/'payload'/fact['path'];row=_semver_read(row_path)
        if row['key']!=key or row['input']!=fact['input'] or row['status']!=fact['status']:
            raise ValueError('row manifest membership differs')
        _semver_validate_row(row,row_path.parent,receipt)
    _semver_required_rows(receipt,receipt['scope'],Path(directory)/'payload')
    return receipt


def _semver_verify_packet(directory, attempt, *, unsigned=False):
    directory=Path(directory);target=attempt['target'];filename='semver-'+target+'.tar';packet=directory/filename
    if packet.is_symlink() or not packet.is_file():
        raise ValueError('fixed expected native packet absent')
    if not unsigned:
        bundle=directory/'bundle.json'
        archive=directory/'artifact.zip'
        expected_artifact=attempt['artifact']
        if _semver_hash(archive)!=expected_artifact['digest'].removeprefix('sha256:'):
            raise ValueError('proof wrapper bytes differ from independently selected artifact')
        with zipfile.ZipFile(archive) as zipped:
            if len(zipped.infolist())!=2 or set(zipped.namelist())!={filename,'bundle.json'}:
                raise ValueError('signed proof wrapper has unexpected subjects')
            if hashlib.sha256(zipped.read(filename)).hexdigest()!=_semver_hash(packet) or hashlib.sha256(zipped.read('bundle.json')).hexdigest()!=_semver_hash(bundle):
                raise ValueError('signed packet/bundle differs from actual wrapper')
        gh=_semver_gh();work=Path(tempfile.mkdtemp(prefix='rdocx-semver-verify-'))
        try:
            if hashlib.sha256(SEMVER_TRUSTED_ROOT_JSONL).hexdigest()!=SEMVER_TRUSTED_ROOT_SHA256:
                raise ValueError('embedded independent trust-root bytes changed')
            root=work/'trusted-root.jsonl'
            descriptor=os.open(root,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
            with os.fdopen(descriptor,'wb') as stream:stream.write(SEMVER_TRUSTED_ROOT_JSONL)
            if _semver_hash(root)!=SEMVER_TRUSTED_ROOT_SHA256:raise ValueError('staged independent root changed')
            env=dict(os.environ)
            for name in ('GH_TOKEN','GITHUB_TOKEN','GH_ENTERPRISE_TOKEN','GITHUB_ENTERPRISE_TOKEN','GH_DEBUG'):env.pop(name,None)
            for key,folder in (('GH_CONFIG_DIR','config'),('XDG_CACHE_HOME','cache'),('XDG_STATE_HOME','state'),('XDG_DATA_HOME','data')):
                (work/folder).mkdir();env[key]=str(work/folder)
            expected={'packet_sha256':_semver_hash(packet),'bundle_sha256':_semver_hash(bundle),'packet_subject_name':_semver_signed_subject(attempt),
                'source_digest':attempt['source_head'],'signer_digest':attempt['source_head'],'build_config_digest':attempt['source_head'],
                'source_ref':attempt['source_ref'],'signer_ref':attempt['source_ref'],'source_workflow':'.github/workflows/ci.yml',
                'signer_workflow':'.github/workflows/ci.yml','event_name':attempt['event'],'run_id':int(attempt['run_id']),'run_attempt':attempt['run_attempt']}
            argv=[gh['executable'],'attestation','verify',str(packet.resolve()),'-R','tensorbee/rdocx','--bundle',str(bundle.resolve()),'--custom-trusted-root',str(root),
                '--signer-workflow','tensorbee/rdocx/.github/workflows/ci.yml','--signer-digest',expected['signer_digest'],'--source-digest',expected['source_digest'],
                '--source-ref',expected['source_ref'],'--cert-oidc-issuer','https://token.actions.githubusercontent.com','--deny-self-hosted-runners','--hostname','github.com','--format','json']
            process=_semver_run(argv,REPO,env,work,'verify',120)
            process.update(executable=gh['executable'],executable_sha256=gh['executable_sha256'],trusted_root_sha256=SEMVER_TRUSTED_ROOT_SHA256,
                packet_sha256=expected['packet_sha256'],bundle_sha256=expected['bundle_sha256'],packet_path=str(packet.resolve()),bundle_path=str(bundle.resolve()),trusted_root_path=str(root),
                verifier_implementation='gh',verifier_version='2.97.0',consumer_host=_semver_host())
            policy={'implementation':'gh','version':'2.97.0','consumer_host':_semver_host(),'executable_sha256':SEMVER_GH_VERIFY_TOOLS[_semver_host()]['executable_sha256']}
            qualification=qualify_semver_attestation((work/'verify.stdout').read_bytes(),process,expected,verifier_policy=policy)
            _semver_write(work/'qualification.json',qualification)
        finally:
            consumer=SCRATCH/'semver-consumers';consumer.mkdir(exist_ok=True)
            saved_failure=Path(tempfile.mkdtemp(prefix='attempt-',dir=consumer))
            for path in work.iterdir():
                if path.is_file():shutil.copyfile(path,saved_failure/path.name)
            shutil.rmtree(work)
    unpack=Path(tempfile.mkdtemp(prefix='rdocx-semver-packet-parent-'))
    try:
        receipt=_semver_unpack_packet(packet,unpack/'packet')
        _semver_validate_producer(receipt,unpack/'packet',attempt)
        return {'receipt':receipt,'root':unpack/'packet','owned_parent':unpack,'packet':str(packet),'packet_sha256':_semver_hash(packet),'attempt':attempt}
    except BaseException:
        shutil.rmtree(unpack)
        raise


def _semver_download_proof(attempt, destination):
    if attempt.get('artifact') is None:
        raise ValueError('required signed native proof artifact absent')
    destination=Path(destination);archive=_semver_download_artifact(attempt['artifact'],destination/'download')
    proof=_semver_safe_zip(archive,destination/'proof',{'semver-'+attempt['target']+'.tar','bundle.json'})
    shutil.copyfile(archive,proof/'artifact.zip')
    _semver_write(proof/'proof-selection.json',{'attempt_id':attempt['id']})
    return proof




def _semver_copy_reused_row(reference, destination, expected_input):
    packet=reference['packet'];source=packet['root']/'payload'/reference['row']['path'];row=_semver_read(source)
    if row['input']!=expected_input or row['status']!='success':
        raise ValueError('reused row no longer matches exact applicable inputs')
    for path in source.parent.rglob('*'):
        if path.is_file() and path.name!='row.json':
            _semver_keep(path,destination,path.relative_to(source.parent).as_posix())
    row['origin_context']=row.get('origin_context',{key:packet['receipt'][key] for key in ('tool','rust','source','target','policy','producer','environment')})
    row['reused_from']={'packet_sha256':packet['packet_sha256'],'original_attempt':packet['attempt']['id'],'original_row':reference['row_key']}
    return row





def _semver_existing_ref(root, relative):
    path=Path(root)/relative
    reference={'path':relative,'sha256':_semver_hash(path),'bytes':path.stat().st_size,'producer_path':str(path.resolve())}
    _semver_bound_file(root,reference)
    return reference


def _semver_preserve_failed_generated(roots, destination, target, error):
    retained=[]
    for role,source,target_root in roots:
        for manifest in (Path(target_root)/'semver-checks').glob('*/Cargo.toml'):
            data=tomllib.loads(manifest.read_text())
            dependencies=data.get('dependencies',{})
            if data.get('package',{}).get('name')!='placeholder' or not any(isinstance(d,dict) and d.get('path') and Path(d['path']).resolve()==Path(source).resolve() for d in dependencies.values()):continue
            prefix='failed/'+role+'/'+manifest.parent.name+'/'
            for name in ('Cargo.toml','Cargo.lock','lib.rs'):
                path=manifest.parent/name
                if path.is_file():retained.append(_semver_keep(path,destination,prefix+name))
            for doc in (manifest.parent/'target'/target/'doc').glob('*.json'):
                if doc.is_file():retained.append(_semver_keep(doc,destination,prefix+'doc/'+doc.name))
    _semver_write(Path(destination)/'failed-retention.json',{'qualification':'available diagnostic bytes only, not completed generation or API coverage',
        'error':type(error).__name__+': '+str(error),'files':retained})


def _semver_preserve_failed_client(work, destination, error):
    retained=[]
    for relative in ('Cargo.toml','Cargo.lock','src/lib.rs'):
        path=Path(work)/relative
        if path.is_file():retained.append(_semver_keep(path,destination,'failed/client/'+relative))
    # Preserve actual available client products without representing them as successful compilation.
    for extension in ('*.rmeta','*.rlib'):
        for path in (Path(work)/'target').rglob(extension):
            if path.is_file():retained.append(_semver_keep(path,destination,'failed/client/'+path.relative_to(work).as_posix()))
    _semver_write(Path(destination)/'failed-retention.json',{'qualification':'available client diagnostic bytes only, not compile-pass',
        'error':type(error).__name__+': '+str(error),'files':retained})


def _semver_signed_subject(attempt):
    target=attempt['target'];native=attempt.get('native_job');artifact=attempt.get('native_artifact')
    if target not in SEMVER_NATIVE_JOBS or not isinstance(native,dict) or type(native.get('id')) is not int or native['id']<1 or not isinstance(artifact,dict) or type(artifact.get('id')) is not int or artifact['id']<1:
        raise ValueError('actual native job/artifact identity absent for signed subject')
    digest=artifact.get('digest')
    expected_name=f"semver-native-{target}-{attempt['run_id']}-{attempt['run_attempt']}"
    workflow=artifact.get('workflow_run',{})
    if artifact.get('name')!=expected_name or not isinstance(digest,str) or re.fullmatch('sha256:[0-9a-f]{64}',digest) is None or workflow.get('id')!=int(attempt['run_id']) or workflow.get('head_sha')!=attempt['source_head']:
        raise ValueError('actual native artifact parent/head/digest association differs')
    return f"semver-{target}-job{native['id']}-artifact{artifact['id']}-{digest.removeprefix('sha256:')}.tar"


def _prepare_release_semver_signing(target, producer_outputs, destination):
    """Trusted reviewed signing step, fixed needs output selection, never user subject paths."""
    if os.environ.get('GITHUB_JOB')!='semver-sign' or target not in SEMVER_NATIVE_JOBS:
        raise ValueError('signer is not its actual fixed workflow job')
    scope=_semver_scope('all');sprint,_state=_semver_authority_location(())
    context,catalog=_collect_release_semver_authority(sprint,scope)
    if context['purpose']!='sign':raise ValueError('actual signing authority purpose absent')
    run=os.environ.get('GITHUB_RUN_ID');number=os.environ.get('GITHUB_RUN_ATTEMPT')
    current=[a for a in catalog['attempts'] if a['run_id']==run and str(a['run_attempt'])==number]
    if len(current)!=3 or {a['target'] for a in current}!=set(SEMVER_NATIVE_JOBS) or any(a['status']!='success' for a in current):
        raise ValueError('actual completed native producer union is not available for signing')
    for native_attempt in current:
        _semver_signed_subject(native_attempt)
        if native_attempt['native_artifact'].get('expired') is not False:
            raise ValueError('required native artifact is unavailable before signing')
    attempt=next(a for a in current if a['target']==target)
    artifact=attempt['native_artifact'];subject=_semver_signed_subject(attempt)
    if str(artifact['id'])!=str(producer_outputs['artifact_id']) or artifact['digest'].removeprefix('sha256:')!=producer_outputs['artifact_digest'] or producer_outputs['input_identity']!=context['source_identity']:
        raise ValueError('fixed needs native artifact outputs differ from actual API/control-plane facts')
    destination=Path(destination)
    archive=_semver_download_artifact(artifact,destination/'download')
    expected='semver-'+target+'.tar'
    files=_semver_safe_zip(archive,destination/'subject',{expected})
    packet=files/expected
    if _semver_hash(packet)!=producer_outputs['packet_sha256']:
        raise ValueError('fixed producer packet digest differs')
    observed=_semver_verify_packet(files,attempt,unsigned=True)
    try:
        if observed['receipt']['source']['sha256']!=context['source_identity']:
            raise ValueError('signing subject current input identity differs')
    finally:shutil.rmtree(observed['owned_parent'])
    result={'packet':str(packet.resolve()),'packet_sha256':_semver_hash(packet),'subject_name':subject,
        'native_job_id':attempt['native_job']['id'],'native_artifact_id':artifact['id'],'native_artifact_digest':artifact['digest'],
        'run_id':run,'run_attempt':number,'target':target}
    _semver_write(destination/'signing-selection.json',result)
    return result


def _prepare_release_semver_evidence(scope, destination):
    """Actual deciding caller's live collection/download phase before frozen CLI consumption."""
    scope=_semver_scope(scope) if isinstance(scope,str) else scope
    sprint,_state=_semver_authority_location(())
    context,catalog=_collect_release_semver_authority(sprint,scope)
    destination=Path(destination);destination.mkdir(parents=True)
    proofs=[]
    for target in SEMVER_NATIVE_TOOLS:
        relevant=[a for a in catalog['attempts'] if a['target']==target and a['current_source_equivalent']]
        choice=select_semver_attempt([{**a,'row_key':'native-union-'+target,'input_sha256':context['source_identity']} for a in relevant])
        proofs.append(_semver_download_proof(choice,destination/target))
    return _semver_download_origin_proofs(proofs,catalog,Path(destination)/'origins')


def _prepare_release_semver_reuse(scope, target, destination):
    """Optional acquisition failure invalidates reuse, never mandatory fresh proof."""
    try:
        scope=_semver_scope(scope) if isinstance(scope,str) else scope
        sprint,_state=_semver_authority_location(())
        context,catalog=_collect_release_semver_authority(sprint,scope)
        choices=[a for a in catalog['attempts'] if a['target']==target and a['status']=='success' and not a.get('active_current_producer') and a.get('artifact')]
        if not choices:return []
        chosen=select_semver_attempt([{**a,'row_key':'reuse-'+target,'input_sha256':context['source_identity']} for a in choices])
        proof=_semver_download_proof(chosen,Path(destination)/target)
        return _semver_download_origin_proofs([proof],catalog,Path(destination)/'origins')
    except (ValueError,OSError,KeyError,TypeError,subprocess.SubprocessError) as error:
        refusal={'error':type(error).__name__+': '+str(error),'decision':'reacquire all missing or invalidated rows'}
        try:
            _semver_write(Path(destination)/'reuse-invalidation.json',refusal)
        except OSError as retention_error:
            print('optional SemVer reuse refusal retention failed: '+str(retention_error),file=sys.stderr)
        print('optional SemVer reuse invalidated: '+refusal['error'],file=sys.stderr)
        return []



def _semver_project_state(manifest):
    manifest=Path(manifest)
    if manifest.is_symlink() or not manifest.is_file():
        raise ValueError('generated manifest is not a regular actual file')
    stat_result=manifest.stat()
    return {'sha256':_semver_hash(manifest),'mtime_ns':stat_result.st_mtime_ns,'ctime_ns':stat_result.st_ctime_ns,'bytes':stat_result.st_size}


def _semver_snapshot_projects(target_root):
    return {str(p.resolve()):_semver_project_state(p) for p in (Path(target_root)/'semver-checks').glob('*/Cargo.toml')}


def _semver_resolve_role(path, roles, target):
    path=_semver_path(path,target)
    matches=[]
    for role in roles:
        prefix=_semver_path(role['producer_prefix'],target)
        if path.is_relative_to(prefix):matches.append((len(prefix.parts),role['role']))
    if not matches:raise ValueError('producer path has no authenticated role')
    length=max(m[0] for m in matches);selected=[role for size,role in matches if size==length]
    if len(selected)!=1:raise ValueError('conflicting producer role prefixes')
    return selected[0]



def _semver_verified_directories(directories, catalog):
    packets=[];seen=set()
    try:
        for directory in dict.fromkeys(str(Path(d).resolve()) for d in directories):
            hints=Path(directory)/'proof-selection.json'
            ident=_semver_read(hints).get('attempt_id') if hints.is_file() else None
            candidates=[a for a in catalog['attempts'] if (ident is None or a['id']==ident) and a.get('artifact') and (Path(directory)/('semver-'+a['target']+'.tar')).is_file()]
            matching=[]
            for candidate in candidates:
                try:packet=_semver_verify_packet(directory,candidate)
                except (ValueError,OSError,KeyError,TypeError):continue
                matching.append(packet)
            if len(matching)!=1:
                for packet in matching:shutil.rmtree(packet['owned_parent'])
                raise ValueError('supplied proof lacks a unique actual certificate/job/artifact association')
            packet=matching[0]
            if packet['packet_sha256'] in seen:
                shutil.rmtree(packet['owned_parent']);continue
            seen.add(packet['packet_sha256']);packets.append(packet)
        return packets
    except BaseException:
        for packet in packets:shutil.rmtree(packet['owned_parent'])
        raise


def _semver_check_reused_origin(packet, row_key, by_digest):
    """Follow only the finite actual immutable row-reference chain, never a success label."""
    visited=set();current=packet;key=row_key
    while True:
        identity=(current['packet_sha256'],key)
        if identity in visited:raise ValueError('cyclic immutable row provenance')
        visited.add(identity)
        fact=current['receipt']['rows'].get(key)
        if fact is None:raise ValueError('original reused row missing')
        row=_semver_read(current['root']/'payload'/fact['path'])
        reference=row.get('reused_from')
        if reference is None:return True
        original=by_digest.get(reference['packet_sha256'])
        if original is None or original['attempt']['id']!=reference['original_attempt']:
            raise ValueError('required original authenticated packet/attempt reference is missing')
        original_fact=original['receipt']['rows'].get(reference['original_row'])
        if original_fact is None:raise ValueError('original authenticated row reference missing')
        original_row=_semver_read(original['root']/'payload'/original_fact['path'])
        context=original_row.get('origin_context',{k:original['receipt'][k] for k in ('tool','rust','source','target','policy','producer','environment')})
        if row['input']!=original_row['input'] or row['result']!=original_row['result'] or row.get('origin_context')!=context:
            raise ValueError('reused raw observation/context differs from its authenticated original')
        current,key=original,reference['original_row']


def _semver_execute_reuse(scope, directories):
    sprint,_state=_semver_authority_location(directories)
    context,catalog=_load_release_semver_authority(sprint,scope,directories)
    packets=_semver_verified_directories(directories,catalog);by_digest={p['packet_sha256']:p for p in packets};available={}
    try:
        for packet in packets:
            for key,fact in packet['receipt']['rows'].items():
                try:
                    _semver_check_reused_origin(packet,key,by_digest)
                    latest=_semver_attempt_for_row(catalog,fact['input'],for_reuse=True)
                except ValueError:continue
                if latest['id']==packet['attempt']['id']:
                    available[key]={'packet':packet,'row':fact,'row_key':key}
        # Unused owned unpacked roots can be removed now; the executor closes selected roots.
        used={str(ref['packet']['owned_parent']) for ref in available.values()}
        for packet in packets:
            if str(packet['owned_parent']) not in used:shutil.rmtree(packet['owned_parent'])
        return available
    except BaseException:
        for packet in packets:
            if packet['owned_parent'].exists():shutil.rmtree(packet['owned_parent'])
        raise


def _semver_download_origin_proofs(proofs, catalog, destination):
    result=list(proofs);queue=list(proofs);seen=set();known={}
    while queue:
        directory=Path(queue.pop(0));hint=_semver_read(directory/'proof-selection.json')
        choices=[a for a in catalog['attempts'] if a['id']==hint['attempt_id']]
        if len(choices)!=1:raise ValueError('original proof attempt absent from independent catalog')
        packet=_semver_verify_packet(directory,choices[0])
        try:
            digest=packet['packet_sha256']
            if digest in seen:continue
            seen.add(digest)
            for key,fact in packet['receipt']['rows'].items():
                row=_semver_read(packet['root']/'payload'/fact['path']);reference=row.get('reused_from')
                if reference is None or reference['packet_sha256'] in seen:continue
                original=[a for a in catalog['attempts'] if a['id']==reference['original_attempt']]
                if len(original)!=1 or original[0].get('artifact') is None:
                    raise ValueError('required original proof transport is unavailable')
                ref_digest=reference['packet_sha256']
                if ref_digest not in known:
                    original_directory=_semver_download_proof(original[0],Path(destination)/ref_digest)
                    if _semver_hash(original_directory/('semver-'+original[0]['target']+'.tar'))!=ref_digest:
                        raise ValueError('downloaded original subject differs from immutable reference')
                    known[ref_digest]=original_directory;result.append(original_directory);queue.append(original_directory)
        finally:shutil.rmtree(packet['owned_parent'])
    return result


def _semver_required_rows(receipt, scope, payload, fresh=None):
    expected=set()
    if not set(scope['packages']).issubset(receipt['scope']['packages']):raise ValueError('native family membership incomplete')
    for member in receipt['members']:
        name=member['name']
        if name not in scope['packages']:continue
        if member['version']!=scope['packages'][name]['version']:raise ValueError('native candidate version differs')
        raw=Path(payload)/'registry'/(name+'.json');recorded=select_semver_inventory(name,member['version'],_semver_read(raw))
        if receipt['registry'][name]['raw_sha256']!=_semver_hash(raw):raise ValueError('frozen registry raw identity differs')
        applicable=fresh[name] if fresh is not None else recorded
        if member['api']=='binary-only-cli':
            if receipt['applicability'][name]['api_tool']!='not-applicable':raise ValueError('binary-only member falsely claims library tool proof')
            continue
        if receipt['applicability'][name]['api_tool']!=recorded['applicability']:raise ValueError('initial-line applicability relabelled')
        for baseline in applicable['baselines']:
            _semver_validate_archive(payload,name,baseline['num'],baseline['checksum'])
            for view in ('default','all'):
                rows=[(key,fact) for key,fact in receipt['rows'].items() if fact['input']['kind']=='api' and fact['input']['crate']==name and fact['input']['version']==member['version'] and fact['input']['baseline']==baseline['num'] and fact['input']['view']==view and fact['input']['extra']['archive_sha256']==baseline['checksum']]
                if len(rows)!=1:raise ValueError('new/missing baseline or feature row makes evidence incomplete')
                expected.add(rows[0][0])
    historical=_semver_read(Path(payload)/'registry/historical-opc.json')
    shared=scope['packages'].get('oxml-opc',{}).get('version')
    if shared is None:
        shared=tomllib.loads((REPO/'crates/oxml-opc/Cargo.toml').read_text())['package']['version']
    select_semver_inventory('oxml-opc',shared,historical)
    for canary,subject in SEMVER_CANARY_OUTCOMES:
        if subject=='current-candidate' and (canary=='declared-arc' and 'rpptx' not in scope['families'] or canary=='word-nominal' and 'rdocx' not in scope['families']):continue
        name='rdocx' if canary=='word-nominal' else 'oxml-opc'
        version=scope['packages'][name]['version'] if subject=='current-candidate' else subject.removeprefix('published-')
        rows=[(key,fact) for key,fact in receipt['rows'].items() if fact['input']['kind']=='canary' and fact['input']['crate']==name and fact['input']['version']==version and fact['input']['baseline']==subject and fact['input']['extra']['canary']==canary]
        if len(rows)!=1:raise ValueError('required exact-version compiler witness missing/ambiguous')
        key,fact=rows[0];expected.add(key)
        if subject!='current-candidate':
            baselines=[b for b in historical['versions'] if b['num']==version]
            if len(baselines)!=1 or fact['input']['extra']['archive_sha256']!=baselines[0]['checksum']:
                raise ValueError('historical witness checksum differs from complete frozen registry')
            _semver_validate_archive(payload,'oxml-opc',version,baselines[0]['checksum'])
    return expected


def _consume_release_semver(scope, directories, context, catalog, *, refresh_registry=False, release_candidate=False):
    if context.get('purpose')=='sign':raise ValueError('signing-only context cannot decide evidence')
    packets=_semver_verified_directories(directories,catalog);by_digest={p['packet_sha256']:p for p in packets};verified={}
    work=Path(tempfile.mkdtemp(prefix='rdocx-semver-registry-')) if refresh_registry else None
    try:
        fresh=_semver_registry(scope,work,publication=release_candidate) if work else None
        for target in SEMVER_NATIVE_TOOLS:
            candidates=[a for a in catalog['attempts'] if a['target']==target and a['current_source_equivalent']]
            choice=select_semver_attempt([{**a,'row_key':'native-union-'+target,'input_sha256':context['source_identity']} for a in candidates])
            matches=[p for p in packets if p['attempt']['id']==choice['id']]
            if len(matches)!=1:raise ValueError('required latest signed native packet missing/ambiguous')
            packet=matches[0];receipt=packet['receipt']
            expected=_semver_required_rows(receipt,scope,packet['root']/'payload',fresh)
            for key in expected:
                _semver_check_reused_origin(packet,key,by_digest)
                latest=_semver_attempt_for_row(catalog,receipt['rows'][key]['input'])
                if latest['id']!=choice['id']:raise ValueError('newer failed/incomparable relevant attempt blocks reuse')
            verified[target]={'packet_sha256':packet['packet_sha256'],'attempt':choice['id'],'rows':len(expected)}
        if release_candidate and context['purpose']!='external':raise ValueError('publication requires completed external CI authority')
        if context['purpose']=='external':
            for target in SEMVER_NATIVE_TOOLS:
                current=[a for a in catalog['attempts'] if a['target']==target and a['source_head']==context['head']]
                select_semver_attempt([{**a,'row_key':'final-ci-'+target,'input_sha256':context['source_identity']} for a in current])
        return {'complete':True,'scope':scope['scope'],'native':verified,'authority_cutoff':context['cutoff'],
            'catalog_sha256':context['catalog_sha256'],'registry_refreshed':bool(fresh),'publication_candidate_checked':release_candidate,
            'qualification':'AUTHENTICATED_NATIVE_SEMVER_UNION_NOT_PUBLICATION_APPROVAL'}
    finally:
        for packet in packets:shutil.rmtree(packet['owned_parent'])
        if work:shutil.rmtree(work)


def validate_unified_release_family(tag: str) -> dict[str, object]:
    """Check the exact selected Rust and Python version carriers for a new tag."""
    match = PYTHON_RELEASE_TAG_RE.fullmatch(tag)
    if match is None or match.group("legacy") is not None:
        raise ValueError(f"{tag!r} is not a unified release tag")
    distribution = "rdocx" if match.group("stable") else "rpptx"
    version = match.group("version")
    workspace = tomllib.loads((REPO / "Cargo.toml").read_text(encoding="utf-8"))
    if distribution == "rdocx" and workspace["workspace"]["package"]["version"] != version:
        raise ValueError("stable workspace version differs from release tag")
    dependencies = workspace["workspace"]["dependencies"]
    selected = set(RELEASE_CRATES[distribution])
    publishable = set()
    lock = tomllib.loads((REPO / "Cargo.lock").read_text(encoding="utf-8"))
    locked = {(package["name"], package["version"]) for package in lock["package"]}
    for member in workspace["workspace"]["members"]:
        manifest = tomllib.loads((REPO / member / "Cargo.toml").read_text(encoding="utf-8"))
        package = manifest["package"]
        carrier = package["version"]
        member_version = workspace["workspace"]["package"]["version"] if isinstance(carrier, dict) and carrier.get("workspace") is True else carrier
        registry = package.get("publish")
        if member_version == version and (registry is None or registry is True or isinstance(registry, list) and "crates-io" in registry):
            publishable.add(package["name"])
    if publishable != selected:
        raise ValueError(f"selected publishable crates differ: missing={sorted(selected - publishable)}, extra={sorted(publishable - selected)}")
    for crate in (*RELEASE_CRATES[distribution], f"{distribution}-py"):
        manifest = tomllib.loads(
            (REPO / "crates" / crate / "Cargo.toml").read_text(encoding="utf-8")
        )
        carrier = manifest["package"]["version"]
        actual = workspace["workspace"]["package"]["version"] if isinstance(carrier, dict) and carrier.get("workspace") is True else carrier
        if actual != version:
            raise ValueError(f"{crate} is {actual}, expected {version}")
        if crate in selected and (crate, version) not in locked:
            raise ValueError(f"Cargo.lock lacks {crate} {version}")
        if crate in dependencies and dependencies[crate]["version"] != version:
            raise ValueError(f"workspace dependency {crate} differs from {version}")
    project = tomllib.loads(
        (REPO / "crates" / f"{distribution}-py" / "pyproject.toml").read_text(encoding="utf-8")
    )
    if project["project"]["version"] != version:
        raise ValueError(f"{distribution} Python distribution differs from {version}")
    return {"tag": tag, "distribution": distribution, "version": version}


def validate_release_package_inventory(tag: str, directory: Path) -> None:
    """Check the unpacked packages left by a workspace publication dry run."""
    family = validate_unified_release_family(tag)
    version = str(family["version"])
    expected = {
        f"{crate}-{version}" for crate in RELEASE_CRATES[str(family["distribution"])]
    }
    actual = {path.name for path in directory.glob(f"*-{version}")}
    if actual != expected:
        raise ValueError(
            f"selected package inventory differs: missing={sorted(expected - actual)}, "
            f"extra={sorted(actual - expected)}"
        )
    for name in expected:
        package = directory / name
        if not package.is_dir() or not (package / "Cargo.toml").is_file():
            raise ValueError(f"dry-run package is incomplete: {name}")


def cmd_unified_release_family(args: argparse.Namespace) -> int:
    try:
        result = validate_unified_release_family(args.tag)
    except (KeyError, OSError, TypeError, ValueError) as error:
        die(f"unified release family: {error}")
    print(f"unified-release-family {result['tag']}: {result['distribution']} {result['version']}, ok")
    return 0


def verify_published_unified_release(tag: str, directory: Path) -> dict[str, object]:
    """Confirm both registries expose the selected exact version and Python bytes."""
    from hashlib import sha256

    family = validate_unified_release_family(tag)
    distribution = str(family["distribution"])
    version = str(family["version"])
    expected_python = {
        path.name: sha256(path.read_bytes()).hexdigest()
        for path in directory.iterdir()
        if path.name.startswith(f"{distribution}-{version}-")
        or path.name == f"{distribution}-{version}.tar.gz"
    }
    if len(expected_python) != 7:
        raise ValueError("reviewed family assets lack seven Python distributions")

    def fetch_json(url: str) -> dict[str, object]:
        request = urllib.request.Request(url, headers={"User-Agent": "rdocx-release/1"})
        with urllib.request.urlopen(request, timeout=30) as response:
            value = json.load(response)
        if not isinstance(value, dict):
            raise ValueError(f"invalid registry response from {url}")
        return value

    pending = set(RELEASE_CRATES[distribution])
    for attempt in range(20):
        for crate in tuple(sorted(pending)):
            try:
                result = fetch_json(f"https://crates.io/api/v1/crates/{crate}/{version}")
                entry = result.get("version")
                if isinstance(entry, dict) and entry.get("num") == version:
                    pending.remove(crate)
            except (urllib.error.HTTPError, urllib.error.URLError):
                pass
        try:
            project = fetch_json(f"https://pypi.org/pypi/{distribution}/{version}/json")
            urls = project.get("urls")
            if not isinstance(urls, list):
                raise ValueError("PyPI release lacks a file list")
            actual_python = {
                entry["filename"]: entry["digests"]["sha256"]
                for entry in urls
                if isinstance(entry, dict)
            }
            python_ready = actual_python == expected_python
        except (KeyError, TypeError, urllib.error.HTTPError, urllib.error.URLError):
            python_ready = False
        if not pending and python_ready:
            return family
        if attempt < 19:
            time.sleep(30)
    raise ValueError(
        f"registry publication incomplete: crates={sorted(pending)}, "
        f"PyPI exact files={python_ready}"
    )


def cmd_published_unified_release(args: argparse.Namespace) -> int:
    try:
        result = verify_published_unified_release(args.tag, Path(args.directory))
    except (OSError, TypeError, ValueError) as error:
        die(f"published unified release: {error}")
    print(f"published-unified-release {result['tag']}: both registries, ok")
    return 0


def validate_python_release_artifacts(tag: str, directory: Path) -> dict[str, object]:
    """Validate the exact selected Python distribution artifact set."""
    match = PYTHON_RELEASE_TAG_RE.fullmatch(tag)
    if match is None:
        raise ValueError(
            f"{tag!r} is not a Python release tag, expected "
            "vX.Y.Z or rpptx-vX.Y.Z"
        )
    distribution = "rdocx" if match.group("stable") else ("rpptx" if match.group("incubating") else match.group("legacy"))
    version = match.group("version")
    if not directory.is_dir():
        raise ValueError(f"Python release artifact directory is missing: {directory}")

    expected_wheels = {
        f"{distribution}-{version}-cp39-abi3-{platform}.whl"
        for platform in PYTHON_RELEASE_PLATFORMS
    }
    expected_sdists = {f"{distribution}-{version}.tar.gz"}
    expected_names = expected_wheels | expected_sdists
    entries = sorted(directory.iterdir(), key=lambda path: path.name)
    actual_names = {path.name for path in entries}
    if len(entries) != len(actual_names):
        raise ValueError("Python release artifact directory has duplicate names")
    if actual_names != expected_names:
        missing = sorted(expected_names - actual_names)
        unexpected = sorted(actual_names - expected_names)
        raise ValueError(
            "Python release artifact set differs from the exact family, "
            f"missing={missing}, unexpected={unexpected}"
        )
    if any(not path.is_file() for path in entries):
        raise ValueError("Python release artifact set contains a non-file entry")

    for name in sorted(expected_wheels):
        path = directory / name
        distribution = name.split("-", 1)[0]
        platform = name.removesuffix(".whl").rsplit("-", 1)[1]
        dist_info = f"{distribution}-{version}.dist-info"
        metadata_name = f"{dist_info}/METADATA"
        wheel_name = f"{dist_info}/WHEEL"
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            metadata_records = [
                member
                for member in names
                if member.endswith(".dist-info/METADATA")
                or member.endswith(".dist-info/WHEEL")
            ]
            if sorted(metadata_records) != sorted((metadata_name, wheel_name)):
                raise ValueError(
                    f"{name} has unexpected distribution metadata records"
                )
            if names.count(metadata_name) != 1 or names.count(wheel_name) != 1:
                raise ValueError(f"{name} lacks one exact METADATA and WHEEL record")
            metadata = Parser().parsestr(
                archive.read(metadata_name).decode("utf-8", errors="strict")
            )
            wheel = Parser().parsestr(
                archive.read(wheel_name).decode("utf-8", errors="strict")
            )
        if metadata.get("Name") != distribution:
            raise ValueError(f"{name} has project name {metadata.get('Name')!r}")
        if metadata.get("Version") != version:
            raise ValueError(f"{name} has project version {metadata.get('Version')!r}")
        validate_python_project_metadata(metadata, distribution, name)
        expected_tag = f"cp39-abi3-{platform}"
        if wheel.get_all("Tag", []) != [expected_tag]:
            raise ValueError(
                f"{name} has wheel tags {wheel.get_all('Tag', [])!r}, "
                f"expected {[expected_tag]!r}"
            )

    for name in sorted(expected_sdists):
        path = directory / name
        distribution = name.split("-", 1)[0]
        root = f"{distribution}-{version}"
        metadata_name = f"{root}/PKG-INFO"
        with tarfile.open(path, mode="r:gz") as archive:
            members = archive.getmembers()
            names = [member.name for member in members]
            if any(
                member.name.startswith("/")
                or ".." in Path(member.name).parts
                or not (
                    member.name == root or member.name.startswith(f"{root}/")
                )
                for member in members
            ):
                raise ValueError(f"{name} contains a path outside {root}/")
            if names.count(metadata_name) != 1:
                raise ValueError(f"{name} lacks one exact PKG-INFO record")
            metadata_member = archive.getmember(metadata_name)
            if not metadata_member.isfile():
                raise ValueError(f"{name} PKG-INFO is not a regular file")
            metadata_file = archive.extractfile(metadata_member)
            if metadata_file is None:
                raise ValueError(f"{name} PKG-INFO is not a regular file")
            metadata = Parser().parsestr(
                metadata_file.read().decode("utf-8", errors="strict")
            )
        if metadata.get("Name") != distribution:
            raise ValueError(f"{name} has project name {metadata.get('Name')!r}")
        if metadata.get("Version") != version:
            raise ValueError(f"{name} has project version {metadata.get('Version')!r}")
        validate_python_project_metadata(metadata, distribution, name)

    return {
        "tag": tag,
        "version": version,
        "distributions": (distribution,),
        "wheels": len(expected_wheels),
        "sdists": len(expected_sdists),
    }


def cmd_python_release_artifacts(args: argparse.Namespace) -> int:
    try:
        result = validate_python_release_artifacts(args.tag, Path(args.directory))
    except (OSError, UnicodeError, ValueError, tarfile.TarError, zipfile.BadZipFile) as error:
        die(f"python release artifacts: {error}")
    print(
        f"python-release-artifacts {result['tag']}: "
        f"{result['wheels']} wheels, {result['sdists']} source distributions, ok"
    )
    return 0


def state_path(sprint: str) -> Path:
    return SCRATCH / f"{sprint}-run.json"


def load(sprint: str) -> dict:
    p = state_path(sprint)
    if not p.exists():
        die(f"no run state for {sprint}. Run `init {sprint}` first.")
    data = json.loads(p.read_text(encoding="utf-8"))
    if data.get("schema_version") != SCHEMA_VERSION:
        die(
            f"run state schema {data.get('schema_version')} != {SCHEMA_VERSION}. "
            "Delete the file and re-init."
        )
    return data


def save(sprint: str, data: dict) -> None:
    SCRATCH.mkdir(parents=True, exist_ok=True)
    state_path(sprint).write_text(
        json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )


def git_head() -> str:
    """Return the commit whose evidence is being recorded or checked."""
    result = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        cwd=REPO,
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout.strip()


def parse_current_sprint() -> tuple[str, list[dict]]:
    if not CURRENT_SPRINT.exists():
        die(f"{CURRENT_SPRINT} not found")
    text = CURRENT_SPRINT.read_text(encoding="utf-8")
    m = SPRINT_RE.search(text)
    if not m:
        die("CURRENT_SPRINT.md has no parseable `# Current Sprint, SNN` heading")
    features = [
        {"fid": r[0], "title": r[1], "size": r[2], "tracker_status": r[3], "owner": r[4]}
        for r in WAVE_ROW_RE.findall(text)
    ]
    if not features and not VALIDATION_ONLY_RE.search(text):
        die(
            "CURRENT_SPRINT.md has no parseable wave rows. An intentionally "
            "empty sprint must declare `**Validation-only**: yes`."
        )
    return m.group(1), features


def backlog_statuses() -> dict[str, str]:
    if not BACKLOG.exists():
        die(f"{BACKLOG} not found")
    return {r[0]: r[2] for r in BACKLOG_ROW_RE.findall(BACKLOG.read_text(encoding="utf-8"))}


def completed_record_problems(sprint: str, fid: str) -> list[str]:
    """Return missing or stale durable records for one completed feature."""
    problems: list[str] = []
    current_id, current_features = parse_current_sprint()
    current = {feature["fid"]: feature for feature in current_features}
    feature = current.get(fid)
    if current_id != sprint:
        problems.append(f"CURRENT_SPRINT.md says {current_id}, run state says {sprint}")
    elif feature is None:
        problems.append(f"{fid} has no CURRENT_SPRINT.md row")
    else:
        if feature["tracker_status"] != "done":
            problems.append(
                f"{fid} is completed in run state but "
                f"'{feature['tracker_status']}' in CURRENT_SPRINT.md"
            )
        if feature["owner"] not in ("", "-"):
            problems.append(
                f"{fid} is completed but CURRENT_SPRINT.md owner is "
                f"'{feature['owner']}'"
            )

    plan = PLANS / f"{fid}-design.md"
    if not plan.exists():
        problems.append(f"{fid} has no design plan at {plan.relative_to(REPO)}")
    elif "**Status**: completed" not in plan.read_text(encoding="utf-8"):
        problems.append(f"{fid} design plan is not completed")

    if not AS_BUILT.exists() or not re.search(
        rf"^### {re.escape(fid)},", AS_BUILT.read_text(encoding="utf-8"), re.MULTILINE
    ):
        problems.append(f"{fid} has no AS_BUILT.md entry")

    tracker_text = (
        SPRINT_TRACKER.read_text(encoding="utf-8") if SPRINT_TRACKER.exists() else ""
    )
    if not re.search(
        rf"^\|\s*{re.escape(fid)}\s*\|\s*{re.escape(sprint)}\s*\|",
        tracker_text,
        re.MULTILINE,
    ):
        problems.append(f"{fid} has no {sprint} row in SPRINT_TRACKER.md")
    return problems


def completed_owner_problems(data: dict) -> list[str]:
    """Return completed features whose run-state owner was not cleared."""
    return [
        f"{fid} is completed but run-state owner is '{feature['owner']}'"
        for fid, feature in data["features"].items()
        if feature["state"] == "completed" and feature.get("owner") is not None
    ]


def cmd_init(args: argparse.Namespace) -> int:
    sprint, features = parse_current_sprint()
    if args.sprint != sprint:
        die(f"CURRENT_SPRINT.md says {sprint}, you asked for {args.sprint}")

    exists = state_path(sprint).exists()
    if exists and args.resume:
        # The point of --resume is that an interrupted /run-sprint picks up
        # where it stopped rather than redoing landed work. Adopt the limits
        # and canonical feature metadata from this invocation, while keeping
        # every state, worker fact, and evidence record.
        data = load(sprint)
        data["max_review_passes"] = args.max_review_passes
        data["max_workers"] = args.max_workers
        reviewer = getattr(args, "reviewer", None)
        if reviewer:
            if data.get("reviewer") not in (None, reviewer) and data["reviews"]:
                die("cannot replace the requested reviewer after review evidence exists")
            data["reviewer"] = reviewer
        for f in features:
            feature = data["features"].setdefault(
                f["fid"],
                {
                    "state": "pending",
                    "size": f["size"],
                    "title": f["title"],
                    "owner": f["owner"] if f["owner"] != "-" else None,
                },
            )
            feature["size"] = f["size"]
            feature["title"] = f["title"]
        save(sprint, data)
        print(
            f"resumed {sprint}, phase={data['phase']}, "
            f"{len(data['features'])} features, {len(data['reviews'])} review pass(es)"
        )
        return 0
    if exists and not args.force:
        die(
            f"run state for {sprint} already exists. Pass --resume to continue "
            "it, or --force to throw it away and recreate."
        )

    data = {
        "schema_version": SCHEMA_VERSION,
        "sprint": sprint,
        "phase": "design",
        "max_review_passes": args.max_review_passes,
        "max_workers": args.max_workers,
        "reviewer": getattr(args, "reviewer", None),
        "features": {
            f["fid"]: {
                "state": "pending",
                "size": f["size"],
                "title": f["title"],
                "owner": f["owner"] if f["owner"] != "-" else None,
            }
            for f in features
        },
        "reviews": [],
        "verifications": [],
    }
    save(sprint, data)
    print(f"initialised {sprint} with {len(features)} features, phase=design")
    return 0


def cmd_status(args: argparse.Namespace) -> int:
    sprint = args.sprint or parse_current_sprint()[0]
    data = load(sprint)
    print(
        f"{sprint}  phase={data['phase']}  "
        f"max-review-passes={data.get('max_review_passes', MAX_REVIEW_PASSES)}  "
        f"max-workers={data.get('max_workers') or 'unbounded'}"
    )
    counts: dict[str, int] = {}
    for fid, f in sorted(data["features"].items()):
        counts[f["state"]] = counts.get(f["state"], 0) + 1
        owner = f" ({f['owner']})" if f.get("owner") else ""
        wave = f" w{f['wave']}" if f.get("wave") is not None else ""
        print(f"  {fid:<8} {f['size']:<2}{wave} {f['state']}{owner}  {f['title']}")
        if args.workers:
            for field in WORKER_FIELDS:
                if field != "wave" and f.get(field):
                    print(f"      {field:<19} {f[field]}")
    summary = ", ".join(f"{k}={v}" for k, v in sorted(counts.items()))
    print("  " + (summary or "features=0"))
    if data["reviews"]:
        last = data["reviews"][-1]
        print(f"  reviews: {len(data['reviews'])} passes, last blocking={last['blocking']}")
    return 0


def cmd_set_phase(args: argparse.Namespace) -> int:
    if args.phase not in PHASES:
        die(f"unknown phase {args.phase}. One of: {', '.join(PHASES)}")
    data = load(args.sprint)
    if data["phase"] == "closed":
        die(f"{args.sprint} is closed. Phases do not move after that.")
    data["phase"] = args.phase
    save(args.sprint, data)
    print(f"{args.sprint} phase={args.phase}")
    return 0


def cmd_mark_feature(args: argparse.Namespace) -> int:
    if not FID_RE.match(args.fid):
        die(f"{args.fid} is not a valid F-ID")
    if args.state not in FEATURE_STATES:
        die(f"unknown state {args.state}. One of: {', '.join(FEATURE_STATES)}")
    data = load(args.sprint)
    f = data["features"].get(args.fid)
    if f is None:
        die(f"{args.fid} is not in {args.sprint}")
    old = f["state"]
    allowed = FEATURE_TRANSITIONS[old]
    if args.state != old and args.state not in allowed:
        die(
            f"{args.fid}: {old} -> {args.state} is not a legal transition. "
            f"From {old} you may go to: {', '.join(sorted(allowed)) or '(nothing)'}"
        )
    f["state"] = args.state
    if args.owner:
        f["owner"] = args.owner
    if args.clear_owner:
        f["owner"] = None
    for field in WORKER_FIELDS:
        value = getattr(args, field, None)
        if value is not None:
            f[field] = value
    save(args.sprint, data)
    print(f"{args.fid}: {old} -> {args.state}")
    return 0


def cmd_record_review(args: argparse.Namespace) -> int:
    data = load(args.sprint)
    head = git_head()
    reviewer = getattr(args, "reviewer", None)
    if data.get("reviewer") and reviewer != data["reviewer"]:
        die(f"sprint review requires the requested reviewer {data['reviewer']}")
    bound = data.get("max_review_passes", MAX_REVIEW_PASSES)
    if args.passno > bound and not args.extend:
        die(
            f"pass {args.passno} exceeds the bound of {bound}. "
            "Another pass means the sprint is not ready. Carry work, or pass "
            "--extend and record why in the review file."
        )
    data["reviews"].append(
        {
            "pass": args.passno,
            "blocking": args.blocking,
            "should_fix": args.should_fix,
            "nice_to_have": args.nice_to_have,
            "head": head,
            "reviewer": reviewer,
        }
    )
    save(args.sprint, data)
    verdict = "clean" if args.blocking == 0 else f"{args.blocking} blocking"
    print(
        f"{args.sprint} review pass {args.passno} recorded at {head[:12]}: {verdict}"
    )
    return 0


def capability_owner_problems(scope: str, backlog: str) -> list[str]:
    """Check the capability ownership facts that completion metadata can change."""
    statuses = {fid: status for fid, _, status in BACKLOG_ROW_RE.findall(backlog)}
    problems = []
    rows = [line for line in scope.splitlines() if line.startswith("| DOCX-")]
    if not rows:
        return ["capability matrix has no parseable DOCX rows"]
    for line in rows:
        cells = [cell.strip() for cell in line.strip("|").split("|")]
        if len(cells) != 19:
            problems.append(f"malformed capability row: {cells[0]}")
            continue
        capability, classification, owner = cells[0], cells[-3], cells[-1]
        if classification in ("partial", "unsupported"):
            status = statuses.get(owner)
            if status not in ("pending", "in-progress"):
                problems.append(f"{capability} incomplete capability owner {owner} is {status or 'missing'}")
        elif owner != "-":
            problems.append(f"{capability} resolved capability still has active owner {owner}")
    return problems


def metadata_reuse_problems(data: dict, head: str, record: dict) -> list[str]:
    """Authenticate a finite delivery-only range against the latest actual full run."""
    full = [(i, item) for i, item in enumerate(data["verifications"])
            if item["scope"] == "full"]
    if not full:
        return ["metadata reuse requires an actual full verification attempt"]
    index, latest = full[-1]
    source = record.get("source_head", "")
    if (record.get("full_attempt") != index or not latest["passed"]
            or latest.get("head") != source):
        return ["metadata reuse must anchor the latest passing full attempt"]
    if not re.fullmatch(r"[0-9a-f]{40}", source) or not re.fullmatch(r"[0-9a-f]{40}", head):
        return ["metadata reuse requires full commit identities"]

    def git(*args: str) -> str:
        return subprocess.run(["git", *args], cwd=REPO, check=True,
                              capture_output=True).stdout.decode("utf-8")

    try:
        if git("status", "--porcelain"):
            return ["metadata reuse requires a clean worktree"]
        git("merge-base", "--is-ancestor", source, head)
        ledgers = {"docs/sprints/" + name + ".md" for name in
                   ("CURRENT_SPRINT", "BACKLOG", "SPRINT_TRACKER", "AS_BUILT")}
        own_reviews = {f".claude/reviews/{data['sprint']}-sprint-review-pass-{review['pass']}.md"
                       for review in data.get("reviews", [])}
        plans = {f".claude/plans/{fid}-design.md" for fid in data["features"]}
        if git("rev-list", "--min-parents=2", f"{source}..{head}"):
            return ["metadata reuse requires a linear delivery-only commit range"]
        for commit in git("rev-list", "--reverse", f"{source}..{head}").splitlines():
            if prose_check.check_text(git("log", "-1", "--format=%B", commit), "commit message"):
                return [f"metadata reuse refuses invalid commit prose: {commit}"]
            parent = git("rev-parse", f"{commit}^").strip()
            paths = git("diff", "--name-only", "--no-renames", "-z", parent, commit).split("\0")
            for path in filter(None, paths):
                # Deletions, symlinks and mode changes cannot be delivery metadata.
                after_entry = git("ls-tree", commit, "--", path).split()
                before_entry = git("ls-tree", parent, "--", path).split()
                if not after_entry or after_entry[0] != "100644":
                    return [f"metadata reuse refuses deletion or non-regular file: {path}"]
                if before_entry and before_entry[0] != "100644":
                    return [f"metadata reuse refuses mode change: {path}"]
                after = git("show", f"{commit}:{path}")
                if prose_check.check_text(after, path):
                    return [f"metadata reuse refuses invalid delivery prose: {path}"]
                if path in own_reviews and not before_entry:
                    continue
                if not before_entry:
                    return [f"metadata reuse refuses new file: {path}"]
                before = git("show", f"{parent}:{path}")
                if path in ledgers:
                    if path.endswith("AS_BUILT.md") and not after.startswith(before):
                        return ["metadata reuse requires append-only AS_BUILT.md"]
                    if path.endswith("CURRENT_SPRINT.md") and not re.search(
                            rf"^# Current Sprint, {re.escape(data['sprint'])}$", after, re.M):
                        return ["metadata reuse cannot change the current sprint identity"]
                    continue
                if path in plans:
                    old_lines, new_lines = before.splitlines(keepends=True), after.splitlines(keepends=True)
                    if len(old_lines) != len(new_lines):
                        return [f"metadata reuse refuses semantic plan edits: {path}"]
                    for old, new in zip(old_lines, new_lines):
                        if old == new or (old.startswith("- [ ] ") and new == old.replace("- [ ] ", "- [x] ", 1)):
                            continue
                        if (re.fullmatch(r"\*\*Status\*\*: (approved|in-progress|reviewed)\n?", old)
                                and new == "**Status**: completed" + ("\n" if old.endswith("\n") else "")):
                            continue
                        return [f"metadata reuse refuses semantic plan edits: {path}"]
                    continue
                return [f"metadata reuse refuses non-delivery change: {path}"]
        problems = capability_owner_problems(
            (REPO / "docs/hld/02-scope-and-non-goals.md").read_text(encoding="utf-8"),
            (REPO / "docs/sprints/BACKLOG.md").read_text(encoding="utf-8"),
        )
        if problems:
            return problems
    except (subprocess.CalledProcessError, UnicodeError, OSError) as error:
        return [f"metadata reuse cannot authenticate commit range: {error}"]
    return []


def cmd_record_verification(args: argparse.Namespace) -> int:
    data = load(args.sprint)
    head = git_head()
    record = {"scope": args.scope, "passed": args.passed,
              "harness": args.harness, "head": head}
    source = getattr(args, "reuse_from", None)
    if args.scope == "metadata":
        if not source:
            die("metadata verification requires --reuse-from with the full source SHA")
        full = [i for i, item in enumerate(data["verifications"]) if item["scope"] == "full"]
        record.update(source_head=source, full_attempt=full[-1] if full else None)
        problems = metadata_reuse_problems(data, head, record)
        if problems:
            die("; ".join(problems))
    elif source:
        die("--reuse-from is only valid for metadata verification")
    data["verifications"].append(record)
    save(args.sprint, data)
    print(f"{args.sprint} verification recorded at {head[:12]}: "
          f"{args.scope}, passed={args.passed}")
    return 0


def closure_evidence_problems(data: dict, head: str) -> list[str]:
    """Return stale or missing evidence, with failures superseding prior green."""
    problems: list[str] = []
    if not data["reviews"]:
        problems.append("no sprint-review pass recorded")
    else:
        review = data["reviews"][-1]
        if review["blocking"] != 0:
            problems.append(f"last review pass has {review['blocking']} blocking finding(s)")
        if data.get("reviewer") and review.get("reviewer") != data["reviewer"]:
            problems.append(f"latest sprint review is not from requested reviewer {data['reviewer']}")
        if review.get("head") != head:
            problems.append("latest sprint review covered "
                            f"{review.get('head') or 'an unrecorded HEAD'}, current HEAD is {head}")
    acceptance = [item for item in data["verifications"] if item["scope"] in ("full", "metadata")]
    if not acceptance or not acceptance[-1]["passed"] or acceptance[-1].get("head") != head:
        problems.append(f"no passing `/verify --full` recorded for current HEAD {head}")
    elif acceptance[-1]["scope"] == "metadata":
        problems.extend(metadata_reuse_problems(data, head, acceptance[-1]))
    return problems


def cmd_validate_handoff(args: argparse.Namespace) -> int:
    """Check a worker's `.claude/handoffs/F-XXX-ready.md` before integration.

    This is a shape check, not a trust check. It catches the half-written
    handoff. Whether the claims inside are true is what /integrate-feature
    verifies against the branch.
    """
    path = Path(args.path)
    if not path.exists():
        die(f"{path} not found. The worker never ran `/complete-feature --prepare`.")

    fields = dict(HANDOFF_FIELD_RE.findall(path.read_text(encoding="utf-8")))
    problems: list[str] = []

    for key in HANDOFF_REQUIRED:
        if not fields.get(key):
            problems.append(f"missing or empty field: **{key}**")

    if fields.get("F-ID") and fields["F-ID"] != args.fid:
        problems.append(f"handoff is for {fields['F-ID']}, you asked for {args.fid}")

    microscope = fields.get("Microscope", "")
    if microscope and not ("0 defects" in microscope and "0 smells" in microscope):
        problems.append(
            "Microscope must cite a pass reporting '0 defects, 0 smells'. "
            f"It says: {microscope}"
        )

    verify = fields.get("Verify", "")
    if verify and "fail" in verify.lower():
        problems.append(f"Verify records a failure: {verify}")
    expected_verify = f"/verify --scoped {args.fid}"
    if verify and expected_verify not in verify:
        problems.append(
            f"Verify must record `{expected_verify}, pass`."
        )

    checkpoint = fields.get("Checkpoint", "")
    external = fields.get("External gates", "")
    test_gate = fields.get("Test gate", "")
    pending = any("external ci pending" in value.lower() for value in (verify, test_gate)) or bool(external)
    if checkpoint and checkpoint not in ("external-ci-source-availability", "sprint-final-acceptance"):
        problems.append("unknown handoff checkpoint")
    if pending and checkpoint != "external-ci-source-availability":
        problems.append("pending external CI requires its source-availability checkpoint")
    final = fields.get("Final checks", "")
    final_pending = any("final sprint checks pending" in value.lower() for value in (verify, test_gate)) or bool(final)
    if final_pending and checkpoint != "sprint-final-acceptance":
        problems.append("pending final sprint checks require their final-acceptance checkpoint")
    if checkpoint == "sprint-final-acceptance":
        if not re.fullmatch(r"[A-Za-z0-9/][A-Za-z0-9 ./_,:()=-]{2,400}, pending", final) or final[:-9].strip().lower() in ("checks", "gates", "final checks", "final sprint checks"):
            problems.append("final checkpoint requires named actual final checks, pending")
        if verify != f"{expected_verify}, local checks pass, final sprint checks pending":
            problems.append("final checkpoint must distinguish local checks from pending sprint acceptance")
        if (not test_gate.endswith(", pass") or len(test_gate[:-6].strip()) < 3
                or re.search(r"\b(?:complet(?:e|ed)|pending)\b", test_gate, re.I)):
            problems.append("final checkpoint requires named actual local regression checks, pass")
        if external:
            problems.append("final checkpoint cannot replace the external CI checkpoint")
    if checkpoint == "external-ci-source-availability":
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9 ./_-]{2,240}, pending", external) or external[:-9].strip().lower() in ("ci", "external ci", "gates"):
            problems.append("external checkpoint requires named actual gates, pending")
        if verify != f"{expected_verify}, local checks pass, external CI pending":
            problems.append("external checkpoint must distinguish local checks from pending scoped acceptance")
        if not test_gate.endswith(", pass, external CI pending") or re.search(r"\bcomplet(?:e|ed)\b", test_gate, re.I):
            problems.append("external checkpoint requires actual local test checks with external CI pending")

    if problems:
        print(f"validate-handoff {path}: REFUSED", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        return 1

    print(
        f"validate-handoff {args.fid}: ok, "
        f"{fields['Base'][:12]}..{fields['Head'][:12]} on {fields['Branch']}"
    )
    return 0


def cmd_close_preflight(args: argparse.Namespace) -> int:
    data = load(args.sprint)
    problems: list[str] = []
    head = git_head()

    if data["phase"] == "blocked":
        problems.append("sprint is blocked. Resolve what blocked it and re-run.")

    # /integrate-feature deletes each handoff it consumed. One still sitting
    # here means a worker was prepared and never integrated.
    stale = sorted(p.name for p in HANDOFFS.glob("F-*-ready.md"))
    if stale:
        problems.append("unconsumed worker handoffs: " + ", ".join(stale))

    incomplete = [
        fid
        for fid, f in data["features"].items()
        if f["state"] not in ("completed", "carried")
    ]
    if incomplete:
        problems.append(
            "features neither completed nor carried: " + ", ".join(sorted(incomplete))
        )
    problems.extend(completed_owner_problems(data))

    problems.extend(closure_evidence_problems(data, head))

    # The durable delivery record must agree with the run state. This catches a
    # completion that updated the JSON but missed one of the canonical ledgers.
    statuses = backlog_statuses()
    for fid, f in data["features"].items():
        want = "done" if f["state"] == "completed" else None
        if want and statuses.get(fid) != want:
            problems.append(
                f"{fid} is completed in run state but '{statuses.get(fid)}' in BACKLOG.md"
            )
        if want:
            problems.extend(completed_record_problems(args.sprint, fid))

    if data["reviews"]:
        last_pass = data["reviews"][-1]["pass"]
        review = REVIEWS / f"{args.sprint}-sprint-review-pass-{last_pass}.md"
        if not review.exists():
            problems.append(
                f"latest sprint review file is missing: {review.relative_to(REPO)}"
            )
        elif not re.search(
            r"^\*\*Verdict\*\*:\s*0 blocking,",
            review.read_text(encoding="utf-8"),
            re.MULTILINE,
        ):
            problems.append(
                "latest sprint review file does not report zero blocking: "
                f"{review.relative_to(REPO)}"
            )

    if problems:
        print(f"close-preflight {args.sprint}: REFUSED", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        return 1

    print(f"close-preflight {args.sprint}: ok, ready to close")
    return 0


COMPATIBILITY_COUNTS = {"BASE": 24, "XML": 14, "LIFE": 16, "AREA": 46}
COMPATIBILITY_MANIFEST_SHA = "7501e001301a703988526f76c9961b35b345a1c17f8e1d383e9477d5947b22e1"
COMPATIBILITY_BODY_SHA = "2f4348dd55dcd1b8889badd8ca407d9f790cf38327c00ec3ae31dd5b02b3f66e"
COMPATIBILITY_FREEZE_SHA = "d3cf81e655dfb1f7d35745487d46e2ceb8c3c826167439cc970911eec3ac21f7"
COMPATIBILITY_DIVERGENCES = {"W-XML-02", "P-AREA-17"}


def compatibility_artifact(reference: dict) -> bytes:
    """Authenticate retained bytes without executing their contents."""
    path = Path(reference["path"])
    digest = reference["sha256"]
    if not path.is_absolute() or not re.fullmatch(r"[0-9a-f]{64}", digest):
        raise ValueError("compatibility evidence needs an absolute path and SHA-256")
    raw = path.read_bytes()
    if hashlib.sha256(raw).hexdigest() != digest:
        raise ValueError(f"compatibility evidence changed: {path}")
    return raw


def compatibility_obligations(record: dict) -> list[str]:
    """Keep each frozen item and its explicit semicolon-separated clauses."""
    declared = record.get("additional_native_obligations", [])
    if isinstance(declared, str):
        declared = [declared]
    if not isinstance(declared, list) or any(not isinstance(text, str) for text in declared):
        raise ValueError("frozen native obligations must be texts")
    return [clause.strip() for text in declared for clause in text.split(";") if clause.strip()]


def compatibility_observations(facts: list) -> None:
    if not isinstance(facts, list) or not facts:
        raise ValueError("qualification needs concrete typed observations")
    for fact in facts:
        if not isinstance(fact, dict) or type(fact.get("observation")) is not str or (
            not fact["observation"].strip()
            or fact.get("kind") not in ("source-model", "save-reopen", "native", "render", "lifetime", "schema")
            or "expected" not in fact or "actual" not in fact
            or not isinstance(fact.get("evidence"), list) or not fact["evidence"]
        ):
            raise ValueError("qualification needs concrete typed observations")
        for reference in fact["evidence"]:
            compatibility_artifact(reference)


def compatibility_git_source(root: Path) -> dict:
    """Capture actual clean Git identity and every tracked working byte."""
    def git(*args):
        return subprocess.check_output(["git", "--no-optional-locks", *args], cwd=root)
    if git("status", "--porcelain", "--untracked-files=no").strip():
        raise ValueError("compatibility clean Git source has tracked changes")
    entries = [entry.split(b"\t", 1) for entry in git("ls-tree", "-r", "-z", "HEAD").split(b"\0") if entry]
    tracked = set(filter(None, git("ls-files", "-z").decode().split("\0")))
    if tracked != {name.decode() for _, name in entries}:
        raise ValueError("compatibility clean Git tracked paths differ from HEAD")
    object_format = git("rev-parse", "--show-object-format").decode().strip()
    files = {}
    for metadata, encoded_name in entries:
        mode, kind, oid = metadata.decode().split()
        name = encoded_name.decode()
        path = root / name
        if kind != "blob" or mode not in ("100644", "100755") or path.is_symlink() or not path.is_file():
            raise ValueError("compatibility source must contain regular tracked files")
        raw = path.read_bytes()
        git_object = hashlib.new(object_format, b"blob " + str(len(raw)).encode() + b"\0" + raw).hexdigest()
        if git_object != oid or stat.S_IMODE(path.stat().st_mode) != int(mode, 8) & 0o777:
            raise ValueError("compatibility clean Git working bytes or mode differ from HEAD")
        files[name] = {"sha256": hashlib.sha256(raw).hexdigest(), "size": len(raw),
                       "mode": stat.S_IMODE(path.stat().st_mode)}
    return {"head": git("rev-parse", "HEAD").decode().strip(),
            "tree": git("rev-parse", "HEAD^{tree}").decode().strip(),
            "index_sha256": hashlib.sha256(git("ls-files", "--stage", "-z")).hexdigest(),
            "files": files}


def compatibility_process(argv: list[str], cwd: Path, env: dict, timeout: int, log: Path) -> dict:
    """Bound this case session and leave an auditable lease for the outer job."""
    signals = []
    timed_out = False
    child = None
    pending_error = None
    old_handler = signal.getsignal(signal.SIGTERM)
    lease_path = log.with_suffix(".lease.json")
    lease = {"argv_sha256": hashlib.sha256(json.dumps(argv).encode()).hexdigest(), "cleanup": None}
    def interrupted(signum, frame):
        raise InterruptedError("compatibility case supervisor interrupted")
    signal.signal(signal.SIGTERM, interrupted)
    with log.open("xb") as output:
        try:
            # Do not expose the small spawn-to-lease interval to SIGTERM.
            old_mask = signal.pthread_sigmask(signal.SIG_BLOCK, {signal.SIGTERM})
            try:
                child = subprocess.Popen(argv, cwd=cwd, env=env, stdout=output, stderr=subprocess.STDOUT,
                                         start_new_session=True)
                identity = subprocess.check_output(["/bin/ps", "-p", str(child.pid), "-o", "uid=,pgid=,lstart="],
                                                   env=dict(os.environ, LC_ALL="C"), timeout=5).decode().strip()
                lease.update(pid=child.pid, pgid=child.pid, process_identity=identity)
                fd = os.open(lease_path, os.O_WRONLY | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW, 0o600)
                with os.fdopen(fd, "w") as stream:
                    json.dump(lease, stream)
            finally:
                signal.pthread_sigmask(signal.SIG_SETMASK, old_mask)
            child.wait(timeout=timeout)
        except BaseException as error:
            timed_out = isinstance(error, subprocess.TimeoutExpired)
            pending_error = None if timed_out else error
        finally:
            signal.signal(signal.SIGTERM, signal.SIG_IGN)
            if child is not None:
                for sig in (signal.SIGTERM, signal.SIGKILL):
                    try:
                        os.killpg(child.pid, sig)
                        signals.append(sig.name)
                    except ProcessLookupError:
                        break
                    try:
                        child.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        pass
                gone = False
                for _ in range(50):
                    try:
                        os.killpg(child.pid, 0)
                    except ProcessLookupError:
                        gone = True
                        break
                    time.sleep(0.1)
                result = {"exit_code": child.returncode, "timed_out": timed_out,
                          "cleanup": {"groupGone": gone, "leaderReaped": child.returncode is not None,
                                      "leaseReleaseSafe": gone and child.returncode is not None, "signals": signals}}
                lease.update(cleanup=result["cleanup"], exit_code=child.returncode)
                lease_path.write_text(json.dumps(lease, indent=2) + "\n")
            signal.signal(signal.SIGTERM, old_handler)
    if pending_error is not None:
        raise pending_error
    if child is None:
        raise ValueError("compatibility case process did not start")
    return result


def compatibility_execute(config: dict) -> dict:
    """Execute the original corpus once, with a separate approved input overlay."""
    if config.get("status") != "admitted":
        raise ValueError("compatibility execution configuration must be admitted")
    expected_identity = config["github_identity"]
    stable_identity = {"GITHUB_REPOSITORY", "GITHUB_REPOSITORY_ID", "GITHUB_ACTOR", "GITHUB_TRIGGERING_ACTOR",
                       "GITHUB_EVENT_NAME", "GITHUB_REF", "GITHUB_SHA", "GITHUB_WORKFLOW_REF",
                       "GITHUB_WORKFLOW_SHA", "GITHUB_JOB"}
    dynamic_identity = {"GITHUB_RUN_ID", "GITHUB_RUN_ATTEMPT"}
    identity = {key: os.environ.get(key) for key in stable_identity | dynamic_identity}
    if set(expected_identity) != stable_identity or any(identity[key] != value for key, value in expected_identity.items()) or (
        identity["GITHUB_REPOSITORY"] != "tensorbee/rdocx" or identity["GITHUB_ACTOR"] != "mantissaman"
        or identity["GITHUB_TRIGGERING_ACTOR"] != "mantissaman" or identity["GITHUB_EVENT_NAME"] != "workflow_dispatch"
        or identity["GITHUB_JOB"] != "owner-compatibility" or identity["GITHUB_SHA"] != config["head"]
        or identity["GITHUB_WORKFLOW_SHA"] != config["head"]
        or not str(identity["GITHUB_REF"]).startswith("refs/heads/")
        or identity["GITHUB_WORKFLOW_REF"] != f"tensorbee/rdocx/.github/workflows/ci.yml@{identity['GITHUB_REF']}"
        or any(not re.fullmatch(r"[1-9][0-9]*", str(identity[key])) for key in dynamic_identity)
    ):
        raise ValueError("compatibility execution GitHub identity is not admitted")
    root = Path(config["source_root"]).resolve()
    before = compatibility_git_source(root)
    for key in ("head", "tree", "index_sha256"):
        if before[key] != config[key]:
            raise ValueError(f"compatibility execution source changed: {key}")
    if before["files"] != json.loads(compatibility_artifact(config["source_manifest"])):
        raise ValueError("compatibility execution tracked source changed")
    build = json.loads(compatibility_artifact(config["build_provenance"]))
    if build.get("status") != "admitted" or build["current_source_manifest_sha256"] != config["source_manifest"]["sha256"]:
        raise ValueError("compatibility build provenance is not admitted for this source")
    capture = json.loads(compatibility_artifact(build["build_source_capture"]))["files"]
    def provider(name):
        return name.startswith(("crates/", ".cargo/")) or name in ("Cargo.toml", "Cargo.lock", "rust-toolchain.toml", "build.rs")
    current = {name: pin for name, pin in before["files"].items() if provider(name)}
    built = {name: {key: pin[key] for key in ("sha256", "size", "mode")} for name, pin in capture.items() if provider(name)}
    if not current or current != built or set(build["products"]) != {"rdocx", "rpptx"}:
        raise ValueError("compatibility compiled provider graph differs from the actual build")
    for reference in build["products"].values():
        compatibility_artifact(reference)
    probes_sha = hashlib.sha256(json.dumps(config["installed_probes"], sort_keys=True).encode()).hexdigest()
    if build["installed_probes_sha256"] != probes_sha:
        raise ValueError("compatibility installed products do not match admitted build provenance")
    manifest_raw = compatibility_artifact(config["manifest"])
    body_raw = compatibility_artifact(config["body_manifest"])
    if hashlib.sha256(manifest_raw).hexdigest() != COMPATIBILITY_MANIFEST_SHA or (
        hashlib.sha256(body_raw).hexdigest() != COMPATIBILITY_BODY_SHA
    ):
        raise ValueError("compatibility execution requires the original frozen manifests")
    manifest = json.loads(manifest_raw)
    records = manifest["records"]
    ids = {row["id"] for row in records}
    if len(records) != 100 or len(ids) != 100 or any(
        sum(row["area"] == area for row in records) != count for area, count in COMPATIBILITY_COUNTS.items()
    ) or sum(row["family"] == "W" for row in records) != 49 or sum(row["family"] == "P" for row in records) != 51:
        raise ValueError("compatibility execution allocation must remain the frozen 100")
    frozen = Path(config["manifest"]["path"]).parent
    support = frozen / "support.py"
    if hashlib.sha256(support.read_bytes()).hexdigest() != manifest["support_sha256"]:
        raise ValueError("compatibility frozen support changed")
    overlay = json.loads(compatibility_artifact(config["runtime_overlay"]))
    changed = {"fixtures/native-notes.pptx", "fixtures/transition-sound.pptx"}
    if overlay.get("status") != "admitted" or set(overlay["replacements"]) != changed:
        raise ValueError("compatibility runtime overlay must admit exactly the two revised native inputs")
    for ref in overlay["native_producer_evidence"]:
        compatibility_artifact(ref)
    if not overlay["native_producer_evidence"]:
        raise ValueError("compatibility overlay requires actual native producer evidence")
    closure_raw = compatibility_artifact(config["frozen_closure"])
    if hashlib.sha256(closure_raw).hexdigest() != COMPATIBILITY_FREEZE_SHA:
        raise ValueError("compatibility frozen executable closure changed")
    closure = json.loads(closure_raw)["files"]
    closure_by_path = {row["path"]: row for row in closure}
    inputs = {"support.py": support.read_bytes()}
    for row in closure:
        name = row["path"]
        if not name.startswith("fixtures/"):
            continue
        if name in changed:
            revision = overlay["replacements"][name]
            if revision["original_sha256"] != row["sha256"]:
                raise ValueError("compatibility overlay lost original native input identity")
            raw = compatibility_artifact(revision["replacement"])
        else:
            raw = compatibility_artifact({"path": str(frozen / name), "sha256": row["sha256"]})
        inputs[name] = raw
    if not changed.issubset(inputs):
        raise ValueError("compatibility frozen executable fixture closure is incomplete")
    for row in records:
        name = row["script_path"]
        if name != f"scripts/{row['id']}.py" or name not in closure_by_path or closure_by_path[name]["sha256"] != row["script_sha256"]:
            raise ValueError("compatibility script identity changed")
        inputs[name] = compatibility_artifact({"path": str(frozen / name), "sha256": row["script_sha256"]})
    guards = config["runtime_guards"]
    if not guards:
        raise ValueError("compatibility execution requires installed products and external tool guards")
    def authenticate():
        for reference in (config["manifest"], config["body_manifest"], config["frozen_closure"], config["runtime_overlay"]):
            compatibility_artifact(reference)
        if support.read_bytes() != inputs["support.py"]:
            raise ValueError("compatibility original frozen support changed")
        for row in records:
            compatibility_artifact({"path": str(frozen / row["script_path"]), "sha256": row["script_sha256"]})
        for name in inputs:
            if not name.startswith("fixtures/"):
                continue
            reference = overlay["replacements"][name]["replacement"] if name in changed else {
                "path": str(frozen / name), "sha256": closure_by_path[name]["sha256"]}
            compatibility_artifact(reference)
        for reference in overlay["native_producer_evidence"]:
            compatibility_artifact(reference)
        for path, digest in guards.items():
            compatibility_artifact({"path": path, "sha256": digest})
        if compatibility_git_source(root) != before:
            raise ValueError("compatibility source changed during execution")
    if set(config["interpreters"]) != {"W", "P"} or set(config["installed_probes"]) != {"W", "P"}:
        raise ValueError("compatibility execution requires both installed products")
    probe = ("import hashlib,importlib,json,pathlib,sys;"
             "from importlib.metadata import distribution;"
             "m=importlib.import_module(sys.argv[1]);d=distribution(sys.argv[1]);"
             "print(json.dumps({'module':str(pathlib.Path(m.__file__).resolve()),'version':d.version,"
             "'files':{str(pathlib.Path(d.locate_file(f)).resolve()):hashlib.sha256(pathlib.Path(d.locate_file(f)).read_bytes()).hexdigest() "
             "for f in d.files if pathlib.Path(d.locate_file(f)).is_file() and not str(f).endswith(('.pyc','.pyo'))}},sort_keys=True))")
    def installed():
        for family, product in (("W", "rdocx"), ("P", "rpptx")):
            python = config["interpreters"][family]
            if python not in guards:
                raise ValueError("compatibility interpreter is unpinned")
            result = json.loads(subprocess.check_output([python, "-I", "-B", "-c", probe, product], timeout=30))
            if not result["files"] or result["module"] not in result["files"] or result != config["installed_probes"][family] or any(guards.get(path) != digest for path, digest in result["files"].items()):
                raise ValueError("compatibility installed product identity changed")
    authenticate()
    installed()
    output = Path(config["output"]).resolve()
    if output.exists() or output == root or output.is_relative_to(root):
        raise ValueError("compatibility execution needs a fresh external output directory")
    timeout = config["case_timeout_seconds"]
    if type(timeout) is not int or not 1 <= timeout <= 3600:
        raise ValueError("compatibility case timeout must be finite")
    output.mkdir(parents=True)
    execution = output / "frozen-execution"
    execution.mkdir()
    for name, raw in inputs.items():
        target = execution / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(raw)
    def artifact(path):
        return {"path": str(path), "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}
    attempts = []
    error = None
    try:
        for row in records:
            authenticate()
            work = output / "cases" / row["id"]
            work.mkdir(parents=True)
            script = execution / row["script_path"]
            env = {key: value for key, value in os.environ.items() if key not in (
                "PYTHONPATH", "PYTHONHOME", "PYTHONSTARTUP", "PYTEST_ADDOPTS", "DYLD_INSERT_LIBRARIES", "BASH_ENV", "ENV")}
            env.update(config["environment"])
            # Preserve the frozen direct-script assertion traceback. The sole
            # import-path override is this authenticated support directory.
            env.update(PYTHONPATH=str(execution), PYTHONNOUSERSITE="1")
            result = compatibility_process([config["interpreters"][row["family"]], "-s", "-B", str(script)],
                                           work, env, timeout, work / "process.log")
            attempts.append(dict(id=row["id"], family=row["family"], area=row["area"], attempt_number=1,
                                 script_sha256=row["script_sha256"], **result, evidence=[artifact(work / "process.log")]))
            (output / "attempts.json").write_text(json.dumps(attempts, indent=2) + "\n")
            if not result["cleanup"]["leaseReleaseSafe"]:
                raise ValueError("compatibility process cleanup is unsafe")
        authenticate()
        installed()
        for name, raw in inputs.items():
            if (execution / name).read_bytes() != raw:
                raise ValueError("compatibility execution altered a frozen body or runtime input")
    except BaseException as failure:
        error = f"{type(failure).__name__}: {failure}"
        raise
    finally:
        receipt = {"source_kind": "git", "dirty": False, "head": before["head"], "tree": before["tree"],
                   "index_sha256": before["index_sha256"], "attempted": len(attempts), "error": error,
                   "assertions_exit_zero": sum(row["exit_code"] == 0 and not row["timed_out"] for row in attempts),
                   "full_acceptance_percent": None, "qualification_status": "pending",
                   "runtime_overlay": config["runtime_overlay"], "historical_score_relabelled": False}
        receipt["github_identity"] = identity
        receipt["build_provenance"] = config["build_provenance"]
        receipt["missing_non_runtime_history"] = [row for row in closure if not row["path"].startswith(("scripts/", "fixtures/"))
                                                   and not (frozen / row["path"]).is_file()]
        (output / "execution-receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    return receipt


def cmd_compatibility_execute(args: argparse.Namespace) -> int:
    """Keep execution and subsequent independently admitted reporting in one job."""
    try:
        if not 0 <= args.qualification_wait_seconds <= 1800:
            raise ValueError("compatibility qualification wait must be finite")
        config_raw = compatibility_artifact({"path": str(Path(args.config).resolve()), "sha256": args.config_sha})
        config = json.loads(config_raw)
        receipt = compatibility_execute(config)
        output = Path(config["output"])
        deadline = time.monotonic() + args.qualification_wait_seconds
        packet_path = output / "owner-admitted-report-packet.json"
        while not packet_path.exists() and time.monotonic() < deadline:
            time.sleep(1)
        if not packet_path.exists():
            raise ValueError("current CI attempts still await independent file/native qualification")
        if packet_path.stat().st_size > 10_000_000:
            raise ValueError("compatibility qualification packet exceeds the bound")
        packet = json.loads(packet_path.read_text())
        attempts = output / "attempts.json"
        if packet["attempts"] != {"path": str(attempts), "sha256": hashlib.sha256(attempts.read_bytes()).hexdigest()}:
            raise ValueError("compatibility reporting must use this actual CI attempt")
        provenance = json.loads(compatibility_artifact(packet["provenance"]))
        if provenance.get("github_identity") != receipt["github_identity"]:
            raise ValueError("compatibility reporting must bind the actual CI run and attempt")
        if provenance["source_manifest"] != config["source_manifest"] or any(
            provenance[key] != receipt[key] for key in ("head", "tree", "index_sha256", "source_kind", "dirty")
        ):
            raise ValueError("compatibility reporting source differs from this actual CI execution")
        result = compatibility_report(packet)
        (output / "qualified-report.json").write_text(json.dumps(result, indent=2) + "\n")
        # Never upload raw captured output, Office resources or saved documents.
        # The job summary contains only this fixed count/identity projection.
        summary = {key: result[key] for key in ("denominator", "by_area", "documentation_subset", "total",
                                               "full_acceptance_percent", "threshold_met")}
        summary["github_identity"] = receipt["github_identity"]
        raw = json.dumps(summary, indent=2)
        if len(raw.encode()) > 16384:
            raise ValueError("compatibility report summary exceeds the upload bound")
        if any(not re.fullmatch(r"[WP]-(BASE|XML|LIFE|AREA)-[0-9]{2}", case_id)
               for row in [summary["total"], summary["documentation_subset"], *summary["by_area"].values()]
               for case_id in row["unexpected_failure_ids"]):
            raise ValueError("compatibility report contains an unreviewed output identity")
        with Path(os.environ["GITHUB_STEP_SUMMARY"]).open("a") as stream:
            stream.write("## Current unchanged compatibility corpus\n\n```json\n" + raw + "\n```\n")
        print(json.dumps(summary, sort_keys=True))
        return 0 if result["threshold_met"] and not result["total"]["unexpected_failure_ids"] else 1
    except (KeyError, TypeError, ValueError, OSError, subprocess.SubprocessError) as error:
        print(f"compatibility-execute: REFUSED: {type(error).__name__}", file=sys.stderr)
        return 1


def compatibility_named_non_successes(packet: dict, source_sha: str, attempts_raw: bytes, by_id: dict) -> set[str]:
    """Admit only reviewed, candidate-specific frozen assertion conflicts."""
    if packet.get("named_non_successes") is None and packet.get("non_success_admission") is None:
        return set()
    raw = compatibility_artifact(packet["named_non_successes"])
    facts = json.loads(raw)
    admission = json.loads(compatibility_artifact(packet["non_success_admission"]))
    attempts_sha = hashlib.sha256(attempts_raw).hexdigest()
    if admission.get("status") != "admitted" or admission.get("disposition_sha256") != hashlib.sha256(raw).hexdigest() or any(
        value.get("source_manifest_sha256") != source_sha or value.get("attempts_sha256") != attempts_sha
        for value in (facts, admission)
    ):
        raise ValueError("named non-success admission does not bind these facts/source/attempts")
    assertions = {
        "W-AREA-19": (18, "assert any('Gelasio' in x for x in names)"),
        "P-AREA-02": (40, "assert len(updates)==1"),
        "P-AREA-14": (17, "assert nv.get('descr')=='Description' and nv.get('title')=='Title'"),
    }
    rows = facts["cases"]
    if not isinstance(rows, list) or not rows or any(not isinstance(row, dict) for row in rows):
        raise ValueError("named non-successes require explicit reviewed case facts")
    declared = set()
    for row in rows:
        case_id = row["id"]
        if case_id not in assertions or case_id in declared:
            raise ValueError("named non-success identity is unreviewed or duplicated")
        attempt = by_id[case_id]
        cleanup = attempt.get("cleanup", {})
        if attempt["exit_code"] != 1 or attempt["timed_out"] or any(
            cleanup.get(field) is not True for field in ("groupGone", "leaderReaped", "leaseReleaseSafe")
        ) or cleanup.get("signals") != []:
            raise ValueError("named non-success needs an actual safely completed assertion failure")
        reference = row["failure_evidence"]
        if reference not in attempt["evidence"]:
            raise ValueError("named non-success failure log is not this original attempt's evidence")
        text = compatibility_artifact(reference).decode()
        frames = re.findall(r'File "([^"\n]+)", line (\d+), in <module>', text)
        line, assertion = assertions[case_id]
        if len(frames) != 1 or not frames[0][0].endswith(f"/scripts/{case_id}.py") or frames[0][1] != str(line) or (
            not text.startswith("Traceback (most recent call last):\n")
            or f"\n    {assertion}\n" not in text or not text.endswith("\nAssertionError\n")
        ):
            raise ValueError("named non-success did not reach its reviewed frozen assertion")
        for field, kind in (("source_model", "source-model"), ("save_reopen", "save-reopen")):
            compatibility_observations(row[field])
            if any(fact["kind"] != kind for fact in row[field]):
                raise ValueError("named non-success companion observation has the wrong scope")
        declared.add(case_id)
    return declared


def compatibility_report(packet: dict) -> dict:
    """Report first attempts. Qualification is separate from script exit status."""
    manifest_raw = compatibility_artifact(packet["manifest"])
    body_raw = compatibility_artifact(packet["body_manifest"])
    if hashlib.sha256(manifest_raw).hexdigest() != COMPATIBILITY_MANIFEST_SHA:
        raise ValueError("compatibility manifest is not the frozen S91 manifest")
    if hashlib.sha256(body_raw).hexdigest() != COMPATIBILITY_BODY_SHA:
        raise ValueError("compatibility bodies are not the frozen S91 bodies")
    manifest = json.loads(manifest_raw)
    records = manifest["records"]
    ids = {row["id"] for row in records}
    if len(records) != 100 or len(ids) != 100:
        raise ValueError("compatibility denominator must be exactly 100 distinct cases")
    for area, count in COMPATIBILITY_COUNTS.items():
        if sum(row["area"] == area for row in records) != count:
            raise ValueError(f"compatibility allocation changed: {area}")
    if sum(row["family"] == "W" for row in records) != 49 or sum(
        row["family"] == "P" for row in records
    ) != 51:
        raise ValueError("compatibility family allocation must remain 49 Word and 51 PPT")
    attempts_raw = compatibility_artifact(packet["attempts"])
    attempts = json.loads(attempts_raw)
    if len(attempts) != 100 or {row["id"] for row in attempts} != ids:
        raise ValueError("compatibility attempts contain duplicate or missing case identities")
    by_id = {row["id"]: row for row in attempts}
    for record in records:
        attempt = by_id[record["id"]]
        if attempt["attempt_number"] != 1 or type(attempt["attempt_number"]) is not int:
            raise ValueError("compatibility credit is limited to each immutable first attempt")
        if attempt["script_sha256"] != record["script_sha256"]:
            raise ValueError(f"compatibility script changed: {record['id']}")
        if attempt["family"] != record["family"] or attempt["area"] != record["area"]:
            raise ValueError(f"compatibility attempt identity changed: {record['id']}")
        if type(attempt["timed_out"]) is not bool or (
            attempt["exit_code"] is not None and type(attempt["exit_code"]) is not int
        ):
            raise ValueError("compatibility exit status must be an actual integer or null")
        for reference in attempt["evidence"]:
            compatibility_artifact(reference)
        if not attempt["evidence"]:
            raise ValueError(f"compatibility attempt has no retained evidence: {record['id']}")
    provenance = json.loads(compatibility_artifact(packet["provenance"]))
    # Historical dirty captures and actual clean Git executions are distinct.
    if provenance["source_kind"] == "working-copy":
        if provenance["dirty"] is not True:
            raise ValueError("compatibility working-copy source must identify the actual dirty candidate")
    elif provenance["source_kind"] == "git":
        if provenance["dirty"] is not False:
            raise ValueError("compatibility clean Git source cannot claim dirty bytes")
        current = compatibility_git_source(Path(provenance["source_root"]))
        for key in ("head", "tree", "index_sha256"):
            if current[key] != provenance[key]:
                raise ValueError(f"compatibility clean Git identity changed: {key}")
        if current["files"] != json.loads(compatibility_artifact(provenance["source_manifest"])):
            raise ValueError("compatibility clean Git tracked bytes changed")
    else:
        raise ValueError("compatibility source kind must be working-copy or git")
    for key, size in (("base", 40), ("head", 40), ("index_sha256", 64)):
        if not re.fullmatch(rf"[0-9a-f]{{{size}}}", provenance[key]):
            raise ValueError(f"compatibility source identity is invalid: {key}")
    for reference in provenance["evidence"]:
        compatibility_artifact(reference)
    source_raw = compatibility_artifact(provenance["source_manifest"])
    source_sha = hashlib.sha256(source_raw).hexdigest()
    if not provenance["evidence"] or not json.loads(source_raw):
        raise ValueError("compatibility source/build/install evidence is missing")
    dependencies = packet["dependencies"]
    required = {"F-X211a", "F-X211b", "F-X211d", "F-X215", "F-X217", "F-X218", "F-X220"}
    required.update(f"F-X{number}" for number in range(200, 211))
    if set(dependencies) != required:
        raise ValueError("compatibility provider dependency identities are incomplete")
    for reference in dependencies.values():
        compatibility_artifact(reference)
    divergences = COMPATIBILITY_DIVERGENCES | compatibility_named_non_successes(packet, source_sha, attempts_raw, by_id)

    # Until a separate reviewed fact packet is admitted, no correct-file score
    # follows from assertion exits, hash gates or package inventories.
    qualified = set()
    full_metric = None
    declared = {row["id"]: compatibility_obligations(row) for row in records if row.get("additional_native_obligations")}
    native_inventory = [
        {"id": case_id, "obligation": text, "obligation_sha256": hashlib.sha256(text.encode()).hexdigest(), "status": "pending"}
        for case_id, texts in declared.items() for text in texts
    ]
    if packet.get("qualification") is not None:
        facts_raw = compatibility_artifact(packet["qualification"])
        facts = json.loads(facts_raw)
        admission = json.loads(compatibility_artifact(packet["qualification_admission"]))
        if admission.get("status") != "admitted":
            raise ValueError("compatibility qualification is not admitted")
        if admission["qualification_sha256"] != hashlib.sha256(facts_raw).hexdigest() or (
            admission["source_manifest_sha256"] != source_sha
            or admission["attempts_sha256"] != hashlib.sha256(attempts_raw).hexdigest()
        ):
            raise ValueError("compatibility admission does not bind these facts/source/attempts")
        if facts["source_manifest_sha256"] != source_sha or (
            facts["attempts_sha256"] != admission["attempts_sha256"]
        ):
            raise ValueError("compatibility facts belong to a different candidate")
        fact_rows = facts["cases"]
        if len(fact_rows) != 100 or {row["id"] for row in fact_rows} != ids:
            raise ValueError("compatibility facts must evaluate all 100 unchanged cases")
        for row in fact_rows:
            if row["status"] not in ("qualified", "unqualified", "pending"):
                raise ValueError("compatibility fact status is invalid")
            for reference in row["evidence"]:
                compatibility_artifact(reference)
            if row["status"] == "qualified":
                proof_fields = ("source_model", "save_reopen")
                if not row["evidence"]:
                    raise ValueError("correct-file credit needs actual source/model and save/reopen facts")
                for field in proof_fields:
                    compatibility_observations(row[field])
                attempt = by_id[row["id"]]
                if attempt["exit_code"] != 0 or attempt["timed_out"] or row["id"] in divergences:
                    raise ValueError("a failed or divergent first attempt cannot earn correct-file credit")
                qualified.add(row["id"])
        # Admission must enumerate the defining owner, safety and native/render
        # authorities. Their concrete proofs remain subject to independent review.
        hard_gates = facts["hard_gates"]
        required_gates = {
            row["id"] for row in records
            if row["area"] in ("XML", "LIFE") or row.get("additional_native_obligations")
        }
        required_gates.add("P-AREA-10")
        if not required_gates.issubset(hard_gates):
            raise ValueError("all current-owner XML, lifetime and declared native/render hard gates are required")
        for gate in hard_gates.values():
            if gate["status"] not in ("qualified", "unqualified", "pending"):
                raise ValueError("compatibility hard gate status is invalid")
            for reference in gate["evidence"]:
                compatibility_artifact(reference)
            if gate["status"] == "qualified":
                if not gate["evidence"]:
                    raise ValueError("qualified hard gates require retained actual observations/evidence")
                compatibility_observations(gate["observations"])
        native_facts = facts["native_obligations"]
        if set(native_facts) != set(declared):
            raise ValueError("qualification must bind every frozen native obligation")
        native_ready = True
        for case_id, texts in declared.items():
            proofs = native_facts[case_id]
            if not isinstance(proofs, list) or len(proofs) != len(texts):
                raise ValueError(f"native obligation proofs are missing: {case_id}")
            for text, proof in zip(texts, proofs):
                digest = hashlib.sha256(text.encode()).hexdigest()
                if proof["obligation"] != text or proof["obligation_sha256"] != digest:
                    raise ValueError(f"native obligation identity changed: {case_id}")
                if proof["status"] not in ("qualified", "unqualified", "pending"):
                    raise ValueError("native obligation status is invalid")
                if proof["status"] == "qualified":
                    compatibility_observations(proof["observations"])
                else:
                    native_ready = False
        native_inventory = [dict(id=case_id, obligation=proof["obligation"],
                                 obligation_sha256=proof["obligation_sha256"], status=proof["status"])
                            for case_id, proofs in native_facts.items() for proof in proofs]
        complete = all(row["status"] != "pending" for row in fact_rows)
        gates_ready = all(gate["status"] == "qualified" for gate in hard_gates.values())
        if complete and gates_ready and native_ready and admission["status"] == "admitted":
            full_metric = len(qualified)
    def counts(selected: list[dict]) -> dict:
        selected_ids = {row["id"] for row in selected}
        exits = {row["id"] for row in selected if row["exit_code"] == 0 and not row["timed_out"]}
        return {"attempted": len(selected), "executed": sum(row["exit_code"] is not None for row in selected),
                "assertions_exit_zero": len(exits), "correct_file": len(selected_ids & qualified),
                "deliberate_divergence": len(selected_ids & divergences),
                "unexpected_failure_ids": sorted(selected_ids - exits - divergences)}
    documentation = {f"W-BASE-{i:02}" for i in range(1, 18)} | {f"P-BASE-{i:02}" for i in range(1, 8)}
    return {"denominator": 100, "body_manifest_sha256": COMPATIBILITY_BODY_SHA,
            "source_manifest_sha256": source_sha, "provenance": provenance,
            "by_area": {area: counts([row for row in attempts if row["area"] == area]) for area in COMPATIBILITY_COUNTS},
            "documentation_subset": counts([row for row in attempts if row["id"] in documentation]),
            "native_obligations": native_inventory,
            "total": counts(attempts), "full_acceptance_percent": full_metric,
            "threshold_met": full_metric is not None and full_metric >= 90}


def cmd_compatibility_report(args: argparse.Namespace) -> int:
    try:
        result = compatibility_report(json.loads(Path(args.packet).read_text()))
    except (KeyError, TypeError, ValueError, OSError) as error:
        print(f"compatibility-report: REFUSED: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result, indent=2, sort_keys=True))
    if args.check and (
        not result["threshold_met"] or result["total"]["unexpected_failure_ids"]
    ):
        print("compatibility-report: qualification incomplete or acceptance failed", file=sys.stderr)
        return 1
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("init"); p.add_argument("sprint"); p.add_argument("--force", action="store_true"); p.add_argument("--resume", action="store_true"); p.add_argument("--max-review-passes", type=int, default=MAX_REVIEW_PASSES); p.add_argument("--max-workers", type=int, default=None); p.add_argument("--reviewer", help="required identity of the independent sprint reviewer"); p.set_defaults(fn=cmd_init)
    p = sub.add_parser("status"); p.add_argument("sprint", nargs="?"); p.add_argument("--workers", action="store_true", help="show branch, worktree and handoff per feature"); p.set_defaults(fn=cmd_status)
    p = sub.add_parser("set-phase"); p.add_argument("sprint"); p.add_argument("phase"); p.set_defaults(fn=cmd_set_phase)
    p = sub.add_parser("mark-feature"); p.add_argument("sprint"); p.add_argument("fid"); p.add_argument("state"); p.add_argument("--owner"); p.add_argument("--clear-owner", action="store_true"); p.add_argument("--wave", type=int); p.add_argument("--branch"); p.add_argument("--worktree"); p.add_argument("--base"); p.add_argument("--head"); p.add_argument("--handoff"); p.add_argument("--integration-commit"); p.set_defaults(fn=cmd_mark_feature)
    p = sub.add_parser("record-review"); p.add_argument("sprint"); p.add_argument("passno", type=int); p.add_argument("--blocking", type=int, required=True); p.add_argument("--should-fix", type=int, default=0); p.add_argument("--nice-to-have", type=int, default=0); p.add_argument("--extend", action="store_true"); p.add_argument("--reviewer", help="actual reviewer identity, not the implementing runner"); p.set_defaults(fn=cmd_record_review)
    p = sub.add_parser("record-verification"); p.add_argument("sprint"); p.add_argument("--scope", choices=["fast", "scoped", "full", "metadata"], required=True); p.add_argument("--reuse-from"); p.add_argument("--passed", action="store_true"); p.add_argument("--harness", default="unchecked"); p.set_defaults(fn=cmd_record_verification)
    p = sub.add_parser("validate-handoff"); p.add_argument("path"); p.add_argument("--fid", required=True); p.set_defaults(fn=cmd_validate_handoff)
    p = sub.add_parser("close-preflight"); p.add_argument("sprint"); p.set_defaults(fn=cmd_close_preflight)
    p = sub.add_parser("release-notes"); p.add_argument("tag"); mode = p.add_mutually_exclusive_group(required=True); mode.add_argument("--check", action="store_true"); mode.add_argument("--render", action="store_true"); p.set_defaults(fn=cmd_release_notes)
    p = sub.add_parser("release-semver", allow_abbrev=False)
    p.add_argument("scope")
    for name in ("target", "tool", "output"):
        p.add_argument("--" + name, action="append")
    p.add_argument("--reuse-evidence", action="append", default=[])
    p.add_argument("--check-evidence", nargs="+", action="append")
    p.add_argument("--refresh-registry", action="count", default=0)
    p.add_argument("--release-candidate", action="count", default=0)
    p.set_defaults(fn=cmd_release_semver)
    p = sub.add_parser("unified-release-family"); p.add_argument("tag"); p.set_defaults(fn=cmd_unified_release_family)
    p = sub.add_parser("published-unified-release"); p.add_argument("tag"); p.add_argument("directory"); p.set_defaults(fn=cmd_published_unified_release)
    p = sub.add_parser("python-release-artifacts"); p.add_argument("tag"); p.add_argument("directory"); p.set_defaults(fn=cmd_python_release_artifacts)
    p = sub.add_parser("compatibility-report"); p.add_argument("packet"); mode = p.add_mutually_exclusive_group(required=True); mode.add_argument("--check", action="store_true"); mode.add_argument("--render", action="store_true"); p.set_defaults(fn=cmd_compatibility_report)
    p = sub.add_parser("compatibility-execute")
    p.add_argument("config")
    p.add_argument("--config-sha", required=True)
    p.add_argument("--qualification-wait-seconds", type=int, default=1800)
    p.set_defaults(fn=cmd_compatibility_execute)

    args = ap.parse_args()
    if getattr(args, "fid", None) and not FID_RE.match(args.fid):
        die(f"{args.fid} is not a valid F-ID")
    if hasattr(args, "sprint") and args.sprint and not SPRINT_ID_RE.match(args.sprint):
        die(f"{args.sprint} is not a valid sprint id, expected SNN")
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())
