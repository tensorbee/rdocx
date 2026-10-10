"""Run the first-draft corpus: scripts written the way an agent writes them
from memory of python-docx, each followed by a check of the saved file.

Each script in ``first_drafts/`` declares its area and one status line:

- ``# expected: pass``: runs unmodified and writes a correct file. A failure
  fails this test.
- ``# needs: #NNN``: passes once the named issue or pull request lands. A
  pass is reported so that the mark is updated.
- ``# divergence: <text>``: a deliberate difference from python-docx. The
  script must fail with an error whose message contains ``<text>``, the
  rdocx way it names.

The code before the ``# --- check`` line is the script. The code after it
checks the saved file, with ``part(path, name)``, ``names(path)`` and ``re``.
The success rate per area and overall is printed on every run.
"""

import re
import struct
import warnings
import zipfile
import zlib
from pathlib import Path

CORPUS = Path(__file__).parent / "first_drafts"
CHECK_MARKER = "# --- check"


def _png(width=2, height=2):
    def chunk(kind, data):
        return (
            struct.pack(">I", len(data))
            + kind
            + data
            + struct.pack(">I", zlib.crc32(kind + data))
        )

    rows = b"".join(b"\x00" + b"\x20\x80\xe0" * width for _ in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(rows))
        + chunk(b"IEND", b"")
    )


def _fixtures(root):
    import rdocx

    (root / "logo.png").write_bytes(_png())
    existing = rdocx.Document()
    existing.add_paragraph("existing")
    existing.save(root / "existing-document-file.docx")
    template = rdocx.Document()
    template.add_paragraph("Dear {{name}},")
    template.add_paragraph("Date: {{date}}")
    template.add_table(rows=3, cols=2)
    for row, values in enumerate(
        (("Metric", "Value"), ("Revenue", "{{value}}"), ("Margin", "{{value}}"))
    ):
        for column, value in enumerate(values):
            template.tables[0].cell(row, column).text = value
    template.add_paragraph("Regards")
    template.save(root / "template.docx")


def part(path, name="word/document.xml"):
    with zipfile.ZipFile(path) as package:
        return package.read(name).decode("utf-8")


def names(path):
    with zipfile.ZipFile(path) as package:
        return package.namelist()


def parse_script(path):
    source = path.read_text()
    header = {}
    for line in source.splitlines():
        match = re.match(r"# (area|expected|needs|divergence): (.+)", line)
        if match:
            header.setdefault(match.group(1), match.group(2).strip())
    body, _, check = source.partition(CHECK_MARKER)
    statuses = [key for key in ("expected", "needs", "divergence") if key in header]
    assert "area" in header and len(statuses) == 1, f"{path.name}: bad header"
    return header, body, check


def run_script(path, root):
    """Run one script in `root` and return None on success or the failure."""
    header, body, check = parse_script(path)
    namespace = {"__name__": "__main__", "__file__": str(path)}
    try:
        exec(compile(body, str(path), "exec"), namespace)
    except Exception as error:  # noqa: BLE001 - every failure is recorded
        return f"script: {type(error).__name__}: {error}".splitlines()[0]
    helpers = {"part": part, "names": names, "re": re}
    try:
        exec(compile(check, f"{path} (check)", "exec"), {**namespace, **helpers})
    except Exception as error:  # noqa: BLE001
        return f"check: {type(error).__name__}: {error}".splitlines()[0][:200]
    return None


def run_corpus(tmp_path, monkeypatch, fixtures):
    scripts = sorted(CORPUS.glob("*.py"))
    results = []
    for path in scripts:
        root = tmp_path / path.stem
        root.mkdir()
        monkeypatch.chdir(root)
        fixtures(root)
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            failure = run_script(path, root)
        header, _, _ = parse_script(path)
        results.append((path.stem, header, failure))
    return results


def report(results, capsys):
    areas = {}
    for _, header, failure in results:
        passed, total = areas.get(header["area"], (0, 0))
        areas[header["area"]] = (passed + (failure is None), total + 1)
    lines = ["", "First-draft corpus success rate (runs unmodified, correct file):"]
    for area, (passed, total) in sorted(areas.items()):
        lines.append(f"  {area:<16} {passed:>3}/{total:<3} {100 * passed / total:5.1f}%")
    passed = sum(failure is None for _, _, failure in results)
    lines.append(f"  {'overall':<16} {passed:>3}/{len(results):<3} {100 * passed / len(results):5.1f}%")
    for name, header, failure in results:
        if failure is not None:
            mark = header.get("needs") or (
                "divergence" if "divergence" in header else "UNEXPECTED"
            )
            lines.append(f"  - {name} [{mark}] {failure}")
    with capsys.disabled():
        print("\n".join(lines))


def judge(results):
    unexpected, stale_marks, silent_divergences = [], [], []
    for name, header, failure in results:
        if "expected" in header and failure is not None:
            unexpected.append(f"{name}: {failure}")
        elif "needs" in header and failure is None:
            stale_marks.append(f"{name} (marked needs {header['needs']})")
        elif "divergence" in header and (
            failure is None or header["divergence"] not in failure
        ):
            silent_divergences.append(f"{name}: {failure}")
    for name in stale_marks:
        warnings.warn(f"first draft now passes, mark it expected: pass: {name}")
    assert not unexpected, "scripts expected to pass failed:\n" + "\n".join(unexpected)
    assert not silent_divergences, (
        "deliberate divergences must fail naming the rdocx way:\n"
        + "\n".join(silent_divergences)
    )


def test_first_draft_corpus(tmp_path, monkeypatch, capsys):
    results = run_corpus(tmp_path, monkeypatch, _fixtures)
    assert len(results) >= 50
    report(results, capsys)
    judge(results)
