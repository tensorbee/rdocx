#!/usr/bin/env python3
"""Validate every workspace README and compile its Rust examples."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
from pathlib import Path
import re
import subprocess
import sys
import tarfile
import tomllib


REPO_ROOT = Path(__file__).resolve().parent.parent
RUST_FENCE = re.compile(r"^```rust(?P<attributes>[^\n]*)$", re.MULTILINE)
EXAMPLE_FENCE = re.compile(
    r"^```(?:rust[^\n]*|toml|python|javascript|sh|text)$",
    re.MULTILINE,
)
WORKSPACE_PACKAGE_COUNT = 27
PUBLISHABLE_PACKAGE_COUNT = 22
LOCAL_PATCHES = (
    ("oxml-core", "crates/oxml-core"),
    ("oxml-drawing", "crates/oxml-drawing"),
    ("oxml-layout", "crates/oxml-layout"),
    ("oxml-media", "crates/oxml-media"),
    ("oxml-opc", "crates/oxml-opc"),
    ("oxml-pdf", "crates/oxml-pdf"),
    ("oxml-sml", "crates/oxml-sml"),
    ("oxml-cli-support", "crates/oxml-cli-support"),
    ("oxml-chart", "crates/oxml-chart"),
    ("rdocx", "crates/rdocx"),
    ("rdocx-cli", "crates/rdocx-cli"),
    ("rdocx-html", "crates/rdocx-html"),
    ("rdocx-layout", "crates/rdocx-layout"),
    ("rdocx-opc", "crates/rdocx-opc"),
    ("rdocx-oxml", "crates/rdocx-oxml"),
    ("rdocx-pdf", "crates/rdocx-pdf"),
    ("rpptx", "crates/rpptx"),
    ("rpptx-cli", "crates/rpptx-cli"),
    ("rpptx-chart", "crates/rpptx-chart"),
    ("rpptx-layout", "crates/rpptx-layout"),
    ("rpptx-oxml", "crates/rpptx-oxml"),
    ("rpptx-render", "crates/rpptx-render"),
)


@dataclass(frozen=True)
class ReadmeCase:
    package: str
    crate_name: str
    readme: Path
    expected_rust_fences: int
    companion_crates: tuple[tuple[str, str], ...] = ()


README_CASES = (
    ReadmeCase("rdocx", "rdocx", REPO_ROOT / "README.md", 7),
    ReadmeCase(
        "rdocx-opc",
        "rdocx_opc",
        REPO_ROOT / "crates/rdocx-opc/README.md",
        1,
    ),
    ReadmeCase(
        "rdocx-oxml",
        "rdocx_oxml",
        REPO_ROOT / "crates/rdocx-oxml/README.md",
        1,
    ),
    ReadmeCase(
        "rdocx-layout",
        "rdocx_layout",
        REPO_ROOT / "crates/rdocx-layout/README.md",
        1,
    ),
    ReadmeCase(
        "rdocx-html",
        "rdocx_html",
        REPO_ROOT / "crates/rdocx-html/README.md",
        1,
    ),
    ReadmeCase(
        "rdocx-pdf",
        "rdocx_pdf",
        REPO_ROOT / "crates/rdocx-pdf/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-cli-support",
        "oxml_cli_support",
        REPO_ROOT / "crates/oxml-cli-support/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-core", "oxml_core", REPO_ROOT / "crates/oxml-core/README.md", 1
    ),
    ReadmeCase(
        "oxml-drawing",
        "oxml_drawing",
        REPO_ROOT / "crates/oxml-drawing/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-layout",
        "oxml_layout",
        REPO_ROOT / "crates/oxml-layout/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-media", "oxml_media", REPO_ROOT / "crates/oxml-media/README.md", 1
    ),
    ReadmeCase(
        "oxml-opc", "oxml_opc", REPO_ROOT / "crates/oxml-opc/README.md", 1
    ),
    ReadmeCase(
        "oxml-pdf",
        "oxml_pdf",
        REPO_ROOT / "crates/oxml-pdf/README.md",
        1,
        (("oxml-layout", "oxml_layout"),),
    ),
    ReadmeCase(
        "oxml-py-support",
        "oxml_py_support",
        REPO_ROOT / "crates/oxml-py-support/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-sml",
        "oxml_sml",
        REPO_ROOT / "crates/oxml-sml/README.md",
        1,
    ),
    ReadmeCase(
        "oxml-chart",
        "oxml_chart",
        REPO_ROOT / "crates/oxml-chart/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx",
        "rpptx",
        REPO_ROOT / "crates/rpptx/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx-chart",
        "rpptx_chart",
        REPO_ROOT / "crates/rpptx-chart/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx-layout",
        "rpptx_layout",
        REPO_ROOT / "crates/rpptx-layout/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx-oxml",
        "rpptx_oxml",
        REPO_ROOT / "crates/rpptx-oxml/README.md",
        1,
    ),
    ReadmeCase(
        "rpptx-render",
        "rpptx_render",
        REPO_ROOT / "crates/rpptx-render/README.md",
        1,
    ),
)

README_REQUIRED_TEXT = {
    REPO_ROOT / "README.md": (
        'rdocx = "0.13.1"',
        'rdocx = { version = "0.13.1", default-features = false }',
        "rdocx convert report.docx --to pdf -o report.pdf",
        "rdocx convert report.docx --to html -o report.html",
        "rdocx convert report.docx --to md -o report.md",
        'rdocx replace report.docx --placeholder "Draft" --value "Final" -o final.docx',
    ),
    REPO_ROOT / "crates/rdocx-cli/README.md": (
        "cargo install rdocx-cli --version '^0.13.1'",
        "rdocx convert report.docx --to pdf -o report.pdf",
    ),
    REPO_ROOT / "crates/rdocx-html/README.md": ('rdocx-html = "0.13.1"',),
    REPO_ROOT / "crates/rdocx-layout/README.md": ('rdocx-layout = "0.13.1"',),
    REPO_ROOT / "crates/rdocx-opc/README.md": (
        'rdocx-opc = "0.13.1"',
        "use rdocx_opc::OpcPackage;",
    ),
    REPO_ROOT / "crates/rdocx-oxml/README.md": ('rdocx-oxml = "0.13.1"',),
    REPO_ROOT / "crates/rdocx-pdf/README.md": (
        'rdocx-pdf = "0.13.1"',
        "use rdocx_pdf::render_to_pdf;",
    ),
    REPO_ROOT / "crates/oxml-cli-support/README.md": (
        'oxml_cli_support::parse_range("2,4-6")?',
    ),
    REPO_ROOT / "crates/oxml-core/README.md": ("Length::inches(8.5)",),
    REPO_ROOT / "crates/oxml-drawing/README.md": ("Fill::from_xml",),
    REPO_ROOT / "crates/oxml-layout/README.md": ('Color::from_hex("3366CC")',),
    REPO_ROOT / "crates/oxml-media/README.md": ("resolve(b\"\\x89PNG",),
    REPO_ROOT / "crates/oxml-opc/README.md": ("ContentTypes::from_xml",),
    REPO_ROOT / "crates/oxml-pdf/README.md": ("render_to_pdf(&layout)",),
    REPO_ROOT / "crates/oxml-py-support/README.md": (
        "emu_from_inches(8.5)",
    ),
    REPO_ROOT / "crates/rdocx-py/README.md": (
        "doc.add_paragraph(\"Hello from Python\")",
        'doc.save("hello.docx")',
    ),
    REPO_ROOT / "crates/rdocx-wasm/README.md": (
        'from "@tensorbee/rdocx-wasm"',
        "doc.toDocxBytes()",
    ),
    REPO_ROOT / "crates/rpptx/README.md": (
        "use rpptx::Presentation;",
        "Presentation::new()?",
    ),
    REPO_ROOT / "crates/oxml-chart/README.md": ("AxisId::new(10_000_001)?",),
    REPO_ROOT / "crates/rpptx-chart/README.md": ("AxisId::new(10_000_001)?",),
    REPO_ROOT / "crates/rpptx-cli/README.md": (
        "cargo install rpptx-cli --version '^0.11.0'",
        "rpptx convert deck.pptx --to pdf -o deck.pdf",
    ),
    REPO_ROOT / "crates/rpptx-py/README.md": (
        'Presentation("deck.pptx")',
        "len(presentation.slides)",
    ),
    REPO_ROOT / "crates/rpptx-layout/README.md": ("ScopedMediaIds::default()",),
    REPO_ROOT / "crates/rpptx-oxml/README.md": ("CT_Presentation::from_xml",),
    REPO_ROOT / "crates/rpptx-render/README.md": ("RelScopes::default()",),
    REPO_ROOT / "crates/rpptx-wasm/README.md": (
        'from "@tensorbee/rpptx-wasm"',
        "deck.toBytes()",
    ),
}


def validate_fences(readme: Path, expected: int) -> bool:
    text = readme.read_text(encoding="utf-8")
    attributes = RUST_FENCE.findall(text)
    if len(attributes) != expected or any(
        attribute != ",no_run" for attribute in attributes
    ):
        print(
            f"README doctest error: expected {expected} "
            f"exact rust,no_run fences, found {len(attributes)} with "
            f"attributes {attributes!r}",
            file=sys.stderr,
        )
        return False
    return True


def cargo_metadata() -> dict[str, object] | None:
    result = subprocess.run(
        ["cargo", "metadata", "--no-deps", "--format-version", "1"],
        cwd=REPO_ROOT,
        stdout=subprocess.PIPE,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        return None
    return json.loads(result.stdout)


def package_readme(package: dict[str, object]) -> Path | None:
    value = package.get("readme")
    manifest = package.get("manifest_path")
    if not isinstance(value, str) or not isinstance(manifest, str):
        return None
    return (Path(manifest).parent / value).resolve()


def validate_local_patches(packages: list[object]) -> bool:
    expected: set[tuple[str, str]] = set()
    for package in packages:
        if not isinstance(package, dict) or package.get("publish") == []:
            continue
        name = package.get("name")
        manifest = package.get("manifest_path")
        if not isinstance(name, str) or not isinstance(manifest, str):
            print(
                "README doctest error: invalid publishable package metadata",
                file=sys.stderr,
            )
            return False
        try:
            package_path = Path(manifest).parent.resolve().relative_to(REPO_ROOT)
        except ValueError:
            print(
                f"README doctest error: {name} is outside the repository",
                file=sys.stderr,
            )
            return False
        expected.add((name, package_path.as_posix()))

    actual = set(LOCAL_PATCHES)
    if len(actual) != len(LOCAL_PATCHES) or actual != expected:
        missing = sorted(expected - actual)
        unexpected = sorted(actual - expected)
        print(
            "README doctest error: local patches differ from publishable "
            f"metadata, missing={missing!r}, unexpected={unexpected!r}",
            file=sys.stderr,
        )
        return False
    return True


def validate_package_archive(package: dict[str, object], readme: Path) -> bool:
    name = package["name"]
    version = package.get("version")
    if not isinstance(name, str) or not isinstance(version, str):
        print("README doctest error: invalid package identity", file=sys.stderr)
        return False
    command = [
        "cargo",
        "package",
        "--locked",
        "--allow-dirty",
        "--no-verify",
        "-p",
        name,
    ]
    for patch_name, patch_path in LOCAL_PATCHES:
        command.extend(
            ["--config", f'patch.crates-io.{patch_name}.path="{patch_path}"']
        )
    result = subprocess.run(
        command,
        cwd=REPO_ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr, end="")
        return False
    archive = REPO_ROOT / f"target/package/{name}-{version}.crate"
    if not archive.is_file():
        print(
            f"README doctest error: missing generated archive {archive}",
            file=sys.stderr,
        )
        return False
    with tarfile.open(archive, "r:gz") as package_archive:
        readmes = [
            member
            for member in package_archive.getmembers()
            if Path(member.name).name == "README.md"
        ]
        if len(readmes) != 1:
            print(
                f"README doctest error: {name} archive has {len(readmes)} "
                "README files",
                file=sys.stderr,
            )
            return False
        packaged = package_archive.extractfile(readmes[0])
        if packaged is None or packaged.read() != readme.read_bytes():
            print(
                f"README doctest error: {name} archive README differs from "
                f"{readme}",
                file=sys.stderr,
            )
            return False
    return True


def validate_inventory() -> bool:
    metadata = cargo_metadata()
    if metadata is None:
        return False
    packages = metadata.get("packages")
    if not isinstance(packages, list) or len(packages) != WORKSPACE_PACKAGE_COUNT:
        observed = len(packages) if isinstance(packages, list) else "invalid"
        print(
            f"README doctest error: expected {WORKSPACE_PACKAGE_COUNT} workspace "
            f"packages, found {observed}",
            file=sys.stderr,
        )
        return False

    valid = True
    readme_paths: set[Path] = set()
    publishable: list[tuple[dict[str, object], Path]] = []
    for package in packages:
        if not isinstance(package, dict) or not isinstance(package.get("name"), str):
            print("README doctest error: invalid Cargo package metadata", file=sys.stderr)
            valid = False
            continue
        name = package["name"]
        readme = package_readme(package)
        manifest_value = package.get("manifest_path")
        if not isinstance(manifest_value, str):
            print(
                f"README doctest error: {name} lacks a manifest path",
                file=sys.stderr,
            )
            valid = False
            continue
        manifest = Path(manifest_value)
        manifest_data = tomllib.loads(manifest.read_text(encoding="utf-8"))
        declared_readme = manifest_data.get("package", {}).get("readme")
        if not isinstance(declared_readme, str):
            print(
                f"README doctest error: {manifest} lacks an explicit package.readme",
                file=sys.stderr,
            )
            valid = False
        if readme is None or not readme.is_file():
            print(
                f"README doctest error: {name} does not declare an existing README",
                file=sys.stderr,
            )
            valid = False
            continue
        readme_paths.add(readme)
        text = readme.read_text(encoding="utf-8")
        if not text.startswith(f"# {name}\n"):
            print(
                f"README doctest error: {readme} must start with '# {name}'",
                file=sys.stderr,
            )
            valid = False
        if name != "rdocx":
            for heading in ("## Use it when", "## Relationship", "## Example"):
                if heading not in text:
                    print(
                        f"README doctest error: {readme} lacks {heading!r}",
                        file=sys.stderr,
                    )
                    valid = False
        if not EXAMPLE_FENCE.search(text):
            print(
                f"README doctest error: {readme} lacks a supported example fence",
                file=sys.stderr,
            )
            valid = False
        if package.get("publish") != []:
            publishable.append((package, readme))

    if len(readme_paths) != WORKSPACE_PACKAGE_COUNT:
        print(
            f"README doctest error: expected {WORKSPACE_PACKAGE_COUNT} distinct "
            f"README sources, found {len(readme_paths)}",
            file=sys.stderr,
        )
        valid = False
    if len(publishable) != PUBLISHABLE_PACKAGE_COUNT:
        print(
            f"README doctest error: expected {PUBLISHABLE_PACKAGE_COUNT} "
            f"publishable packages, found {len(publishable)}",
            file=sys.stderr,
        )
        valid = False
    if not validate_local_patches(packages):
        valid = False

    for readme, required_items in README_REQUIRED_TEXT.items():
        text = readme.read_text(encoding="utf-8")
        for required in required_items:
            if required not in text:
                print(
                    f"README doctest error: {readme} does not contain "
                    f"{required!r}",
                    file=sys.stderr,
                )
                valid = False
    if valid:
        for package, readme in sorted(
            publishable, key=lambda item: str(item[0]["name"])
        ):
            if not validate_package_archive(package, readme):
                valid = False
    if valid:
        print(
            f"readme_doctests: {WORKSPACE_PACKAGE_COUNT} distinct workspace "
            f"READMEs and {PUBLISHABLE_PACKAGE_COUNT} publishable package "
            "inventories validated"
        )
    return valid


def build_rlibs(package: str, crate_names: tuple[str, ...]) -> dict[str, Path] | None:
    command = [
        "cargo",
        "build",
        "--locked",
        "-p",
        package,
        "--message-format=json-render-diagnostics",
    ]
    result = subprocess.run(
        command,
        cwd=REPO_ROOT,
        stdout=subprocess.PIPE,
        text=True,
        check=False,
    )
    artifacts: dict[str, set[Path]] = {name: set() for name in crate_names}
    for line in result.stdout.splitlines():
        try:
            message = json.loads(line)
        except json.JSONDecodeError:
            continue
        if message.get("reason") == "compiler-message":
            rendered = message.get("message", {}).get("rendered")
            if rendered:
                print(rendered, file=sys.stderr, end="")
        if message.get("reason") != "compiler-artifact":
            continue
        target = message.get("target", {})
        crate_name = target.get("name")
        if crate_name not in artifacts or "lib" not in target.get("crate_types", []):
            continue
        artifacts[crate_name].update(
            Path(filename).resolve()
            for filename in message.get("filenames", [])
            if filename.endswith(".rlib")
        )
    if result.returncode != 0:
        return None
    for crate_name, crate_artifacts in artifacts.items():
        if len(crate_artifacts) != 1:
            print(
                f"README doctest error: expected one {crate_name} rlib from "
                f"the {package} build, found "
                f"{sorted(str(path) for path in crate_artifacts)!r}",
                file=sys.stderr,
            )
            return None
    return {name: paths.pop() for name, paths in artifacts.items()}


def compile_readme(case: ReadmeCase) -> bool:
    if not validate_fences(case.readme, case.expected_rust_fences):
        return False

    crate_names = (case.crate_name,) + tuple(
        crate_name for _, crate_name in case.companion_crates
    )
    rlibs = build_rlibs(case.package, crate_names)
    if rlibs is None:
        return False
    resolved_rlibs = list(rlibs.items())
    rlib = rlibs[case.crate_name]
    dependency_dir = rlib.parent / "deps"
    if not dependency_dir.is_dir():
        dependency_dir = rlib.parent
    command = [
        "rustdoc",
        "--test",
        str(case.readme),
        "--crate-name",
        f"{case.crate_name}_readme",
        "--edition=2024",
        "-Dwarnings",
        "-L",
        f"dependency={dependency_dir}",
    ]
    for crate_name, crate_rlib in resolved_rlibs:
        command.extend(("--extern", f"{crate_name}={crate_rlib}"))
    result = subprocess.run(
        command,
        cwd=REPO_ROOT,
        check=False,
    )
    if result.returncode == 0:
        print(
            f"readme_doctests: {case.expected_rust_fences} Rust examples "
            f"compiled from {case.readme}"
        )
    return result.returncode == 0


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("readme", nargs="?", type=Path)
    args = parser.parse_args()
    if args.readme is not None:
        case = ReadmeCase("rdocx", "rdocx", args.readme.resolve(), 7)
        return 0 if compile_readme(case) else 1

    if not validate_inventory():
        return 1
    return 0 if all(compile_readme(case) for case in README_CASES) else 1


if __name__ == "__main__":
    raise SystemExit(main())
