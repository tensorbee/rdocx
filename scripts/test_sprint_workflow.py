from __future__ import annotations

import argparse
import contextlib
import fnmatch
import hashlib
import io
import json
import os
import re
import subprocess
import tarfile
import tempfile
import tomllib
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts import fetch_docx_corpus
from scripts import readme_doctests
from scripts import install_pinned_libreoffice
from scripts import install_pinned_pandoc
from scripts import install_pinned_poppler
from scripts import sprint_workflow as workflow


class SprintWorkflowTests(unittest.TestCase):
    LARGE_DOCUMENT_TEST = (
        "a_thousand_page_document_paginates_and_renders_within_the_declared_limits"
    )

    def yaml_block(self, source: str, header: str) -> str:
        lines = source.splitlines()
        matches = [index for index, line in enumerate(lines) if line == header]
        self.assertEqual(len(matches), 1, header)
        start = matches[0]
        indentation = len(header) - len(header.lstrip())
        end = len(lines)
        for index in range(start + 1, len(lines)):
            line = lines[index]
            if line.strip() and len(line) - len(line.lstrip()) <= indentation:
                end = index
                break
        return "\n".join(lines[start:end]) + "\n"

    def yaml_step(self, job: str, name: str) -> str:
        return self.yaml_block(job, f"      - name: {name}")

    def yaml_direct_lines(self, block: str, indentation: int) -> tuple[str, ...]:
        return tuple(
            line.strip().split(" #", 1)[0].rstrip()
            for line in block.splitlines()[1:]
            if line.strip()
            and not line.lstrip().startswith("#")
            and len(line) - len(line.lstrip()) == indentation
        )

    def operative_lines(self, block: str) -> tuple[str, ...]:
        return tuple(
            line.strip().split(" #", 1)[0].rstrip()
            for line in block.splitlines()
            if line.strip() and not line.lstrip().startswith("#")
        )

    def yaml_mapping_key_count(self, source: str, key: str) -> int:
        count = 0
        for line in source.splitlines():
            stripped = line.strip().split(" #", 1)[0].rstrip()
            if not stripped or stripped.startswith("#") or ":" not in stripped:
                continue
            candidate = stripped.split(":", 1)[0].strip().strip("'\"").strip()
            if candidate == key:
                count += 1
        return count

    def yaml_steps(self, job: str) -> tuple[str, ...]:
        steps = self.yaml_block(job, "    steps:")
        lines = steps.splitlines()[1:]
        starts = tuple(
            index
            for index, line in enumerate(lines)
            if line.strip().startswith("- ")
            and not line.lstrip().startswith("#")
            and len(line) - len(line.lstrip()) == 6
        )
        self.assertTrue(starts)
        return tuple(
            "\n".join(lines[start:end]) + "\n"
            for start, end in zip(starts, starts[1:] + (len(lines),))
        )

    def yaml_step_identity(self, step: str, position: int) -> str:
        header = self.operative_lines(step)[0]
        if header.startswith("- name: "):
            return header.removeprefix("- name: ")
        if header.startswith("- id: "):
            return "id:" + header.removeprefix("- id: ")
        return f"step:{position}"

    def yaml_step_actions(self, step: str) -> tuple[str, ...]:
        actions = []
        for line in step.splitlines():
            indentation = len(line) - len(line.lstrip())
            stripped = line.strip()
            if indentation == 6 and stripped.startswith("- uses: "):
                value = stripped.removeprefix("- uses: ")
            elif indentation == 8 and stripped.startswith("uses: "):
                value = stripped.removeprefix("uses: ")
            else:
                continue
            actions.append(value.split(" #", 1)[0].rstrip())
        return tuple(actions)

    def yaml_run_lines(self, step: str) -> tuple[str, ...]:
        run = self.yaml_block(step, "        run: |")
        return self.operative_lines(run)[1:]

    def yaml_run_script(self, step: str) -> str:
        run = self.yaml_block(step, "        run: |")
        return "\n".join(
            line[10:] if line.startswith(" " * 10) else line
            for line in run.splitlines()[1:]
        )

    def ci_filters(self, ci: str) -> dict[str, tuple[str, ...]]:
        changes = self.yaml_block(ci, "  changes:")
        step = next(
            step
            for step in self.yaml_steps(changes)
            if "dorny/paths-filter@" in step
        )
        filters = self.yaml_block(step, "          filters: |")
        parsed: dict[str, list[str]] = {}
        current = ""
        for line in filters.splitlines()[1:]:
            indentation = len(line) - len(line.lstrip())
            stripped = line.strip()
            if indentation == 12 and stripped.endswith(":"):
                current = stripped[:-1]
                parsed[current] = []
            elif indentation == 14 and stripped.startswith("- "):
                self.assertTrue(current)
                parsed[current].append(stripped[2:].strip("'\""))
        return {name: tuple(paths) for name, paths in parsed.items()}

    def ci_filter_matches(self, patterns: tuple[str, ...], path: str) -> bool:
        return any(fnmatch.fnmatchcase(path, pattern) for pattern in patterns)

    def ci_gate_environment(
        self,
        *,
        selected: dict[str, bool],
        results: dict[str, str],
        event_name: str = "pull_request",
        changes_result: str = "success",
    ) -> dict[str, str]:
        environment = os.environ.copy()
        environment.update(
            {
                "EVENT_NAME": event_name,
                "CHANGES_RESULT": changes_result,
            }
        )
        for job in (
            "test",
            "msrv",
            "wasm",
            "python_bindings",
            "presentation_fidelity",
            "word_fidelity",
            "hash_harness",
            "supply_chain",
            "prose",
        ):
            key = job.upper()
            environment[f"{key}_SELECTED"] = str(selected.get(job, False)).lower()
            environment[f"{key}_RESULT"] = results.get(job, "skipped")
        return environment

    def assert_no_success_short_circuit(self, lines: tuple[str, ...]) -> None:
        for line in lines:
            tokens = tuple(
                token.strip("'\"()")
                for token in line.replace(";", " ")
                .replace("&&", " ")
                .replace("||", " ")
                .split()
            )
            self.assertNotIn("true", tokens, line)
            for index, token in enumerate(tokens[:-1]):
                self.assertFalse(
                    token in ("exit", "return") and tokens[index + 1] == "0",
                    line,
                )

    def test_each_filtered_ci_job_has_a_must_trigger_and_must_not_trigger_path(
        self,
    ) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        filters = self.ci_filters(ci)
        cases = {
            "test": ("crates/rdocx/src/lib.rs", "docs/hld/00-vision.md"),
            "msrv": ("crates/oxml-core/src/lib.rs", "docs/hld/00-vision.md"),
            "wasm": ("crates/rdocx-wasm/src/lib.rs", "docs/hld/00-vision.md"),
            "python_bindings": (
                "crates/rdocx-py/pyproject.toml",
                "docs/hld/00-vision.md",
            ),
            "presentation_fidelity": (
                "scripts/pptx-corpus-manifest.tsv",
                "docs/hld/00-vision.md",
            ),
            "word_fidelity": (
                "scripts/docx_ssim_harness.py",
                "docs/hld/00-vision.md",
            ),
            "hash_harness": (
                "scripts/hash_baseline.json",
                "docs/hld/00-vision.md",
            ),
            "supply_chain": ("Cargo.lock", "docs/hld/00-vision.md"),
            "prose": ("docs/hld/00-vision.md", "crates/rdocx/src/lib.rs"),
        }
        self.assertEqual(set(filters), set(cases))
        for job, (must_trigger, must_not_trigger) in cases.items():
            with self.subTest(job=job, path=must_trigger):
                self.assertTrue(self.ci_filter_matches(filters[job], must_trigger))
                self.assertTrue(
                    self.ci_filter_matches(
                        filters[job], ".github/workflows/ci.yml"
                    )
                )
            with self.subTest(job=job, path=must_not_trigger):
                self.assertFalse(
                    self.ci_filter_matches(filters[job], must_not_trigger)
                )
            narrowed = tuple(
                pattern
                for pattern in filters[job]
                if not fnmatch.fnmatchcase(must_trigger, pattern)
            )
            with self.subTest(job=job, mutation="narrowed"):
                self.assertFalse(self.ci_filter_matches(narrowed, must_trigger))

        changes = self.yaml_block(ci, "  changes:")
        permissions = self.yaml_block(changes, "    permissions:")
        self.assertEqual(
            self.yaml_direct_lines(permissions, 6),
            ("contents: read", "pull-requests: read"),
        )
        self.assertEqual(self.yaml_mapping_key_count(ci, "pull-requests"), 1)
        path_filter_step = next(
            step
            for step in self.yaml_steps(changes)
            if "dorny/paths-filter@" in step
        )
        self.assertEqual(
            self.yaml_step_actions(path_filter_step),
            (
                "dorny/paths-filter@"
                "ceb8a2b8f2d89434be7ff52d3de7ec3738c5cc9d",
            ),
        )

    def test_epubcheck_ci_gate_pins_artifact_and_invokes_exact_oracle(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        test_job = self.yaml_block(ci, "  test:")
        install = self.yaml_step(test_job, "Install pinned EPUBCheck 5.3.0")
        gate = self.yaml_step(test_job, "Run pinned EPUBCheck gate")
        self.assertIn(
            "https://github.com/w3c/epubcheck/releases/download/v5.3.0/"
            "epubcheck-5.3.0.zip",
            install,
        )
        self.assertIn(
            "6c07e68584b2e2ce2f89fe06e1246dfead3eb36b46b340e7d93524f29dcff6c5",
            install,
        )
        self.assertIn(
            "f7f96617c929371821609b88c8484d6dc9f24fe916499863c46094c5fb778a65",
            install,
        )
        self.assertIn("EPUBCHECK_JAR=$jar", install)
        self.assertIn(
            "epubcheck_5_3_0_accepts_the_source_built_publication", gate
        )
        self.assertIn("--ignored --nocapture", gate)
        self.assertNotIn("continue-on-error", install + gate)

    def test_large_document_ci_gate_is_exact_release_ignored_and_single_threaded(
        self,
    ) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_large_document_ci_gate(ci)

    def assert_large_document_ci_gate(self, ci: str) -> None:
        test_job = self.yaml_block(ci, "  test:")
        gate = self.yaml_step(test_job, "Run thousand-page performance gate")
        self.assertEqual(
            self.operative_lines(gate),
            (
                "- name: Run thousand-page performance gate",
                "run: >-",
                "cargo test --locked --release -p rdocx",
                "--test regression_test",
                self.LARGE_DOCUMENT_TEST,
                "-- --ignored --exact --test-threads=1 --nocapture",
            ),
        )

    def test_large_document_ci_gate_rejects_weakened_invocations(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        test_job = self.yaml_block(ci, "  test:")
        gate = self.yaml_step(test_job, "Run thousand-page performance gate")
        mutations = {
            "missing": "",
            "unlocked": gate.replace("--locked ", ""),
            "debug": gate.replace("--release ", ""),
            "not-ignored": gate.replace("--ignored ", ""),
            "not-exact": gate.replace("--exact ", ""),
            "parallel": gate.replace("--test-threads=1", "--test-threads=2"),
            "swallowed": gate.replace(
                "        run: >-\n", "        continue-on-error: true\n        run: >-\n"
            ),
        }
        for name, mutated_gate in mutations.items():
            mutated_ci = ci.replace(gate, mutated_gate, 1)
            self.assertNotEqual(mutated_ci, ci, name)
            with self.subTest(mutation=name), self.assertRaises(AssertionError):
                self.assert_large_document_ci_gate(mutated_ci)

    def test_docs_only_changes_skip_expensive_jobs_and_still_report_the_ci_gate(
        self,
    ) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        filters = self.ci_filters(ci)
        selected = {
            job
            for job, patterns in filters.items()
            if self.ci_filter_matches(patterns, "docs/hld/12-testing-strategy.md")
        }
        self.assertEqual(selected, {"prose"})

        output_names = {
            line.split(":", 1)[0]
            for line in self.yaml_direct_lines(
                self.yaml_block(
                    self.yaml_block(ci, "  changes:"), "    outputs:"
                ),
                6,
            )
        }
        self.assertEqual(output_names, set(filters))
        for job_name in (
            "test",
            "msrv",
            "wasm",
            "python-bindings",
            "presentation-fidelity",
            "word-fidelity",
            "hash-harness",
            "prose",
        ):
            job = self.yaml_block(ci, f"  {job_name}:")
            output = job_name.replace("-", "_")
            direct = self.yaml_direct_lines(job, 4)
            self.assertIn("needs: changes", direct)
            self.assertIn(
                f"if: needs.changes.outputs.{output} == 'true'", direct
            )
        supply_chain = self.yaml_block(ci, "  supply-chain:")
        self.assertIn("needs: changes", self.yaml_direct_lines(supply_chain, 4))
        self.assertIn("github.event_name == 'schedule'", supply_chain)
        self.assertIn("needs.changes.outputs.supply_chain == 'true'", supply_chain)

        ci_gate = self.yaml_block(ci, "  ci-gate:")
        self.assertIn("name: CI gate", self.yaml_direct_lines(ci_gate, 4))
        self.assertIn("if: always()", self.yaml_direct_lines(ci_gate, 4))

    def test_ci_gate_rejects_failed_selected_jobs_and_accepts_unselected_skips(
        self,
    ) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        gate = self.yaml_block(ci, "  ci-gate:")
        needs = {
            line.removeprefix("- ")
            for line in self.operative_lines(self.yaml_block(gate, "    needs:"))[1:]
        }
        filtered_jobs = {
            "test",
            "msrv",
            "wasm",
            "python-bindings",
            "presentation-fidelity",
            "word-fidelity",
            "hash-harness",
            "supply-chain",
            "prose",
        }
        self.assertEqual(needs, filtered_jobs | {"changes"})
        step = self.yaml_step(gate, "Validate filtered jobs")
        environment = self.yaml_block(step, "        env:")
        self.assertEqual(
            self.yaml_direct_lines(environment, 10),
            (
                "EVENT_NAME: ${{ github.event_name }}",
                "CHANGES_RESULT: ${{ needs.changes.result }}",
                "TEST_SELECTED: ${{ needs.changes.outputs.test }}",
                "TEST_RESULT: ${{ needs.test.result }}",
                "MSRV_SELECTED: ${{ needs.changes.outputs.msrv }}",
                "MSRV_RESULT: ${{ needs.msrv.result }}",
                "WASM_SELECTED: ${{ needs.changes.outputs.wasm }}",
                "WASM_RESULT: ${{ needs.wasm.result }}",
                "PYTHON_BINDINGS_SELECTED: ${{ needs.changes.outputs.python_bindings }}",
                "PYTHON_BINDINGS_RESULT: ${{ needs['python-bindings'].result }}",
                "PRESENTATION_FIDELITY_SELECTED: ${{ needs.changes.outputs.presentation_fidelity }}",
                "PRESENTATION_FIDELITY_RESULT: ${{ needs['presentation-fidelity'].result }}",
                "WORD_FIDELITY_SELECTED: ${{ needs.changes.outputs.word_fidelity }}",
                "WORD_FIDELITY_RESULT: ${{ needs['word-fidelity'].result }}",
                "HASH_HARNESS_SELECTED: ${{ needs.changes.outputs.hash_harness }}",
                "HASH_HARNESS_RESULT: ${{ needs['hash-harness'].result }}",
                "SUPPLY_CHAIN_SELECTED: ${{ needs.changes.outputs.supply_chain }}",
                "SUPPLY_CHAIN_RESULT: ${{ needs['supply-chain'].result }}",
                "PROSE_SELECTED: ${{ needs.changes.outputs.prose }}",
                "PROSE_RESULT: ${{ needs.prose.result }}",
            ),
        )
        script = self.yaml_run_script(step)

        selected = {"test": True}
        results = {"test": "success"}
        completed = subprocess.run(
            ("bash", "-eu", "-o", "pipefail", "-c", script),
            env=self.ci_gate_environment(selected=selected, results=results),
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)

        for bad_result in ("failure", "cancelled", "skipped"):
            with self.subTest(selected_result=bad_result):
                completed = subprocess.run(
                    ("bash", "-eu", "-o", "pipefail", "-c", script),
                    env=self.ci_gate_environment(
                        selected=selected,
                        results={"test": bad_result},
                    ),
                    capture_output=True,
                    text=True,
                    check=False,
                )
                self.assertNotEqual(completed.returncode, 0)

        completed = subprocess.run(
            ("bash", "-eu", "-o", "pipefail", "-c", script),
            env=self.ci_gate_environment(
                selected={},
                results={"test": "success"},
            ),
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertNotEqual(completed.returncode, 0)

        completed = subprocess.run(
            ("bash", "-eu", "-o", "pipefail", "-c", script),
            env=self.ci_gate_environment(
                selected={},
                results={"supply_chain": "success"},
                event_name="schedule",
                changes_result="skipped",
            ),
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)

        completed = subprocess.run(
            ("bash", "-eu", "-o", "pipefail", "-c", script),
            env=self.ci_gate_environment(
                selected={},
                results={},
                changes_result="failure",
            ),
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertNotEqual(completed.returncode, 0)

    def assert_pinned_poppler_installer_contract(self, installer: str) -> None:
        self.assertIn('POPLER_VERSION = "26.01.0"', installer)
        self.assertIn(
            'POPLER_SHA256 = "1cb944a4b88847f5fb6551683bc799db59f04990f5d8be07aba2acbf38601089"',
            installer,
        )
        self.assertIn("MAX_DOWNLOAD_BYTES = 8 * 1024 * 1024", installer)
        self.assertIn("MAX_ARCHIVE_MEMBERS = 2_048", installer)
        self.assertIn("MAX_EXTRACTED_BYTES = 64 * 1024 * 1024", installer)
        for tool in ("pdftoppm", "pdfinfo", "pdftotext"):
            self.assertIn(tool, installer)
        self.assertIn("safe_extract", installer)
        self.assertIn("-DENABLE_UTILS=ON", installer)
        self.assertIn('expected = f"{tool} version {POPLER_VERSION}"', installer)

    def assert_pinned_pandoc_installer_contract(self, installer: str) -> None:
        for required in (
            'PANDOC_VERSION = "3.10"',
            'PANDOC_SHA256 = "e0f8af62d0f267d22baa5bcefe6d5dda3a097ccc60de794b759fe03159923244"',
            "MAX_DOWNLOAD_BYTES = 40 * 1024 * 1024",
            "MAX_ARCHIVE_MEMBERS = 256",
            "MAX_EXTRACTED_BYTES = 160 * MEBIBYTE",
            'f"{ARCHIVE_ROOT}/bin/pandoc-lua": "pandoc"',
            'f"{ARCHIVE_ROOT}/bin/pandoc-server": "pandoc"',
            "member.issym()",
            "AUTHENTICATED_SKIPPED_SYMLINKS.get(member.name) == member.linkname",
            "if extracted_bytes > MAX_EXTRACTED_BYTES:",
            "if not member_path.parts or member_path.parts[0] != ARCHIVE_ROOT:",
            "if not member.isfile():",
            "if not executable.is_file():",
            'identity[0] != f"pandoc {PANDOC_VERSION}"',
            'platform.system() != "Linux"',
            'platform.machine() not in ("x86_64", "amd64")',
            'os.environ.get("GITHUB_PATH")',
            'os.environ.get("GITHUB_ENV")',
        ):
            self.assertIn(required, installer)

    def assert_ci_oxml_layout_package_inventory(self, ci: str) -> None:
        job = self.yaml_block(ci, "  package-oxml-layout:")
        inventory = self.yaml_step(job, "Check package inventory")
        script = self.yaml_run_script(inventory)
        listed_fonts = tuple(
            re.findall(
                r"^\s+(fonts/[^\s]+\.ttf)(?:\s*\\|\)\")$",
                script,
                re.MULTILINE,
            )
        )
        listed_legal = tuple(
            re.findall(
                r"^\s+(fonts/(?:LICENSE|NOTICE)-[^\s\\\)]+)(?:\s*\\|\)\")$",
                script,
                re.MULTILINE,
            )
        )
        fonts = workflow.REPO / "crates/oxml-layout/fonts"
        expected_fonts = tuple(
            f"fonts/{path.name}" for path in sorted(fonts.glob("*.ttf"))
        )
        expected_legal = tuple(
            f"fonts/{path.name}"
            for path in sorted((*fonts.glob("LICENSE-*"), *fonts.glob("NOTICE-*")))
        )
        self.assertEqual(len(expected_fonts), 24)
        self.assertEqual(len(expected_legal), 6)
        self.assertEqual(listed_fonts, expected_fonts)
        self.assertEqual(listed_legal, expected_legal)
        self.assertIn("diff -u", script)

    def assert_pinned_libreoffice_installer_contract(self, installer: str) -> None:
        self.assertIn('LIBREOFFICE_VERSION = "26.2.5.2"', installer)
        self.assertIn(
            'LIBREOFFICE_SHA256 = "2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e"',
            installer,
        )
        self.assertIn(
            '"https://download.documentfoundation.org/libreoffice/stable/26.2.5/"',
            installer,
        )
        self.assertIn("MAX_DOWNLOAD_BYTES = 224 * 1024 * 1024", installer)
        self.assertIn("MAX_ARCHIVE_MEMBERS = 256", installer)
        self.assertIn("MAX_EXTRACTED_BYTES = 256 * 1024 * 1024", installer)
        self.assertIn('INSTALL_ROOT = Path("/opt/libreoffice26.2")', installer)
        self.assertEqual(
            install_pinned_libreoffice.SYSTEM_RUNTIME_PACKAGES,
            (
                "libcairo2",
                "libcups2t64",
                "libdbus-1-3",
                "libfontconfig1",
                "libfreetype6",
                "libglib2.0-0t64",
                "libgssapi-krb5-2",
                "libnspr4",
                "libnss3",
                "libx11-6",
                "libx11-xcb1",
                "libxext6",
                "libxinerama1",
            ),
        )
        self.assertIn("safe_extract", installer)
        self.assertIn("apt-get", installer)
        self.assertIn("--no-install-recommends", installer)
        self.assertIn(
            '"LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb"',
            installer,
        )
        self.assertIn('os.environ.get("GITHUB_PATH")', installer)

    def assert_libreoffice_consumers_contract(self, ci: str) -> None:
        consumers = {
            "test": "Run full workspace suite",
            "msrv": "Run full workspace suite",
            "word-fidelity": "Run all-page Word SSIM trend and completeness gate",
        }
        for job_name, use_step in consumers.items():
            job = self.yaml_block(ci, f"  {job_name}:")
            self.assertIn("runs-on: ubuntu-24.04", job)
            install = self.yaml_step(job, "Install pinned LibreOffice 26.2.5.2")
            self.assertEqual(
                self.yaml_direct_lines(install, 8),
                ("run: python3 scripts/install_pinned_libreoffice.py",),
            )
            test_step = self.yaml_step(job, use_step)
            self.assertLess(job.index(install), job.index(test_step))
            self.assertNotIn("continue-on-error", install)
            self.assert_no_success_short_circuit(self.operative_lines(install))
        self.assertEqual(ci.count("python3 scripts/install_pinned_libreoffice.py"), 3)

    def assert_poppler_consumers_contract(self, ci: str) -> None:
        consumers = {
            "test": "cargo test --workspace",
            "python-bindings": "Run full Python binding suite",
            "presentation-fidelity": "Run all-slide SSIM trend and completeness gate",
            "word-fidelity": "Run all-page Word SSIM trend and completeness gate",
            "msrv": "cargo test --workspace",
        }
        for job_name, use_marker in consumers.items():
            job = self.yaml_block(ci, f"  {job_name}:")
            step = self.yaml_step(job, "Install pinned Poppler 26.01.0")
            self.assertEqual(
                self.yaml_direct_lines(step, 8),
                ("shell: bash", "run: |"),
            )
            lines = self.yaml_run_lines(step)
            self.assertIn("python3 scripts/install_pinned_poppler.py", lines)
            self.assert_no_success_short_circuit(lines)
            self.assertLess(job.index(step), job.index(use_marker))
            self.assertFalse(
                any(
                    "continue-on-error:" in line
                    for line in self.operative_lines(job)
                )
            )
        self.assertEqual(ci.count("python3 scripts/install_pinned_poppler.py"), 5)
        self.assertNotIn("brew install poppler", ci)
        self.assertNotIn("apt-get install poppler-utils", ci)

    def assert_word_fidelity_ci_contract(self, ci: str) -> None:
        job = self.yaml_block(ci, "  word-fidelity:")
        self.assertIn("runs-on: ubuntu-24.04", self.yaml_direct_lines(job, 4))
        steps = self.yaml_steps(job)
        self.assertEqual(
            self.yaml_step_actions(steps[0]),
            ("actions/checkout@de0fac2e4500dabe0009e67214ff5f5447ce83dd",),
        )
        self.assertEqual(
            self.yaml_step_actions(steps[1]),
            ("dtolnay/rust-toolchain@4360b52568e2003a75bf9bc1d59f33a8e3fc893c",),
        )
        self.assertIn("toolchain: 1.97.1", self.operative_lines(steps[1]))
        self.assertEqual(
            self.yaml_step_actions(steps[2]),
            ("Swatinem/rust-cache@c19371144df3bb44fab255c43d04cbc2ab54d1c4",),
        )
        prime = self.yaml_step(job, "Prime locked Cargo dependencies")
        self.assertEqual(
            self.yaml_direct_lines(prime, 8),
            ("run: cargo fetch --locked",),
        )
        self.assertEqual(steps[3], prime)
        self.assertEqual(
            self.operative_lines(job).count("run: cargo fetch --locked"), 1
        )
        fetch = self.yaml_step(job, "Fetch pinned Word corpus")
        gate = self.yaml_step(
            job, "Run all-page Word SSIM trend and completeness gate"
        )
        upload = self.yaml_step(job, "Retain Word fidelity evidence")
        self.assertEqual(
            self.yaml_direct_lines(fetch, 8),
            ("run: python3 scripts/fetch_docx_corpus.py",),
        )
        self.assertEqual(
            self.operative_lines(gate),
            (
                "- name: Run all-page Word SSIM trend and completeness gate",
                "run: >-",
                "python3 scripts/docx_ssim_harness.py --check",
                '--output-dir "${RUNNER_TEMP}/word-fidelity"',
            ),
        )
        self.assertIn("if: always()", self.operative_lines(upload))
        self.assertIn(
            "${{ runner.temp }}/word-fidelity/gate-evidence.json", upload
        )
        self.assertIn("${{ runner.temp }}/word-fidelity/ssim-results.tsv", upload)
        self.assertIn("if-no-files-found: error", self.operative_lines(upload))
        self.assertNotIn("continue-on-error", job)
        self.assertLess(job.index(prime), job.index(fetch))
        self.assertLess(job.index(fetch), job.index(gate))
        self.assertLess(job.index(gate), job.index(upload))

    def test_word_fidelity_primes_locked_dependencies_before_offline_build(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_word_fidelity_ci_contract(ci)

    def test_word_fidelity_ci_gate_rejects_weakened_invocations(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_word_fidelity_ci_contract(ci)
        job = self.yaml_block(ci, "  word-fidelity:")
        gate = self.yaml_step(
            job, "Run all-page Word SSIM trend and completeness gate"
        )
        upload = self.yaml_step(job, "Retain Word fidelity evidence")
        prime = self.yaml_step(job, "Prime locked Cargo dependencies")
        test_job = self.yaml_block(ci, "  test:")
        test_checkout = self.yaml_steps(test_job)[0]
        wrong_job_test = test_job.replace(
            test_checkout, test_checkout + prime, 1
        )
        mutations = {
            "missing-fetch": ci.replace(prime, "", 1),
            "unlocked-fetch": ci.replace(
                prime, prime.replace("cargo fetch --locked", "cargo fetch", 1), 1
            ),
            "duplicate-fetch": ci.replace(prime, prime + prime, 1),
            "misplaced-fetch": ci.replace(prime, "", 1).replace(
                gate, gate + prime, 1
            ),
            "wrong-job-fetch": ci.replace(prime, "", 1).replace(
                test_job, wrong_job_test, 1
            ),
            "missing-gate": ci.replace(gate, "", 1),
            "self-test-only": ci.replace(
                gate, gate.replace("--check", "--self-test", 1), 1
            ),
            "no-evidence-output": ci.replace(
                '          --output-dir "${RUNNER_TEMP}/word-fidelity"\n', "", 1
            ),
            "swallowed": ci.replace(
                gate,
                gate.replace(
                    "        run: >-\n",
                    "        continue-on-error: true\n        run: >-\n",
                    1,
                ),
                1,
            ),
            "missing-json": ci.replace(
                "            ${{ runner.temp }}/word-fidelity/gate-evidence.json\n",
                "",
                1,
            ),
            "warning-artifact": ci.replace(
                upload, upload.replace("if-no-files-found: error", "if-no-files-found: warn"), 1
            ),
        }
        for label, mutated in mutations.items():
            self.assertNotEqual(mutated, ci, label)
            with self.subTest(mutation=label), self.assertRaises(AssertionError):
                self.assert_word_fidelity_ci_contract(mutated)

    def assert_workspace_oracle_environment_contract(self, ci: str) -> None:
        setup_action = (
            "astral-sh/setup-uv@20cfd1bf945f4377ade1205e4dbc17946fc9a30d"
        )
        for job_name in ("test", "msrv"):
            job = self.yaml_block(ci, f"  {job_name}:")
            steps = self.yaml_steps(job)
            setup_steps = tuple(
                step
                for step in steps
                if self.yaml_step_actions(step) == (setup_action,)
            )
            self.assertEqual(len(setup_steps), 1, job_name)
            setup = setup_steps[0]
            self.assertEqual(
                self.yaml_direct_lines(setup, 8),
                (f"uses: {setup_action}", "with:"),
            )
            setup_with = self.yaml_block(setup, "        with:")
            self.assertEqual(
                self.yaml_direct_lines(setup_with, 10),
                ('version: "0.10.2"', "enable-cache: false"),
            )

            test_step = self.yaml_step(job, "Run full workspace suite")
            self.assertEqual(
                self.yaml_direct_lines(test_step, 8),
                ("env:", "run: >-"),
            )
            environment = self.yaml_block(test_step, "        env:")
            self.assertEqual(
                self.yaml_direct_lines(environment, 10),
                (
                    'UV_CACHE_DIR: "${{ runner.temp }}/uv-cache"',
                    'RUST_MIN_STACK: "8388608"',
                ),
            )
            self.assertIn("cargo test --workspace", test_step)
            self.assert_no_success_short_circuit(self.operative_lines(test_step))
            self.assertLess(job.index(setup), job.index(test_step))
            self.assertNotIn("continue-on-error", setup + test_step)
        self.assertEqual(self.yaml_mapping_key_count(ci, "RUST_MIN_STACK"), 2)

    def assert_python_pr_job_contract(self, ci: str) -> None:
        triggers = self.yaml_block(ci, "on:")
        trigger_keys = tuple(
            line.split(":", 1)[0]
            for line in self.yaml_direct_lines(triggers, 2)
        )
        self.assertEqual(trigger_keys, ("push", "pull_request", "schedule"))
        pull_request = self.yaml_block(triggers, "  pull_request:")
        self.assertEqual(self.yaml_direct_lines(pull_request, 4), ())

        root_permissions = self.yaml_block(ci, "permissions:")
        self.assertEqual(
            self.yaml_direct_lines(root_permissions, 2),
            ("contents: read",),
        )
        operative_ci = self.operative_lines(ci)
        self.assertFalse(any("id-token:" in line for line in operative_ci))
        self.assertFalse(any("write-all" in line for line in operative_ci))
        self.assertFalse(any("PYTEST_ADDOPTS" in line for line in operative_ci))

        job = self.yaml_block(ci, "  python-bindings:")
        direct = self.yaml_direct_lines(job, 4)
        self.assertEqual(
            direct,
            (
                "needs: changes",
                "if: needs.changes.outputs.python_bindings == 'true'",
                "name: Python bindings (${{ matrix.package.distribution }})",
                "runs-on: macos-26",
                "strategy:",
                "steps:",
            ),
        )
        self.assertFalse(
            any("continue-on-error:" in line for line in self.operative_lines(job))
        )

        strategy = self.yaml_block(job, "    strategy:")
        self.assertEqual(
            self.yaml_direct_lines(strategy, 6),
            ("fail-fast: false", "matrix:"),
        )
        matrix = self.yaml_block(strategy, "      matrix:")
        self.assertEqual(self.yaml_direct_lines(matrix, 8), ("package:",))
        package = self.yaml_block(matrix, "        package:")
        self.assertEqual(
            self.yaml_direct_lines(package, 10),
            (
                "- { distribution: rdocx, crate: rdocx-py, "
                'oracle: "python-docx==1.2.0" }',
                "- { distribution: rpptx, crate: rpptx-py, "
                'oracle: "python-pptx==1.0.2" }',
            ),
        )

        steps = self.yaml_steps(job)
        identities = tuple(
            self.yaml_step_identity(step, position)
            for position, step in enumerate(steps)
        )
        required_order = (
            "step:0",
            "step:1",
            "step:2",
            "Set up Python 3.12",
            "Install pinned Poppler 26.01.0",
            "Create isolated binding environment",
            "Build Python extension",
            "Run full Python binding suite",
        )
        self.assertEqual(identities, required_order)

        action_contract = (
            (
                steps[0],
                "actions/checkout@de0fac2e4500dabe0009e67214ff5f5447ce83dd",
            ),
            (
                steps[1],
                "dtolnay/rust-toolchain@4360b52568e2003a75bf9bc1d59f33a8e3fc893c",
            ),
            (
                steps[2],
                "Swatinem/rust-cache@c19371144df3bb44fab255c43d04cbc2ab54d1c4",
            ),
        )
        for action_step, expected_action in action_contract:
            self.assertEqual(self.yaml_step_actions(action_step), (expected_action,))
            self.assertEqual(self.yaml_direct_lines(action_step, 8), ())

        setup = self.yaml_step(job, "Set up Python 3.12")
        self.assertEqual(
            self.yaml_step_actions(setup),
            ("actions/setup-python@a309ff8b426b58ec0e2a45f0f869d46889d02405",),
        )
        self.assertEqual(
            self.yaml_direct_lines(setup, 8),
            (
                "uses: actions/setup-python@"
                "a309ff8b426b58ec0e2a45f0f869d46889d02405",
                "with:",
            ),
        )
        setup_with = self.yaml_block(setup, "        with:")
        self.assertEqual(
            self.yaml_direct_lines(setup_with, 10),
            ('python-version: "3.12.9"',),
        )

        poppler = self.yaml_step(job, "Install pinned Poppler 26.01.0")
        self.assertEqual(
            self.yaml_run_lines(poppler),
            (
                "brew install \\",
                "cmake ninja pkg-config fontconfig freetype jpeg-turbo \\",
                "libpng libtiff little-cms2 openjpeg",
                "python3 scripts/install_pinned_poppler.py",
            ),
        )

        environment = self.yaml_step(job, "Create isolated binding environment")
        self.assertEqual(
            self.yaml_run_lines(environment),
            (
                'binding_venv="${RUNNER_TEMP}/${{ matrix.package.distribution }}-venv"',
                'python -m venv "$binding_venv"',
                'binding_python="$binding_venv/bin/python"',
                '"$binding_python" -m pip install \\',
                'maturin==1.13.3 \\',
                'pytest==9.1.1 \\',
                '"${{ matrix.package.oracle }}"',
            ),
        )

        build = self.yaml_step(job, "Build Python extension")
        build_lines = self.yaml_run_lines(build)
        self.assertEqual(
            build_lines,
            (
                'binding_venv="${RUNNER_TEMP}/${{ matrix.package.distribution }}-venv"',
                'binding_python="$binding_venv/bin/python"',
                'VIRTUAL_ENV="$binding_venv" \\',
                '"$binding_python" -m maturin develop --locked \\',
                '--manifest-path "crates/${{ matrix.package.crate }}/Cargo.toml"',
            ),
        )

        tests = self.yaml_step(job, "Run full Python binding suite")
        self.assertEqual(
            self.yaml_direct_lines(tests, 8),
            ("shell: bash", "run: |"),
        )
        test_lines = self.yaml_run_lines(tests)
        self.assertEqual(
            test_lines,
            (
                'binding_venv="${RUNNER_TEMP}/${{ matrix.package.distribution }}-venv"',
                'binding_python="$binding_venv/bin/python"',
                '"$binding_python" -m pytest "crates/${{ matrix.package.crate }}/tests"',
            ),
        )
        self.assert_no_success_short_circuit(build_lines + test_lines)
        self.assertNotIn("|| true", job)
        self.assertNotIn("set +e", job)

        for rust_job_name in ("test", "clippy", "doc", "msrv"):
            rust_job = self.yaml_block(ci, f"  {rust_job_name}:")
            self.assertIn("--all-features", rust_job, rust_job_name)
            self.assertIn("--exclude rdocx-py", rust_job, rust_job_name)
            self.assertIn("--exclude rpptx-py", rust_job, rust_job_name)

    def test_python_pr_job_builds_both_extensions_before_pytest(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_python_pr_job_contract(ci)

    def test_pinned_poppler_installer_contract(self) -> None:
        installer_path = workflow.REPO / "scripts/install_pinned_poppler.py"
        self.assertTrue(
            installer_path.is_file(),
            "F-X012 requires one shared pinned Poppler installer",
        )
        installer = installer_path.read_text(encoding="utf-8")
        self.assert_pinned_poppler_installer_contract(installer)

        mutations = {
            "wrong-version": installer.replace("26.01.0", "26.02.0"),
            "wrong-checksum": installer.replace(
                "1cb944a4b88847f5fb6551683bc799db59f04990f5d8be07aba2acbf38601089",
                "0" * 64,
            ),
            "missing-member-bound": installer.replace(
                "MAX_ARCHIVE_MEMBERS = 2_048",
                "MAX_ARCHIVE_MEMBERS = len(members)",
            ),
            "missing-runtime-identity": installer.replace(
                'expected = f"{tool} version {POPLER_VERSION}"',
                'expected = tool',
            ),
        }
        for label, mutated in mutations.items():
            with self.subTest(mutation=label), self.assertRaises(AssertionError):
                self.assert_pinned_poppler_installer_contract(mutated)

    def assert_pinned_pandoc_ci_gate(self, ci: str) -> None:
        job = self.yaml_block(ci, "  test:")
        install = self.yaml_step(job, "Install pinned Pandoc 3.10")
        gate = self.yaml_step(job, "Run pinned Pandoc texmath gate")
        workspace = self.yaml_step(job, "Run full workspace suite")
        self.assertEqual(
            self.operative_lines(install),
            (
                "- name: Install pinned Pandoc 3.10",
                "run: python3 scripts/install_pinned_pandoc.py",
            ),
        )
        self.assertEqual(
            self.operative_lines(gate),
            (
                "- name: Run pinned Pandoc texmath gate",
                "run: >-",
                "cargo test --locked -p rdocx",
                "math::tests::mathml_and_latex_conversion_matches_pinned_pandoc_texmath_trees",
                "--lib -- --ignored --exact --nocapture",
            ),
        )
        self.assertLess(job.index(install), job.index(gate))
        self.assertLess(job.index(gate), job.index(workspace))
        for step in (install, gate):
            self.assertNotIn("continue-on-error", step)
            self.assertNotIn("if: false", step)
            self.assert_no_success_short_circuit(self.operative_lines(step))

    def test_pinned_pandoc_installer_and_ci_gate_are_exact(self) -> None:
        installer_path = workflow.REPO / "scripts/install_pinned_pandoc.py"
        self.assertTrue(installer_path.is_file())
        installer = installer_path.read_text(encoding="utf-8")
        self.assert_pinned_pandoc_installer_contract(installer)

        mutations = {
            "missing-platform-rejection": installer.replace(
                'if platform.system() != "Linux" or platform.machine() not in ("x86_64", "amd64"):',
                "if False:",
                1,
            ),
            "missing-archive-root-rejection": installer.replace(
                "if not member_path.parts or member_path.parts[0] != ARCHIVE_ROOT:",
                "if False:",
                1,
            ),
            "missing-non-file-rejection": installer.replace(
                "if not member.isfile():",
                "if False:",
                1,
            ),
            "missing-extracted-size-cap": installer.replace(
                "if extracted_bytes > MAX_EXTRACTED_BYTES:",
                "if False:",
                1,
            ),
            "missing-expected-executable-check": installer.replace(
                "if not executable.is_file():",
                "if False:",
                1,
            ),
            "missing-runtime-identity": installer.replace(
                'identity[0] != f"pandoc {PANDOC_VERSION}"',
                "False",
                1,
            ),
        }
        for label, mutated in mutations.items():
            self.assertNotEqual(mutated, installer, label)
            with self.subTest(mutation=label), self.assertRaises(AssertionError):
                self.assert_pinned_pandoc_installer_contract(mutated)

        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_pinned_pandoc_ci_gate(ci)

    def test_ci_oxml_layout_package_inventory_matches_bundled_assets(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_ci_oxml_layout_package_inventory(ci)

        noto_entries = (
            "fonts/NotoSansArabic.ttf",
            "fonts/NotoSansDevanagari.ttf",
            "fonts/NotoSansSC-FX058-subset.ttf",
            "fonts/NotoSansThai.ttf",
            "fonts/LICENSE-Noto",
            "fonts/NOTICE-Noto",
        )
        for entry in noto_entries:
            mutated = ci.replace(entry, "", 1)
            self.assertNotEqual(mutated, ci, entry)
            with self.subTest(missing=entry), self.assertRaises(AssertionError):
                self.assert_ci_oxml_layout_package_inventory(mutated)

    def test_pinned_pandoc_installer_accepts_authenticated_archive_with_bounded_headroom(
        self,
    ) -> None:
        installer = (
            workflow.REPO / "scripts/install_pinned_pandoc.py"
        ).read_text(encoding="utf-8")
        self.assert_pinned_pandoc_installer_contract(installer)

        authenticated_payload_bytes = 162_406_703
        ceiling = install_pinned_pandoc.MAX_EXTRACTED_BYTES
        self.assertEqual(ceiling, 160 * install_pinned_pandoc.MEBIBYTE)
        self.assertGreater(ceiling, authenticated_payload_bytes)
        self.assertLess(ceiling - authenticated_payload_bytes, ceiling // 25)

        for replacement in ("154 * MEBIBYTE", "256 * MEBIBYTE"):
            mutated = installer.replace("160 * MEBIBYTE", replacement, 1)
            self.assertNotEqual(mutated, installer, replacement)
            with self.subTest(bound=replacement), self.assertRaises(AssertionError):
                self.assert_pinned_pandoc_installer_contract(mutated)

        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)

            def write_archive(
                path: Path,
                members: tuple[tuple[str, bytes, str], ...],
            ) -> None:
                with tarfile.open(path, mode="w:gz") as archive:
                    executable = tarfile.TarInfo("pandoc-3.10/bin/pandoc")
                    executable.mode = 0o755
                    executable.size = len(b"reviewed pandoc")
                    archive.addfile(executable, io.BytesIO(b"reviewed pandoc"))
                    for name, member_type, linkname in members:
                        member = tarfile.TarInfo(name)
                        member.type = member_type
                        member.linkname = linkname
                        archive.addfile(member)

            exact_aliases = (
                ("pandoc-3.10/bin/pandoc-server", tarfile.SYMTYPE, "pandoc"),
                ("pandoc-3.10/bin/pandoc-lua", tarfile.SYMTYPE, "pandoc"),
            )
            accepted_archive = root / "accepted.tar.gz"
            write_archive(accepted_archive, exact_aliases)
            accepted_root = root / "accepted"
            executable = install_pinned_pandoc.safe_extract(
                accepted_archive,
                accepted_root,
            )
            self.assertEqual(executable.read_bytes(), b"reviewed pandoc")
            for name, _, _ in exact_aliases:
                alias = accepted_root / name
                self.assertFalse(alias.exists(), name)
                self.assertFalse(alias.is_symlink(), name)

            rejected_members = (
                (
                    "wrong-name",
                    "pandoc-3.10/bin/pandoc-other",
                    tarfile.SYMTYPE,
                    "pandoc",
                ),
                (
                    "wrong-target",
                    "pandoc-3.10/bin/pandoc-server",
                    tarfile.SYMTYPE,
                    "other",
                ),
                (
                    "hardlink",
                    "pandoc-3.10/bin/pandoc-server",
                    tarfile.LNKTYPE,
                    "pandoc",
                ),
                ("character-device", "pandoc-3.10/dev", tarfile.CHRTYPE, ""),
                ("block-device", "pandoc-3.10/block", tarfile.BLKTYPE, ""),
                ("fifo", "pandoc-3.10/fifo", tarfile.FIFOTYPE, ""),
            )
            for label, name, member_type, linkname in rejected_members:
                archive = root / f"{label}.tar.gz"
                write_archive(archive, ((name, member_type, linkname),))
                with self.subTest(member=label), self.assertRaisesRegex(
                    RuntimeError,
                    "non-file entry",
                ):
                    install_pinned_pandoc.safe_extract(
                        archive,
                        root / f"extract-{label}",
                    )

    def test_pinned_pandoc_ci_gate_rejects_disabled_or_noop_steps(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        job = self.yaml_block(ci, "  test:")
        install = self.yaml_step(job, "Install pinned Pandoc 3.10")
        gate = self.yaml_step(job, "Run pinned Pandoc texmath gate")
        workspace = self.yaml_step(job, "Run full workspace suite")
        mutations = {
            "missing-gate": ci.replace(gate, "", 1),
            "install-if-false": ci.replace(
                install,
                install.replace("        run:", "        if: false\n        run:", 1),
                1,
            ),
            "install-continue-on-error": ci.replace(
                install,
                install.replace(
                    "        run:",
                    "        continue-on-error: true\n        run:",
                    1,
                ),
                1,
            ),
            "install-noop": ci.replace(
                install,
                install.replace(
                    "run: python3 scripts/install_pinned_pandoc.py",
                    "run: python3 -c 'pass'",
                ),
                1,
            ),
            "gate-if-false": ci.replace(
                gate,
                gate.replace("        run:", "        if: false\n        run:", 1),
                1,
            ),
            "gate-continue-on-error": ci.replace(
                gate,
                gate.replace(
                    "        run:",
                    "        continue-on-error: true\n        run:",
                    1,
                ),
                1,
            ),
            "gate-noop": ci.replace(
                gate,
                gate.replace(
                    "cargo test --locked -p rdocx",
                    "python3 -c 'pass'",
                ),
                1,
            ),
            "gate-after-workspace": ci.replace(gate, "", 1).replace(
                workspace,
                workspace + gate,
                1,
            ),
        }
        for label, mutated in mutations.items():
            self.assertNotEqual(mutated, ci, label)
            with self.subTest(mutation=label), self.assertRaises(AssertionError):
                self.assert_pinned_pandoc_ci_gate(mutated)

    def test_pinned_pandoc_installer_enforces_archive_guards(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            with patch.object(
                install_pinned_pandoc.urllib.request,
                "urlopen",
                return_value=io.BytesIO(b"not the reviewed Pandoc source"),
            ):
                with self.assertRaisesRegex(RuntimeError, "SHA-256"):
                    install_pinned_pandoc.download_archive(root / "wrong.tar.gz")

            archive_path = root / "unsafe.tar.gz"
            with tarfile.open(archive_path, mode="w:gz") as archive:
                member = tarfile.TarInfo("pandoc-3.10/../../escape")
                member.size = 0
                archive.addfile(member, io.BytesIO())
            with self.assertRaisesRegex(RuntimeError, "unsafe path"):
                install_pinned_pandoc.safe_extract(
                    archive_path,
                    root / "extract",
                )

            occupied = root / "occupied"
            executable = occupied / "bin" / "pandoc"
            executable.parent.mkdir(parents=True)
            executable.write_text("not the reviewed executable", encoding="utf-8")
            with (
                patch.object(
                    install_pinned_pandoc.platform,
                    "system",
                    return_value="Linux",
                ),
                patch.object(
                    install_pinned_pandoc.platform,
                    "machine",
                    return_value="x86_64",
                ),
                patch.object(install_pinned_pandoc, "download_archive") as download,
                self.assertRaisesRegex(RuntimeError, "prefix must be absent or empty"),
            ):
                install_pinned_pandoc.install(occupied)
            download.assert_not_called()

    def test_workspace_viewer_jobs_install_pinned_libreoffice(self) -> None:
        installer_path = workflow.REPO / "scripts/install_pinned_libreoffice.py"
        self.assertTrue(
            installer_path.is_file(),
            "F-X012 requires one pinned Linux LibreOffice installer",
        )
        installer = installer_path.read_text(encoding="utf-8")
        self.assert_pinned_libreoffice_installer_contract(installer)

        mutations = {
            "wrong-version": installer.replace("26.2.5.2", "26.2.6.0"),
            "wrong-checksum": installer.replace(
                "2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e",
                "0" * 64,
            ),
            "missing-member-bound": installer.replace(
                "MAX_ARCHIVE_MEMBERS = 256",
                "MAX_ARCHIVE_MEMBERS = len(members)",
            ),
            "recommended-packages": installer.replace(
                '"--no-install-recommends",', '"--install-recommends",'
            ),
        }
        for label, mutated in mutations.items():
            with self.subTest(installer_mutation=label), self.assertRaises(
                AssertionError
            ):
                self.assert_pinned_libreoffice_installer_contract(mutated)

        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_libreoffice_consumers_contract(ci)
        for job_name in ("test", "msrv", "word-fidelity"):
            job = self.yaml_block(ci, f"  {job_name}:")
            install = self.yaml_step(job, "Install pinned LibreOffice 26.2.5.2")
            for label, mutated_install in {
                "missing": "",
                "if-false": install.replace(
                    "        run:", "        if: false\n        run:", 1
                ),
                "continue-on-error": install.replace(
                    "        run:",
                    "        continue-on-error: true\n        run:",
                    1,
                ),
                "exit-zero": install.replace(
                    "        run: python3",
                    "        run: exit 0\n          python3",
                    1,
                ),
            }.items():
                mutated = ci.replace(install, mutated_install, 1)
                with self.subTest(job=job_name, mutation=label), self.assertRaises(
                    AssertionError
                ):
                    self.assert_libreoffice_consumers_contract(mutated)

    def test_pinned_libreoffice_installer_enforces_runtime_guards(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            with patch.object(
                install_pinned_libreoffice.urllib.request,
                "urlopen",
                return_value=io.BytesIO(b"not the reviewed LibreOffice source"),
            ):
                with self.assertRaisesRegex(RuntimeError, "SHA-256"):
                    install_pinned_libreoffice.download_archive(root / "wrong.tar.gz")

            with (
                patch.object(install_pinned_libreoffice, "MAX_DOWNLOAD_BYTES", 8),
                patch.object(
                    install_pinned_libreoffice.urllib.request,
                    "urlopen",
                    return_value=io.BytesIO(b"x" * 9),
                ),
            ):
                with self.assertRaisesRegex(RuntimeError, "download bound"):
                    install_pinned_libreoffice.download_archive(root / "large.tar.gz")

            archive_path = root / "too-many.tar.gz"
            with tarfile.open(archive_path, mode="w:gz") as archive:
                for index in range(
                    install_pinned_libreoffice.MAX_ARCHIVE_MEMBERS + 1
                ):
                    member = tarfile.TarInfo(
                        f"{install_pinned_libreoffice.ARCHIVE_ROOT}/member-{index}"
                    )
                    member.size = 0
                    archive.addfile(member, io.BytesIO())
            with patch.object(
                tarfile.TarFile,
                "getmembers",
                side_effect=AssertionError("unbounded member-table allocation"),
            ):
                with self.assertRaisesRegex(RuntimeError, "member-count"):
                    install_pinned_libreoffice.safe_extract(
                        archive_path,
                        root / "extract",
                    )

            unsafe_archive = root / "unsafe.tar.gz"
            with tarfile.open(unsafe_archive, mode="w:gz") as archive:
                member = tarfile.TarInfo(
                    f"{install_pinned_libreoffice.ARCHIVE_ROOT}/../../escape"
                )
                member.size = 0
                archive.addfile(member, io.BytesIO())
            with self.assertRaisesRegex(RuntimeError, "unsafe path"):
                install_pinned_libreoffice.safe_extract(
                    unsafe_archive,
                    root / "unsafe-extract",
                )

            unsupported_archive = root / "unsupported.tar.gz"
            with tarfile.open(unsupported_archive, mode="w:gz") as archive:
                member = tarfile.TarInfo(
                    f"{install_pinned_libreoffice.ARCHIVE_ROOT}/unsupported"
                )
                member.type = tarfile.SYMTYPE
                member.linkname = "target"
                archive.addfile(member)
            with self.assertRaisesRegex(RuntimeError, "non-file entry"):
                install_pinned_libreoffice.safe_extract(
                    unsupported_archive,
                    root / "unsupported-extract",
                )

            core_package = (
                f"libobasis26.2-core_{install_pinned_libreoffice.LIBREOFFICE_VERSION}"
                "-2_amd64.deb"
            )
            impress_package = (
                f"libreoffice26.2-impress_{install_pinned_libreoffice.LIBREOFFICE_VERSION}"
                "-2_amd64.deb"
            )
            for missing, present in (
                ("libobasis26.2-core", impress_package),
                ("libreoffice26.2-impress", core_package),
            ):
                incomplete_archive = root / f"missing-{missing}.tar.gz"
                with tarfile.open(incomplete_archive, mode="w:gz") as archive:
                    member = tarfile.TarInfo(
                        f"{install_pinned_libreoffice.ARCHIVE_ROOT}/DEBS/{present}"
                    )
                    member.size = 0
                    archive.addfile(member, io.BytesIO())
                with self.subTest(missing=missing), self.assertRaisesRegex(
                    RuntimeError, f"missing {missing}"
                ):
                    install_pinned_libreoffice.safe_extract(
                        incomplete_archive,
                        root / f"missing-{missing}-extract",
                    )

            oversized_member = tarfile.TarInfo(
                f"{install_pinned_libreoffice.ARCHIVE_ROOT}/oversized"
            )
            oversized_member.size = (
                install_pinned_libreoffice.MAX_EXTRACTED_BYTES + 1
            )
            with patch.object(
                install_pinned_libreoffice.tarfile,
                "open",
                return_value=contextlib.nullcontext((oversized_member,)),
            ):
                with self.assertRaisesRegex(RuntimeError, "extracted-size"):
                    install_pinned_libreoffice.safe_extract(
                        root / "unused.tar.gz",
                        root / "oversized-extract",
                    )

            fake_soffice = root / "soffice"
            fake_soffice.write_text("not used", encoding="utf-8")
            wrong_identity = subprocess.CompletedProcess(
                [str(fake_soffice), "--version"],
                0,
                stdout="LibreOffice 99.0.0\n",
                stderr="",
            )
            with patch.object(
                install_pinned_libreoffice.subprocess,
                "run",
                return_value=wrong_identity,
            ):
                with self.assertRaisesRegex(RuntimeError, "unexpected LibreOffice"):
                    install_pinned_libreoffice.verify_soffice(fake_soffice)

            populated = root / "populated-prefix"
            populated.mkdir()
            with (
                patch.object(install_pinned_libreoffice, "INSTALL_ROOT", populated),
                patch.object(install_pinned_libreoffice.platform, "system", return_value="Linux"),
                patch.object(install_pinned_libreoffice.platform, "machine", return_value="x86_64"),
                patch.object(
                    install_pinned_libreoffice,
                    "download_archive",
                    side_effect=AssertionError("populated prefix must fail first"),
                ),
            ):
                with self.assertRaisesRegex(RuntimeError, "prefix must be absent"):
                    install_pinned_libreoffice.install()

            deb_root = root / "complete" / "DEBS"
            deb_root.mkdir(parents=True)
            for package in (core_package, impress_package):
                (deb_root / package).write_bytes(b"package")
            events: list[str] = []

            def record_download(_destination: Path) -> None:
                events.append("download")

            def record_install(command: list[str], *, check: bool) -> None:
                self.assertTrue(check)
                self.assertEqual(command[:5], ["sudo", "apt-get", "install", "--yes", "--no-install-recommends"])
                for package in install_pinned_libreoffice.SYSTEM_RUNTIME_PACKAGES:
                    self.assertIn(package, command)
                self.assertIn(str(deb_root / core_package), command)
                self.assertIn(str(deb_root / impress_package), command)
                events.append("install")

            with (
                patch.object(install_pinned_libreoffice, "INSTALL_ROOT", root / "absent"),
                patch.object(install_pinned_libreoffice.platform, "system", return_value="Linux"),
                patch.object(install_pinned_libreoffice.platform, "machine", return_value="x86_64"),
                patch.dict(os.environ, {"RUNNER_TEMP": str(root)}),
                patch.object(install_pinned_libreoffice, "download_archive", side_effect=record_download),
                patch.object(
                    install_pinned_libreoffice,
                    "safe_extract",
                    side_effect=lambda _archive, _destination: (
                        events.append("extract") or deb_root
                    ),
                ),
                patch.object(install_pinned_libreoffice.subprocess, "run", side_effect=record_install),
                patch.object(
                    install_pinned_libreoffice,
                    "verify_soffice",
                    side_effect=lambda: events.append("verify"),
                ),
                patch.object(
                    install_pinned_libreoffice,
                    "expose_soffice",
                    side_effect=lambda: events.append("expose"),
                ),
            ):
                install_pinned_libreoffice.install()
            self.assertEqual(
                events,
                ["download", "extract", "install", "verify", "expose"],
            )

    def test_pinned_poppler_installer_enforces_its_runtime_guards(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)

            with patch.object(
                install_pinned_poppler.urllib.request,
                "urlopen",
                return_value=io.BytesIO(b"not the reviewed Poppler source"),
            ):
                with self.assertRaisesRegex(RuntimeError, "SHA-256"):
                    install_pinned_poppler.download_archive(root / "poppler.tar.xz")

            with patch.object(
                install_pinned_poppler.urllib.request,
                "urlopen",
                return_value=io.BytesIO(
                    b"x" * (install_pinned_poppler.MAX_DOWNLOAD_BYTES + 1)
                ),
            ):
                with self.assertRaisesRegex(RuntimeError, "download bound"):
                    install_pinned_poppler.download_archive(root / "too-large.tar.xz")

            archive_path = root / "too-many-members.tar.xz"
            with tarfile.open(archive_path, mode="w:xz") as archive:
                for index in range(install_pinned_poppler.MAX_ARCHIVE_MEMBERS + 1):
                    member = tarfile.TarInfo(f"poppler-26.01.0/member-{index}")
                    member.size = 0
                    archive.addfile(member, io.BytesIO())
            with patch.object(
                tarfile.TarFile,
                "getmembers",
                side_effect=AssertionError("unbounded member-table allocation"),
            ):
                with self.assertRaisesRegex(RuntimeError, "member-count"):
                    install_pinned_poppler.safe_extract(
                        archive_path,
                        root / "extract",
                    )

            oversized_member = tarfile.TarInfo("poppler-26.01.0/oversized")
            oversized_member.size = install_pinned_poppler.MAX_EXTRACTED_BYTES + 1
            with patch.object(
                install_pinned_poppler.tarfile,
                "open",
                return_value=contextlib.nullcontext((oversized_member,)),
            ):
                with self.assertRaisesRegex(RuntimeError, "extracted-size"):
                    install_pinned_poppler.safe_extract(
                        root / "unused.tar.xz",
                        root / "oversized-extract",
                    )

            for wrong_tool in install_pinned_poppler.TOOLS:
                prefix = root / f"wrong-{wrong_tool}"
                binary_root = prefix / "bin"
                binary_root.mkdir(parents=True)
                for tool in install_pinned_poppler.TOOLS:
                    version = "99.0.0" if tool == wrong_tool else "26.01.0"
                    executable = binary_root / tool
                    executable.write_text(
                        f"#!/bin/sh\necho '{tool} version {version}' >&2\n",
                        encoding="utf-8",
                    )
                    executable.chmod(0o755)
                with self.subTest(wrong_tool=wrong_tool), self.assertRaisesRegex(
                    RuntimeError,
                    f"unexpected {wrong_tool} identity",
                ):
                    install_pinned_poppler.verify_tools(prefix)

    def test_pinned_poppler_installer_rejects_a_populated_prefix(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            prefix = Path(temp) / "populated"
            binary_root = prefix / "bin"
            binary_root.mkdir(parents=True)
            for tool in install_pinned_poppler.TOOLS:
                executable = binary_root / tool
                executable.write_text(
                    f"#!/bin/sh\necho '{tool} version 26.01.0' >&2\n",
                    encoding="utf-8",
                )
                executable.chmod(0o755)
            with patch.object(
                install_pinned_poppler,
                "download_archive",
                side_effect=AssertionError("download must not be bypassed"),
            ):
                with self.assertRaisesRegex(RuntimeError, "prefix must be empty"):
                    install_pinned_poppler.build(prefix)

    def test_every_poppler_consumer_uses_the_pinned_installer(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_poppler_consumers_contract(ci)
        for job_name in (
            "test",
            "python-bindings",
            "presentation-fidelity",
            "word-fidelity",
            "msrv",
        ):
            marker = f"  {job_name}:"
            job = self.yaml_block(ci, marker)
            step = self.yaml_step(job, "Install pinned Poppler 26.01.0")
            mutated_job = job.replace(
                "          python3 scripts/install_pinned_poppler.py\n", "", 1
            )
            mutated = ci.replace(job, mutated_job, 1)
            with self.subTest(missing_consumer=job_name), self.assertRaises(
                AssertionError
            ):
                self.assert_poppler_consumers_contract(mutated)
            for policy in ("if: false", "continue-on-error: true"):
                weakened_step = step.replace(
                    "        shell: bash\n",
                    f"        {policy}\n        shell: bash\n",
                    1,
                )
                weakened = ci.replace(step, weakened_step, 1)
                with self.subTest(job=job_name, policy=policy), self.assertRaises(
                    AssertionError
                ):
                    self.assert_poppler_consumers_contract(weakened)
            short_circuited_step = step.replace(
                "        run: |\n",
                "        run: |\n          exit 0\n",
                1,
            )
            short_circuited = ci.replace(step, short_circuited_step, 1)
            with self.subTest(job=job_name, policy="exit 0"), self.assertRaises(
                AssertionError
            ):
                self.assert_poppler_consumers_contract(short_circuited)

    def test_workspace_oracle_jobs_pin_uv_cache_and_stack(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_workspace_oracle_environment_contract(ci)
        mutations = {
            "global-stack": ci.replace(
                "  CARGO_TERM_COLOR: always\n",
                "  CARGO_TERM_COLOR: always\n  RUST_MIN_STACK: \"8388608\"\n",
                1,
            ),
            "global-quoted-stack": ci.replace(
                "  CARGO_TERM_COLOR: always\n",
                "  CARGO_TERM_COLOR: always\n  \"RUST_MIN_STACK\": \"8388608\"\n",
                1,
            ),
            "global-spaced-stack": ci.replace(
                "  CARGO_TERM_COLOR: always\n",
                "  CARGO_TERM_COLOR: always\n  RUST_MIN_STACK : \"8388608\"\n",
                1,
            ),
            "global-quoted-spaced-stack": ci.replace(
                "  CARGO_TERM_COLOR: always\n",
                "  CARGO_TERM_COLOR: always\n  \"RUST_MIN_STACK\" : \"8388608\"\n",
                1,
            ),
        }
        for job_name in ("test", "msrv"):
            job = self.yaml_block(ci, f"  {job_name}:")
            test_step = self.yaml_step(job, "Run full workspace suite")
            mutations.update(
                {
                    f"{job_name}-wrong-action": ci.replace(
                        job,
                        job.replace(
                            "astral-sh/setup-uv@"
                            "20cfd1bf945f4377ade1205e4dbc17946fc9a30d",
                            "astral-sh/setup-uv@main",
                            1,
                        ),
                        1,
                    ),
                    f"{job_name}-wrong-uv-version": ci.replace(
                        job,
                        job.replace(
                            'version: "0.10.2"', 'version: "latest"', 1
                        ),
                        1,
                    ),
                    f"{job_name}-shared-home-cache": ci.replace(
                        job,
                        job.replace(
                            'UV_CACHE_DIR: "${{ runner.temp }}/uv-cache"',
                            'UV_CACHE_DIR: "~/.cache/uv"',
                            1,
                        ),
                        1,
                    ),
                    f"{job_name}-default-stack": ci.replace(
                        job,
                        job.replace(
                            '          RUST_MIN_STACK: "8388608"\n', "", 1
                        ),
                        1,
                    ),
                    f"{job_name}-exit-zero": ci.replace(
                        test_step,
                        test_step.replace(
                            "        run: >-\n",
                            "        run: >-\n          exit 0\n",
                            1,
                        ),
                        1,
                    ),
                }
            )
        for label, mutated in mutations.items():
            with self.subTest(mutation=label), self.assertRaises(AssertionError):
                self.assert_workspace_oracle_environment_contract(mutated)

    def test_wasm_job_accepts_the_official_binaryen_125_identity(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        job = self.yaml_block(ci, "  wasm:")
        install = self.yaml_step(job, "Install wasm-opt 125")
        self.assertIn(
            'wasm-opt version 125 (version_125)',
            "\n".join(self.yaml_run_lines(install)),
        )

    def test_workspace_test_jobs_fetch_the_pinned_presentation_corpus(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        for job_name in ("test", "msrv"):
            with self.subTest(job=job_name):
                job = self.yaml_block(ci, f"  {job_name}:")
                fetch = self.yaml_step(job, "Fetch pinned presentation corpus")
                self.assertEqual(
                    self.yaml_direct_lines(fetch, 8),
                    ("run: python3 scripts/fetch_pptx_corpus.py",),
                )
                test_steps = tuple(
                    step
                    for step in self.yaml_steps(job)
                    if "cargo test --workspace" in step
                )
                self.assertEqual(len(test_steps), 1)
                self.assertLess(job.index(fetch), job.index(test_steps[0]))
                self.assertNotIn("continue-on-error", fetch)

    def test_word_corpus_fetcher_verifies_every_checksum_and_refuses_a_mismatch(
        self,
    ) -> None:
        content = {
            category: f"{category} content".encode()
            for category in sorted(fetch_docx_corpus.EXPECTED_CATEGORIES)
        }
        entries = [
            (
                f"{category}.docx",
                category,
                "test-producer",
                "Apache-2.0",
                fetch_docx_corpus.LICENCE_URLS["Apache-2.0"],
                hashlib.sha256(content[category]).hexdigest(),
                f"https://example.com/{category}.docx",
            )
            for category in sorted(fetch_docx_corpus.EXPECTED_CATEGORIES)
        ]
        with tempfile.TemporaryDirectory(dir=workflow.REPO) as directory:
            corpus = Path(directory)
            for entry in entries:
                (corpus / entry[0]).write_bytes(content[entry[1]])
            fetch_docx_corpus.verify_directory(corpus, entries)
            for entry in entries:
                source = corpus / entry[0]
                source.write_bytes(b"changed")
                with self.subTest(checksum=entry[0]), self.assertRaisesRegex(
                    ValueError, "digest mismatch"
                ):
                    fetch_docx_corpus.verify_directory(corpus, entries)
                source.write_bytes(content[entry[1]])

            destination = corpus / entries[-1][0]
            destination.write_bytes(b"stale")
            temporary = corpus / f".{entries[-1][0]}.download"
            with contextlib.redirect_stdout(io.StringIO()):
                with patch.object(
                    fetch_docx_corpus.urllib.request,
                    "urlopen",
                    return_value=io.BytesIO(b"wrong download"),
                ), self.assertRaisesRegex(ValueError, "digest mismatch"):
                    fetch_docx_corpus.fetch(corpus, entries)
            self.assertEqual(destination.read_bytes(), b"stale")
            self.assertFalse(temporary.exists())

    def test_word_corpus_fetcher_refuses_missing_extra_and_unlicensed_inputs(
        self,
    ) -> None:
        header = (
            "path\tcategory\tproducer\tlicence\tlicence_url\tsha256\turl\n"
        )
        lines = [
            "\t".join(
                (
                    f"{category}.docx",
                    category,
                    "test-producer",
                    "Apache-2.0",
                    fetch_docx_corpus.LICENCE_URLS["Apache-2.0"],
                    hashlib.sha256(category.encode()).hexdigest(),
                    f"https://example.com/{category}.docx",
                )
            )
            for category in sorted(fetch_docx_corpus.EXPECTED_CATEGORIES)
        ]
        valid = header + "\n".join(lines) + "\n"
        with tempfile.TemporaryDirectory(dir=workflow.REPO) as directory:
            root = Path(directory)
            manifest = root / "manifest.tsv"
            manifest.write_text(valid, encoding="utf-8")
            entries = fetch_docx_corpus.load_manifest(manifest)

            corpus = root / "corpus"
            corpus.mkdir()
            for entry in entries:
                (corpus / entry[0]).write_bytes(entry[1].encode())
            (corpus / entries[0][0]).unlink()
            with self.assertRaisesRegex(ValueError, "corpus is missing"):
                fetch_docx_corpus.verify_directory(corpus, entries)
            (corpus / entries[0][0]).write_bytes(entries[0][1].encode())
            (corpus / "extra.docx").write_bytes(b"extra")
            with self.assertRaisesRegex(ValueError, "unmanifested files"):
                fetch_docx_corpus.verify_directory(corpus, entries)

            mutations = {
                "unsafe": valid.replace("business-letter.docx", "../unsafe.docx", 1),
                "unlicensed": valid.replace("Apache-2.0", "GPL-3.0-only", 1),
                "incomplete": valid.replace("\tmulti-script\t", "\treport\t", 1),
                "insecure-source": valid.replace(
                    "https://example.com/business-letter.docx",
                    "http://example.com/business-letter.docx",
                    1,
                ),
                "unapproved-licence-url": valid.replace(
                    fetch_docx_corpus.LICENCE_URLS["Apache-2.0"],
                    "https://example.com/LICENSE",
                    1,
                ),
            }
            for label, mutated in mutations.items():
                with self.subTest(mutation=label):
                    manifest.write_text(mutated, encoding="utf-8")
                    with self.assertRaises(ValueError):
                        fetch_docx_corpus.load_manifest(manifest)

    def test_workspace_test_jobs_fetch_the_pinned_word_corpus(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        for job_name in ("test", "msrv"):
            with self.subTest(job=job_name):
                job = self.yaml_block(ci, f"  {job_name}:")
                fetch = self.yaml_step(job, "Fetch pinned Word corpus")
                self.assertEqual(
                    self.yaml_direct_lines(fetch, 8),
                    ("run: python3 scripts/fetch_docx_corpus.py",),
                )
                test_steps = tuple(
                    step
                    for step in self.yaml_steps(job)
                    if "cargo test --workspace" in step
                )
                self.assertEqual(len(test_steps), 1)
                self.assertLess(job.index(fetch), job.index(test_steps[0]))
                self.assertNotIn("continue-on-error", fetch)

    def test_python_pr_job_rejects_failure_swallowing_and_incomplete_cells(
        self,
    ) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_python_pr_job_contract(ci)
        python_job = self.yaml_block(ci, "  python-bindings:")

        def mutate_job(old: str, new: str) -> str:
            self.assertIn(old, python_job)
            return ci.replace(python_job, python_job.replace(old, new, 1), 1)

        mutations = {
            "missing-pull-request-trigger": ci.replace(
                "  pull_request:\n", "", 1
            ),
            "commented-pull-request-trigger": ci.replace(
                "  pull_request:\n", "  # pull_request:\n", 1
            ),
            "root-contents-write": ci.replace(
                "  contents: read\n", "  contents: write\n", 1
            ),
            "root-contents-write-with-required-comment": ci.replace(
                "  contents: read\n",
                "  contents: write # contents: read\n",
                1,
            ),
            "root-id-token-write": ci.replace(
                "  contents: read\n",
                "  contents: read\n  id-token: write\n",
                1,
            ),
            "job-if-false": ci.replace(
                "    name: Python bindings (${{ matrix.package.distribution }})\n",
                "    name: Python bindings (${{ matrix.package.distribution }})\n"
                "    if: false\n",
                1,
            ),
            "job-if-true": ci.replace(
                "    name: Python bindings (${{ matrix.package.distribution }})\n",
                "    name: Python bindings (${{ matrix.package.distribution }})\n"
                "    if: true\n",
                1,
            ),
            "job-pytest-collect-only-environment": ci.replace(
                "    runs-on: macos-26\n",
                "    runs-on: macos-26\n"
                "    env:\n"
                "      PYTEST_ADDOPTS: --collect-only\n",
                1,
            ),
            "root-pytest-collect-only-environment": ci.replace(
                "env:\n  CARGO_TERM_COLOR: always\n",
                "env:\n"
                "  CARGO_TERM_COLOR: always\n"
                "  PYTEST_ADDOPTS: --collect-only\n",
                1,
            ),
            "missing-rpptx-cell": ci.replace(
                '          - { distribution: rpptx, crate: rpptx-py, oracle: "python-pptx==1.0.2" }\n',
                "",
                1,
            ),
            "cancel-other-package-on-failure": ci.replace(
                "fail-fast: false", "fail-fast: true", 1
            ),
            "unversioned-pytest": ci.replace("pytest==9.1.1", "pytest", 1),
            "wrong-python-version": ci.replace(
                'python-version: "3.12.9"', 'python-version: "3.13"', 1
            ),
            "wrong-rdocx-oracle": ci.replace(
                "python-docx==1.2.0", "python-docx==1.1.2", 1
            ),
            "missing-develop": ci.replace(
                "maturin develop --locked", "maturin --version", 1
            ),
            "single-test-file": ci.replace(
                '"crates/${{ matrix.package.crate }}/tests"',
                '"crates/${{ matrix.package.crate }}/tests/test_core.py"',
                1,
            ),
            "continue-on-error": ci.replace(
                "      - name: Run full Python binding suite\n",
                "      - name: Run full Python binding suite\n        continue-on-error: true\n",
                1,
            ),
            "continue-on-error-false": ci.replace(
                "      - name: Run full Python binding suite\n",
                "      - name: Run full Python binding suite\n"
                "        continue-on-error: false\n",
                1,
            ),
            "pytest-if-false": ci.replace(
                "      - name: Run full Python binding suite\n",
                "      - name: Run full Python binding suite\n"
                "        if: false\n",
                1,
            ),
            "pytest-if-true": ci.replace(
                "      - name: Run full Python binding suite\n",
                "      - name: Run full Python binding suite\n"
                "        if: true\n",
                1,
            ),
            "pytest-step-environment": ci.replace(
                "      - name: Run full Python binding suite\n",
                "      - name: Run full Python binding suite\n"
                "        env:\n"
                "          PYTEST_ADDOPTS: --collect-only\n",
                1,
            ),
            "successful-pytest-fallback": ci.replace(
                '"$binding_python" -m pytest "crates/${{ matrix.package.crate }}/tests"',
                '"$binding_python" -m pytest "crates/${{ matrix.package.crate }}/tests" || true',
                1,
            ),
            "wrong-checkout-sha": mutate_job(
                "de0fac2e4500dabe0009e67214ff5f5447ce83dd",
                "0000000000000000000000000000000000000000",
            ),
            "wrong-checkout-sha-with-required-comment": mutate_job(
                "actions/checkout@de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2",
                "actions/checkout@0000000000000000000000000000000000000000 "
                "# v6.0.2 de0fac2e4500dabe0009e67214ff5f5447ce83dd",
            ),
            "checkout-ref-input": mutate_job(
                "      - uses: actions/checkout@"
                "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n",
                "      - uses: actions/checkout@"
                "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n"
                "        with:\n"
                "          ref: main\n",
            ),
            "wrong-rust-toolchain-sha": ci.replace(
                "4360b52568e2003a75bf9bc1d59f33a8e3fc893c",
                "0000000000000000000000000000000000000000",
                1,
            ),
            "rust-toolchain-input": ci.replace(
                "      - uses: dtolnay/rust-toolchain@"
                "4360b52568e2003a75bf9bc1d59f33a8e3fc893c # stable\n",
                "      - uses: dtolnay/rust-toolchain@"
                "4360b52568e2003a75bf9bc1d59f33a8e3fc893c # stable\n"
                "        with:\n"
                "          toolchain: nightly\n",
                1,
            ),
            "wrong-rust-cache-sha": ci.replace(
                "c19371144df3bb44fab255c43d04cbc2ab54d1c4",
                "0000000000000000000000000000000000000000",
                1,
            ),
            "rust-cache-input": ci.replace(
                "      - uses: Swatinem/rust-cache@"
                "c19371144df3bb44fab255c43d04cbc2ab54d1c4 # v2.9.1\n",
                "      - uses: Swatinem/rust-cache@"
                "c19371144df3bb44fab255c43d04cbc2ab54d1c4 # v2.9.1\n"
                "        with:\n"
                "          workspaces: crates/rdocx-py\n",
                1,
            ),
            "wrong-setup-python-sha": ci.replace(
                "a309ff8b426b58ec0e2a45f0f869d46889d02405",
                "0000000000000000000000000000000000000000",
                1,
            ),
            "wrong-setup-python-sha-with-required-comment": ci.replace(
                "actions/setup-python@"
                "a309ff8b426b58ec0e2a45f0f869d46889d02405 # v6.2.0",
                "actions/setup-python@"
                "0000000000000000000000000000000000000000 "
                "# v6.2.0 a309ff8b426b58ec0e2a45f0f869d46889d02405",
                1,
            ),
            "setup-python-extra-input": ci.replace(
                '          python-version: "3.12.9"\n',
                '          python-version: "3.12.9"\n'
                '          architecture: "x64"\n',
                1,
            ),
            "setup-python-comment-smuggled-version": ci.replace(
                '          python-version: "3.12.9"\n',
                '          python-version: "3.13" # python-version: "3.12.9"\n',
                1,
            ),
            "missing-rpptx-exclusion": ci.replace(
                "--exclude rdocx-py --exclude rpptx-py",
                "--exclude rdocx-py",
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, ci, name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_python_pr_job_contract(mutated)

    def assert_wasm_pr_job_contract(self, ci: str) -> None:
        triggers = self.yaml_block(ci, "on:")
        trigger_keys = tuple(
            line.split(":", 1)[0]
            for line in self.yaml_direct_lines(triggers, 2)
        )
        self.assertEqual(trigger_keys, ("push", "pull_request", "schedule"))
        pull_request = self.yaml_block(triggers, "  pull_request:")
        self.assertEqual(self.yaml_direct_lines(pull_request, 4), ())

        root_permissions = self.yaml_block(ci, "permissions:")
        self.assertEqual(
            self.yaml_direct_lines(root_permissions, 2),
            ("contents: read",),
        )
        operative_ci = self.operative_lines(ci)
        self.assertFalse(any("id-token:" in line for line in operative_ci))
        self.assertFalse(any("write-all" in line for line in operative_ci))

        job = self.yaml_block(ci, "  wasm:")
        self.assertEqual(
            self.yaml_direct_lines(job, 4),
            (
                "needs: changes",
                "if: needs.changes.outputs.wasm == 'true'",
                "name: WASM",
                "runs-on: ubuntu-latest",
                "steps:",
            ),
        )
        self.assertFalse(
            any("continue-on-error:" in line for line in self.operative_lines(job))
        )

        steps = self.yaml_steps(job)
        identities = tuple(
            self.yaml_step_identity(step, position)
            for position, step in enumerate(steps)
        )
        self.assertEqual(
            identities,
            (
                "step:0",
                "step:1",
                "step:2",
                "Set up Node 24.11.1",
                "Install wasm-pack 0.15.0",
                "Install wasm-opt 125",
                "Check WASM targets",
                "Run WASM Node tests",
                "Build and install local WASM packages",
            ),
        )

        action_contract = (
            (
                steps[0],
                "actions/checkout@de0fac2e4500dabe0009e67214ff5f5447ce83dd",
            ),
            (
                steps[1],
                "dtolnay/rust-toolchain@4360b52568e2003a75bf9bc1d59f33a8e3fc893c",
            ),
            (
                steps[2],
                "Swatinem/rust-cache@c19371144df3bb44fab255c43d04cbc2ab54d1c4",
            ),
        )
        for action_step, expected_action in action_contract:
            self.assertEqual(self.yaml_step_actions(action_step), (expected_action,))

        rust_inputs = self.yaml_block(steps[1], "        with:")
        self.assertEqual(
            self.yaml_direct_lines(rust_inputs, 10),
            ("targets: wasm32-unknown-unknown",),
        )
        self.assertEqual(self.yaml_direct_lines(steps[0], 8), ())
        self.assertEqual(self.yaml_direct_lines(steps[2], 8), ())

        node = self.yaml_step(job, "Set up Node 24.11.1")
        self.assertEqual(
            self.yaml_step_actions(node),
            (
                "actions/setup-node@"
                "249970729cb0ef3589644e2896645e5dc5ba9c38",
            ),
        )
        self.assertEqual(
            self.yaml_direct_lines(node, 8),
            (
                "uses: actions/setup-node@"
                "249970729cb0ef3589644e2896645e5dc5ba9c38",
                "with:",
            ),
        )
        node_inputs = self.yaml_block(node, "        with:")
        self.assertEqual(
            self.yaml_direct_lines(node_inputs, 10),
            ('node-version: "24.11.1"',),
        )

        install = self.yaml_step(job, "Install wasm-pack 0.15.0")
        install_optimizer = self.yaml_step(job, "Install wasm-opt 125")
        checks = self.yaml_step(job, "Check WASM targets")
        node_tests = self.yaml_step(job, "Run WASM Node tests")
        packages = self.yaml_step(job, "Build and install local WASM packages")
        for command_step in (install, install_optimizer, checks, node_tests, packages):
            self.assertEqual(
                self.yaml_direct_lines(command_step, 8),
                ("shell: bash", "run: |"),
            )
        install_lines = self.yaml_run_lines(install)
        optimizer_lines = self.yaml_run_lines(install_optimizer)
        check_lines = self.yaml_run_lines(checks)
        node_test_lines = self.yaml_run_lines(node_tests)
        self.assertEqual(
            install_lines,
            ("cargo install wasm-pack --version 0.15.0 --locked",),
        )
        self.assertEqual(
            optimizer_lines,
            (
                'binaryen_archive="${RUNNER_TEMP}/binaryen-version_125-x86_64-linux.tar.gz"',
                'binaryen_root="${RUNNER_TEMP}/binaryen-version_125"',
                "curl --fail --location --silent --show-error "
                '"https://github.com/WebAssembly/binaryen/releases/download/'
                'version_125/binaryen-version_125-x86_64-linux.tar.gz" '
                '--output "$binaryen_archive"',
                'echo "7c3bc16599c8274a04d34a504fe4be2047884f900e0e2da2f6fb9cd667183be4  '
                '$binaryen_archive" | sha256sum --check',
                'mkdir -p "$binaryen_root"',
                'tar --extract --gzip --file "$binaryen_archive" --directory '
                '"$binaryen_root" --strip-components=1',
                'echo "$binaryen_root/bin" >> "$GITHUB_PATH"',
                'test "$("$binaryen_root/bin/wasm-opt" --version)" = '
                '"wasm-opt version 125 (version_125)"',
            ),
        )
        self.assertEqual(
            check_lines,
            (
                "cargo check --locked --target wasm32-unknown-unknown -p rdocx-wasm",
                "cargo check --locked --target wasm32-unknown-unknown -p rpptx-wasm",
            ),
        )
        self.assertEqual(
            node_test_lines,
            (
                "wasm-pack test --node crates/rdocx-wasm",
                "wasm-pack test --node crates/rpptx-wasm",
            ),
        )
        self.assert_no_success_short_circuit(
            install_lines + optimizer_lines + check_lines + node_test_lines
        )
        package_lines = self.yaml_run_lines(packages)
        for expected in (
            'package_root="${RUNNER_TEMP}/wasm-packages"',
            'tarball_root="${RUNNER_TEMP}/wasm-tarballs"',
            'npm_cache="${RUNNER_TEMP}/npm-cache"',
            "wasm-pack build --target bundler --scope tensorbee --release "
            '--out-dir "$package_root/rdocx-wasm" crates/rdocx-wasm --locked',
            "wasm-pack build --target bundler --scope tensorbee --release "
            '--out-dir "$package_root/rpptx-wasm" crates/rpptx-wasm --locked',
            'verify_package "$package_root/rdocx-wasm" "@tensorbee/rdocx-wasm" '
            '"0.13.0" "rdocx_wasm"',
            'verify_package "$package_root/rpptx-wasm" "@tensorbee/rpptx-wasm" '
            '"0.11.0" "rpptx_wasm"',
            "npm install --prefix \"$consumer_root\" --cache \"$npm_cache\" "
            "--ignore-scripts --no-audit --no-fund --package-lock=false "
            '"$tarball_root/$tarball"',
        ):
            self.assertEqual(package_lines.count(expected), 1, expected)
        self.assertIn(
            'npm pack "$package_dir" --cache "$npm_cache" --ignore-scripts '
            '--pack-destination "$tarball_root"',
            packages,
        )
        self.assertIn('import(\\"$expected_name\\")', packages)
        self.assertIn('consumer_root="$(mktemp -d ', packages)
        self.assertIn('manifest.name !== expectedName', packages)
        self.assertIn('manifest.version !== expectedVersion', packages)
        self.assertIn('${stem}_bg.wasm', packages)
        self.assertIn('${stem}.js', packages)
        self.assertIn('${stem}.d.ts', packages)
        forbidden = (
            "npm publish",
            "npm login",
            "npm adduser",
            "npm token",
            "wasm-pack publish",
            "NODE_AUTH_TOKEN",
            "NPM_TOKEN",
            "--registry",
            "id-token:",
            "git tag",
            "gh release",
        )
        operative_job = "\n".join(self.operative_lines(job))
        for command in forbidden:
            self.assertNotIn(command, operative_job)
        self.assert_no_success_short_circuit(package_lines)
        self.assertNotIn("|| true", job)
        self.assertNotIn("set +e", job)

    def assert_wasm_optimizer_metadata_contract(
        self, manifest_overrides: dict[str, str] | None = None
    ) -> None:
        manifest_overrides = manifest_overrides or {}
        expected = {
            "wasm-opt": [
                "-Oz",
                "--enable-bulk-memory",
                "--enable-nontrapping-float-to-int",
            ]
        }
        for member in ("crates/rdocx-wasm", "crates/rpptx-wasm"):
            manifest = tomllib.loads(
                manifest_overrides.get(
                    member,
                    (workflow.REPO / member / "Cargo.toml").read_text(
                        encoding="utf-8"
                    ),
                )
            )
            wasm_pack = manifest["package"].get("metadata", {}).get("wasm-pack", {})
            release = wasm_pack.get("profile", {}).get("release")
            self.assertEqual(
                release,
                expected,
                member,
            )

    def test_wasm_pr_job_checks_both_targets_and_runs_node_tests(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_wasm_pr_job_contract(ci)

    def test_wasm_packages_use_the_reviewed_release_optimizer(self) -> None:
        self.assert_wasm_optimizer_metadata_contract()

    def test_wasm_package_contract_rejects_optimizer_mutations(self) -> None:
        for member in ("crates/rdocx-wasm", "crates/rpptx-wasm"):
            manifest = (workflow.REPO / member / "Cargo.toml").read_text(
                encoding="utf-8"
            )
            mutations = {
                "missing-bulk-memory": manifest.replace(
                    '["-Oz", "--enable-bulk-memory", '
                    '"--enable-nontrapping-float-to-int"]',
                    '["-Oz", "--enable-nontrapping-float-to-int"]',
                    1,
                ),
                "missing-nontrapping-float-to-int": manifest.replace(
                    '["-Oz", "--enable-bulk-memory", '
                    '"--enable-nontrapping-float-to-int"]',
                    '["-Oz", "--enable-bulk-memory"]',
                    1,
                ),
                "wrong-size-level": manifest.replace(
                    '["-Oz", "--enable-bulk-memory", '
                    '"--enable-nontrapping-float-to-int"]',
                    '["-Os", "--enable-bulk-memory", '
                    '"--enable-nontrapping-float-to-int"]',
                    1,
                ),
            }
            for name, mutated in mutations.items():
                self.assertNotEqual(mutated, manifest, f"{member}:{name}")
                with self.subTest(member=member, name=name), self.assertRaises(
                    AssertionError
                ):
                    self.assert_wasm_optimizer_metadata_contract({member: mutated})

    def assert_wasm_setup_node_provenance_contract(
        self, ci: str, testing_hld: str
    ) -> None:
        reviewed_sha = "249970729cb0ef3589644e2896645e5dc5ba9c38"
        reviewed_tag = "v6.5.0"
        job = self.yaml_block(ci, "  wasm:")
        provenance_line = (
            f"        uses: actions/setup-node@{reviewed_sha} # {reviewed_tag}"
        )
        self.assertEqual(job.count(provenance_line), 1)
        self.assertIn(f"setup-node {reviewed_tag}", testing_hld)
        self.assertNotIn("setup-node v6.1.0", testing_hld)

    def test_wasm_setup_node_provenance_matches_the_testing_hld(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        testing_hld = (workflow.REPO / "docs/hld/12-testing-strategy.md").read_text(
            encoding="utf-8"
        )
        self.assert_wasm_setup_node_provenance_contract(ci, testing_hld)

        mutations = {
            "stale-workflow-comment": (
                ci.replace(
                    "249970729cb0ef3589644e2896645e5dc5ba9c38 # v6.5.0",
                    "249970729cb0ef3589644e2896645e5dc5ba9c38 # v6.1.0",
                    1,
                ),
                testing_hld,
            ),
            "stale-hld-label": (
                ci,
                testing_hld.replace("setup-node v6.5.0", "setup-node v6.1.0", 1),
            ),
        }
        for name, (mutated_ci, mutated_hld) in mutations.items():
            self.assertTrue(
                mutated_ci != ci or mutated_hld != testing_hld,
                name,
            )
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_wasm_setup_node_provenance_contract(
                    mutated_ci, mutated_hld
                )

    def test_wasm_pr_job_rejects_skipped_or_weakened_gates(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_wasm_pr_job_contract(ci)
        wasm_job = self.yaml_block(ci, "  wasm:")

        def mutate_job(old: str, new: str) -> str:
            self.assertIn(old, wasm_job)
            return ci.replace(wasm_job, wasm_job.replace(old, new, 1), 1)

        mutations = {
            "missing-pull-request-trigger": ci.replace(
                "  pull_request:\n", "", 1
            ),
            "commented-pull-request-trigger": ci.replace(
                "  pull_request:\n", "  # pull_request:\n", 1
            ),
            "root-contents-write": ci.replace(
                "  contents: read\n", "  contents: write\n", 1
            ),
            "root-id-token-write": ci.replace(
                "  contents: read\n",
                "  contents: read\n  id-token: write\n",
                1,
            ),
            "job-condition": mutate_job(
                "    name: WASM\n", "    name: WASM\n    if: true\n"
            ),
            "wrong-checkout-sha": mutate_job(
                "de0fac2e4500dabe0009e67214ff5f5447ce83dd",
                "0000000000000000000000000000000000000000",
            ),
            "wrong-rust-toolchain-sha": mutate_job(
                "4360b52568e2003a75bf9bc1d59f33a8e3fc893c",
                "0000000000000000000000000000000000000000",
            ),
            "wrong-rust-cache-sha": mutate_job(
                "c19371144df3bb44fab255c43d04cbc2ab54d1c4",
                "0000000000000000000000000000000000000000",
            ),
            "wrong-setup-node-sha": mutate_job(
                "249970729cb0ef3589644e2896645e5dc5ba9c38",
                "0000000000000000000000000000000000000000",
            ),
            "wrong-node-version": mutate_job("24.11.1", "24"),
            "unlocked-wasm-pack-install": mutate_job(
                "cargo install wasm-pack --version 0.15.0 --locked",
                "cargo install wasm-pack --version 0.15.0",
            ),
            "floating-wasm-pack-version": mutate_job(
                "cargo install wasm-pack --version 0.15.0 --locked",
                "cargo install wasm-pack --locked",
            ),
            "wrong-wasm-opt-version": mutate_job(
                "binaryen-version_125-x86_64-linux.tar.gz",
                "binaryen-version_124-x86_64-linux.tar.gz",
            ),
            "wrong-wasm-opt-checksum": mutate_job(
                "7c3bc16599c8274a04d34a504fe4be2047884f900e0e2da2f6fb9cd667183be4",
                "0000000000000000000000000000000000000000000000000000000000000000",
            ),
            "missing-wasm-opt-version-check": mutate_job(
                '          test "$("$binaryen_root/bin/wasm-opt" --version)" = '
                '"wasm-opt version 125 (version_125)"\n',
                "",
            ),
            "unlocked-target-check": mutate_job(
                "cargo check --locked --target wasm32-unknown-unknown -p rdocx-wasm",
                "cargo check --target wasm32-unknown-unknown -p rdocx-wasm",
            ),
            "missing-rdocx-target-check": mutate_job(
                "          cargo check --locked --target wasm32-unknown-unknown -p rdocx-wasm\n",
                "",
            ),
            "missing-rpptx-target-check": mutate_job(
                "          cargo check --locked --target wasm32-unknown-unknown -p rpptx-wasm\n",
                "",
            ),
            "missing-rdocx-node-test": mutate_job(
                "          wasm-pack test --node crates/rdocx-wasm\n", ""
            ),
            "missing-rpptx-node-test": mutate_job(
                "          wasm-pack test --node crates/rpptx-wasm\n", ""
            ),
            "missing-node-runner": mutate_job(
                "wasm-pack test --node crates/rdocx-wasm",
                "wasm-pack test crates/rdocx-wasm",
            ),
            "listing-only-node-test": mutate_job(
                "wasm-pack test --node crates/rdocx-wasm",
                "wasm-pack test --node crates/rdocx-wasm -- --list",
            ),
            "check-condition": mutate_job(
                "      - name: Check WASM targets\n",
                "      - name: Check WASM targets\n        if: true\n",
            ),
            "node-test-condition": mutate_job(
                "      - name: Run WASM Node tests\n",
                "      - name: Run WASM Node tests\n        if: true\n",
            ),
            "continue-on-error": mutate_job(
                "      - name: Run WASM Node tests\n",
                "      - name: Run WASM Node tests\n"
                "        continue-on-error: true\n",
            ),
            "successful-fallback": mutate_job(
                "wasm-pack test --node crates/rdocx-wasm",
                "wasm-pack test --node crates/rdocx-wasm || true",
            ),
            "early-success": mutate_job(
                "        run: |\n"
                "          wasm-pack test --node crates/rdocx-wasm",
                "        run: |\n"
                "          exit 0\n"
                "          wasm-pack test --node crates/rdocx-wasm",
            ),
            "missing-rdocx-package": mutate_job(
                "          wasm-pack build --target bundler --scope tensorbee "
                "--release --out-dir \"$package_root/rdocx-wasm\" "
                "crates/rdocx-wasm --locked\n",
                "",
            ),
            "wrong-package-target": mutate_job(
                "wasm-pack build --target bundler",
                "wasm-pack build --target nodejs",
            ),
            "wrong-package-scope": mutate_job(
                "--scope tensorbee --release",
                "--scope other --release",
            ),
            "unlocked-package-build": mutate_job(
                "crates/rdocx-wasm --locked",
                "crates/rdocx-wasm",
            ),
            "missing-clean-install": mutate_job(
                "          npm install --prefix \"$consumer_root\" --cache "
                "\"$npm_cache\" --ignore-scripts --no-audit --no-fund "
                "--package-lock=false \"$tarball_root/$tarball\"\n",
                "",
            ),
            "registry-authentication": mutate_job(
                "      - name: Build and install local WASM packages\n",
                "      - name: Build and install local WASM packages\n"
                "        env:\n"
                "          NPM_TOKEN: forbidden\n",
            ),
            "npm-publish-authority": mutate_job(
                '          assert_inventory "$package_dir" "$expected_name" '
                '"$expected_version" "$stem"\n',
                '          assert_inventory "$package_dir" "$expected_name" '
                '"$expected_version" "$stem"\n'
                '          npm publish "$package_dir"\n',
            ),
            "release-tag-authority": mutate_job(
                "      - name: Build and install local WASM packages\n",
                "      - name: Build and install local WASM packages\n"
                "        env:\n"
                "          RELEASE_COMMAND: git tag v0.0.0\n",
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, ci, name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_wasm_pr_job_contract(mutated)

    def assert_wheels_workflow_contract(self, workflow_bytes: bytes) -> None:
        self.assertEqual(
            hashlib.sha256(workflow_bytes).hexdigest(),
            "56491248b4ffa7ea40abe75b04a16fcfd5c24744d16ccb9a8c6f7110d39be35a",
        )
        wheels = workflow_bytes.decode("utf-8", errors="strict")
        expected_packages = (
            ("rdocx", "rdocx-py", "rdocx"),
            ("rpptx", "rpptx-py", "rpptx"),
        )
        expected_platforms = (
            (
                "manylinux_2_28-x86_64",
                "ubuntu-24.04",
                "x86_64-unknown-linux-gnu",
                "2_28",
                "native",
            ),
            (
                "manylinux_2_28-aarch64",
                "ubuntu-24.04-arm",
                "aarch64-unknown-linux-gnu",
                "2_28",
                "native",
            ),
            (
                "musllinux_1_2-x86_64",
                "ubuntu-24.04",
                "x86_64-unknown-linux-musl",
                "musllinux_1_2",
                "musl",
            ),
            (
                "macos-x86_64",
                "macos-15-intel",
                "x86_64-apple-darwin",
                "off",
                "native",
            ),
            (
                "macos-arm64",
                "macos-14",
                "aarch64-apple-darwin",
                "off",
                "native",
            ),
            (
                "windows-x86_64",
                "windows-2025",
                "x86_64-pc-windows-msvc",
                "off",
                "native",
            ),
        )
        triggers = self.yaml_block(wheels, "on:")
        trigger_keys = tuple(
            line.split(":", 1)[0]
            for line in self.yaml_direct_lines(triggers, 2)
        )
        self.assertEqual(trigger_keys, ("push", "workflow_dispatch"))
        push_trigger = self.yaml_block(triggers, "  push:")
        self.assertEqual(
            self.yaml_direct_lines(push_trigger, 4),
            ('tags: ["py-v*"]',),
        )
        root_permissions = self.yaml_block(wheels, "permissions:")
        self.assertEqual(
            self.yaml_direct_lines(root_permissions, 2),
            ("contents: read",),
        )
        self.assertNotIn("secrets.", wheels)
        self.assertNotIn("git tag", wheels)
        self.assertNotIn("git push", wheels)
        self.assertNotIn("write-all", wheels)
        self.assertNotIn("continue-on-error:", wheels)

        jobs = self.yaml_block(wheels, "jobs:")
        job_keys = tuple(
            line.split(":", 1)[0]
            for line in self.yaml_direct_lines(jobs, 2)
        )
        self.assertEqual(job_keys, ("build-wheels", "build-sdists", "publish"))
        build_wheels = self.yaml_block(wheels, "  build-wheels:")
        build_sdists = self.yaml_block(wheels, "  build-sdists:")
        publish = self.yaml_block(wheels, "  publish:")
        job_names = {
            key: tuple(
                line.removeprefix("name: ")
                for line in self.yaml_direct_lines(job, 4)
                if line.startswith("name: ")
            )
            for key, job in (
                ("build-wheels", build_wheels),
                ("build-sdists", build_sdists),
                ("publish", publish),
            )
        }
        self.assertEqual(
            job_names,
            {
                "build-wheels": (
                    "${{ matrix.package.distribution }} "
                    "${{ matrix.platform.label }}",
                ),
                "build-sdists": ("${{ matrix.package.distribution }} sdist",),
                "publish": ("Publish Python distributions",),
            },
        )
        publication_jobs = tuple(
            key
            for key, names in job_names.items()
            if "publish" in key.lower()
            or any("publish" in name.lower() for name in names)
        )
        self.assertEqual(publication_jobs, ("publish",))

        for build_job in (build_wheels, build_sdists):
            direct = self.yaml_direct_lines(build_job, 4)
            self.assertFalse(any(line.startswith("if:") for line in direct))
            self.assertFalse(any(line.startswith("permissions:") for line in direct))

        cp39_setup = self.yaml_block(build_wheels, "      - id: cp39")
        cp312_setup = self.yaml_block(build_wheels, "      - id: cp312")
        sdist_steps = self.yaml_steps(build_sdists)
        self.assertGreaterEqual(len(sdist_steps), 2)
        sdist_setup = sdist_steps[1]
        self.assertEqual(self.yaml_step_identity(sdist_setup, 2), "step:2")
        for setup, version, condition in (
            (cp39_setup, '"3.9"', ()),
            (
                cp312_setup,
                '"3.12"',
                ("if: matrix.platform.install == 'native'",),
            ),
            (sdist_setup, '"3.9"', ()),
        ):
            setup_conditions = tuple(
                line
                for line in self.yaml_direct_lines(setup, 8)
                if line.startswith("if:")
            )
            self.assertEqual(setup_conditions, condition)
            setup_inputs = self.yaml_block(setup, "        with:")
            self.assertEqual(
                self.yaml_direct_lines(setup_inputs, 10),
                (f"python-version: {version}",),
            )

        matrix = self.yaml_block(build_wheels, "      matrix:")
        matrix_axes = tuple(
            line.strip().split(":", 1)[0]
            for line in matrix.splitlines()[1:]
            if line.strip() and len(line) - len(line.lstrip()) == 8
        )
        self.assertEqual(matrix_axes, ("package", "platform"))

        package_matrix = self.yaml_block(matrix, "        package:")
        package_entries = tuple(
            line.strip()
            for line in package_matrix.splitlines()[1:]
            if len(line) - len(line.lstrip()) == 10
        )
        expected_package_entries = tuple(
            (
                f"- {{ distribution: {distribution}, crate: {crate}, "
                f"module: {module} }}"
            )
            for distribution, crate, module in expected_packages
        )
        self.assertEqual(package_entries, expected_package_entries)

        platform_matrix = self.yaml_block(matrix, "        platform:")
        platform_entries = tuple(
            line.strip()
            for line in platform_matrix.splitlines()[1:]
            if len(line) - len(line.lstrip()) == 10
        )
        expected_platform_entries = tuple(
            (
                f"- {{ label: {label}, runner: {runner}, target: {target}, "
                f"manylinux: {manylinux}, install: {install} }}"
            )
            for label, runner, target, manylinux, install in expected_platforms
        )
        self.assertEqual(platform_entries, expected_platform_entries)

        wheel_build = self.yaml_step(build_wheels, "Build cp39-abi3 wheel")
        self.assertFalse(
            any(
                line.startswith("if:")
                for line in self.yaml_direct_lines(wheel_build, 8)
            )
        )
        wheel_build_inputs = self.yaml_block(wheel_build, "        with:")
        self.assertEqual(
            self.yaml_direct_lines(wheel_build_inputs, 10),
            (
                "command: build",
                "maturin-version: v1.13.3",
                "target: ${{ matrix.platform.target }}",
                "manylinux: ${{ matrix.platform.manylinux }}",
                "working-directory: crates/${{ matrix.package.crate }}",
                "args: --release --locked --compatibility pypi --out ../../dist",
            ),
        )

        wheel_metadata = self.yaml_step(build_wheels, "Validate wheel metadata")
        native_install = self.yaml_step(build_wheels, "Install and test native wheel")
        typing = self.yaml_step(build_wheels, "Validate installed typing surface")
        for step in (native_install, typing):
            conditions = tuple(
                line.strip()
                for line in step.splitlines()
                if len(line) - len(line.lstrip()) == 8
                and line.strip().startswith("if:")
            )
            self.assertEqual(
                conditions,
                ("if: matrix.platform.install == 'native'",),
            )
        musllinux_install = self.yaml_step(
            build_wheels, "Install and test musllinux wheel"
        )
        musllinux_conditions = tuple(
            line
            for line in self.yaml_direct_lines(musllinux_install, 8)
            if line.startswith("if:")
        )
        self.assertEqual(
            musllinux_conditions,
            ("if: matrix.platform.install == 'musl'",),
        )

        metadata_run = self.yaml_run_lines(wheel_metadata)
        native_run = self.yaml_run_lines(native_install)
        typing_run = self.yaml_run_lines(typing)
        musllinux_run = self.yaml_run_lines(musllinux_install)
        for run in (metadata_run, native_run, typing_run, musllinux_run):
            self.assert_no_success_short_circuit(run)
        self.assertEqual(
            metadata_run,
            (
                "python - <<'PY'",
                "from pathlib import Path",
                "import zipfile",
                'wheel = next(Path("dist").glob('
                '"${{ matrix.package.distribution }}-*.whl"))',
                "with zipfile.ZipFile(wheel) as archive:",
                "wheel_metadata = next(",
                "name for name in archive.namelist() "
                'if name.endswith(".dist-info/WHEEL")',
                ")",
                "metadata = archive.read(wheel_metadata).decode()",
                'assert "Tag: cp39-abi3-" in metadata, metadata',
                "PY",
            ),
        )
        self.assertEqual(
            native_run,
            (
                '"${{ steps.cp39.outputs.python-path }}" -m venv .wheel-venv',
                'if [[ "${{ runner.os }}" == "Windows" ]]; then',
                "venv_python=.wheel-venv/Scripts/python",
                "else",
                "venv_python=.wheel-venv/bin/python",
                "fi",
                '"$venv_python" -m pip install --upgrade pip',
                '"$venv_python" -m pip install '
                'dist/${{ matrix.package.distribution }}-*.whl pytest '
                "python-docx==1.2.0 python-pptx==1.0.2",
                '"$venv_python" -c "import ${{ matrix.package.module }}"',
                'if [[ "${{ matrix.package.distribution }}" == "rdocx" ]]; then',
                '"$venv_python" -m pytest \\',
                "crates/rdocx-py/tests/test_core.py " + chr(92),
                "crates/rdocx-py/tests/test_formatting_tables.py " + chr(92),
                "crates/rdocx-py/tests/test_shared.py " + chr(92),
                "crates/rdocx-py/tests/test_python_docx_parity.py",
                "else",
                '"$venv_python" -m pytest '
                "crates/rpptx-py/tests/test_documented_examples.py",
                "fi",
            ),
        )
        self.assertEqual(
            typing_run,
            (
                '"${{ steps.cp312.outputs.python-path }}" -m venv .typing-venv',
                'if [[ "${{ runner.os }}" == "Windows" ]]; then',
                "typing_python=.typing-venv/Scripts/python",
                "else",
                "typing_python=.typing-venv/bin/python",
                "fi",
                '"$typing_python" -m pip install --upgrade pip',
                '"$typing_python" -m pip install '
                'dist/${{ matrix.package.distribution }}-*.whl mypy==2.3.0',
                '"$typing_python" -m mypy --strict \\',
                "crates/${{ matrix.package.crate }}/tests/typing_smoke.py "
                + chr(92),
                "crates/${{ matrix.package.crate }}/python/"
                "${{ matrix.package.module }}",
                'if [[ "${{ matrix.package.distribution }}" == "rdocx" ]]; then',
                '"$typing_python" -m mypy.stubtest rdocx',
                "else",
                '"$typing_python" -m mypy.stubtest rpptx',
                "fi",
            ),
        )
        self.assertEqual(
            musllinux_run,
            (
                "docker run --rm " + chr(92),
                '-v "$PWD:/workspace:ro" ' + chr(92),
                '-v "$PWD/dist:/dist:ro" ' + chr(92),
                "-w /workspace " + chr(92),
                '-e PACKAGE_DISTRIBUTION="${{ matrix.package.distribution }}" '
                + chr(92),
                '-e PACKAGE_MODULE="${{ matrix.package.module }}" ' + chr(92),
                "python:3.9-alpine " + chr(92),
                "sh -euxc '",
                "python -m venv /tmp/wheel-venv",
                "venv_python=/tmp/wheel-venv/bin/python",
                '"$venv_python" -m pip install --upgrade pip',
                '"$venv_python" -m pip install '
                "/dist/${PACKAGE_DISTRIBUTION}-*.whl pytest "
                "python-docx==1.2.0 python-pptx==1.0.2",
                '"$venv_python" -c "import ${PACKAGE_MODULE}"',
                'if [ "$PACKAGE_DISTRIBUTION" = rdocx ]; then',
                '"$venv_python" -m pytest ' + chr(92),
                "crates/rdocx-py/tests/test_core.py " + chr(92),
                "crates/rdocx-py/tests/test_formatting_tables.py " + chr(92),
                "crates/rdocx-py/tests/test_shared.py " + chr(92),
                "crates/rdocx-py/tests/test_python_docx_parity.py",
                "else",
                '"$venv_python" -m pytest '
                "crates/rpptx-py/tests/test_documented_examples.py",
                "fi",
                "'",
            ),
        )

        distribution_branch = (
            'if [[ "${{ matrix.package.distribution }}" == "rdocx" ]]; then'
        )
        self.assertEqual(native_install.count(distribution_branch), 1)
        self.assertIn("crates/rdocx-py/tests/test_python_docx_parity.py", native_install)
        self.assertIn(
            "crates/rpptx-py/tests/test_documented_examples.py", native_install
        )
        self.assertIn(
            'if [ "$PACKAGE_DISTRIBUTION" = rdocx ]; then', musllinux_install
        )
        self.assertIn(
            "crates/rdocx-py/tests/test_python_docx_parity.py",
            musllinux_install,
        )
        self.assertIn(
            "crates/rpptx-py/tests/test_documented_examples.py",
            musllinux_install,
        )
        self.assertEqual(typing.count(distribution_branch), 1)
        self.assertIn('"$typing_python" -m mypy.stubtest rdocx\n', typing)
        self.assertIn('"$typing_python" -m mypy.stubtest rpptx\n', typing)

        upload_wheel = self.yaml_step(build_wheels, "Upload wheel")
        self.assertFalse(
            any(
                line.startswith("if:")
                for line in self.yaml_direct_lines(upload_wheel, 8)
            )
        )
        self.assertIn(
            "uses: actions/upload-artifact@"
            "ea165f8d65b6e75b540449e92b4886f43607fa02",
            upload_wheel,
        )
        upload_wheel_inputs = self.yaml_block(upload_wheel, "        with:")
        self.assertEqual(
            self.yaml_direct_lines(upload_wheel_inputs, 10),
            (
                "name: artifact-wheel-${{ matrix.package.distribution }}-"
                "${{ matrix.platform.label }}",
                "path: dist/*.whl",
                "if-no-files-found: error",
            ),
        )

        self.assertIn("Tag: cp39-abi3-", wheels)
        self.assertIn('python-version: "3.9"', wheels)
        self.assertIn('python-version: "3.12"', wheels)
        self.assertIn('-m venv .wheel-venv', wheels)
        self.assertIn('-m venv .typing-venv', wheels)
        self.assertIn("python-docx==1.2.0", wheels)
        self.assertIn("python-pptx==1.0.2", wheels)
        self.assertIn("test_python_docx_parity.py", wheels)
        self.assertIn("test_documented_examples.py", wheels)
        self.assertIn('"$typing_python" -m mypy --strict', wheels)
        self.assertIn('"$typing_python" -m mypy.stubtest rdocx\n', wheels)
        self.assertIn('"$typing_python" -m mypy.stubtest rpptx\n', wheels)
        self.assertNotIn("mypy.stubtest rdocx rdocx.", wheels)
        self.assertNotIn("mypy.stubtest rpptx rpptx.", wheels)

        sdist_matrix = self.yaml_block(build_sdists, "      matrix:")
        sdist_axes = tuple(
            line.split(":", 1)[0]
            for line in self.yaml_direct_lines(sdist_matrix, 8)
        )
        self.assertEqual(sdist_axes, ("package",))
        sdist_packages = self.yaml_block(sdist_matrix, "        package:")
        self.assertEqual(
            self.yaml_direct_lines(sdist_packages, 10), expected_package_entries
        )
        sdist_build = self.yaml_step(build_sdists, "Build source distribution")
        self.assertFalse(
            any(
                line.startswith("if:")
                for line in self.yaml_direct_lines(sdist_build, 8)
            )
        )
        sdist_build_inputs = self.yaml_block(sdist_build, "        with:")
        self.assertEqual(
            self.yaml_direct_lines(sdist_build_inputs, 10),
            (
                "command: sdist",
                "maturin-version: v1.13.3",
                "working-directory: crates/${{ matrix.package.crate }}",
                "args: --out ../../dist",
            ),
        )
        upload_sdist = self.yaml_step(build_sdists, "Upload source distribution")
        self.assertFalse(
            any(
                line.startswith("if:")
                for line in self.yaml_direct_lines(upload_sdist, 8)
            )
        )
        self.assertIn(
            "uses: actions/upload-artifact@"
            "ea165f8d65b6e75b540449e92b4886f43607fa02",
            upload_sdist,
        )
        upload_sdist_inputs = self.yaml_block(upload_sdist, "        with:")
        self.assertEqual(
            self.yaml_direct_lines(upload_sdist_inputs, 10),
            (
                "name: artifact-sdist-${{ matrix.package.distribution }}",
                "path: dist/*.tar.gz",
                "if-no-files-found: error",
            ),
        )
        self.assertIn("artifact-sdist-${{ matrix.package.distribution }}", wheels)
        self.assertIn("artifact-wheel-${{ matrix.package.distribution }}-", wheels)
        self.assertIn("pattern: artifact-*", wheels)
        self.assertIn("merge-multiple: true", wheels)
        self.assertIn("assert len(wheels) == 12", wheels)
        self.assertIn("assert len(sdists) == 2", wheels)

        publish_conditions = tuple(
            line.strip()
            for line in publish.splitlines()[1:]
            if len(line) - len(line.lstrip()) == 4
            and line.strip().startswith("if:")
        )
        self.assertEqual(
            publish_conditions,
            (
                "if: github.event_name == 'push' && "
                "startsWith(github.ref, 'refs/tags/py-v')",
            ),
        )
        publish_header = publish[: publish.index("    steps:\n")]
        publish_needs = tuple(
            line.removeprefix("needs: ").split(" #", 1)[0].rstrip()
            for line in self.yaml_direct_lines(publish_header, 4)
            if line.startswith("needs:")
        )
        self.assertEqual(publish_needs, ("[build-wheels, build-sdists]",))
        environments = tuple(
            line.removeprefix("environment: ")
            for line in self.yaml_direct_lines(publish_header, 4)
            if line.startswith("environment:")
        )
        self.assertEqual(environments, ("pypi",))
        publish_permissions = self.yaml_block(publish_header, "    permissions:")
        self.assertEqual(
            self.yaml_direct_lines(publish_permissions, 6),
            ("contents: read", "id-token: write"),
        )
        self.assertNotIn(
            "id-token: write", wheels[: wheels.index("  publish:\n")]
        )
        publication_validation = self.yaml_step(
            publish, "Validate complete publication set"
        )
        publication_download = self.yaml_step(
            publish, "Download all distributions"
        )
        publication_action = self.yaml_step(
            publish, "Publish to PyPI with trusted publishing"
        )
        publish_steps = self.yaml_block(publish, "    steps:")
        self.assertEqual(
            tuple(
                line
                for line in self.yaml_direct_lines(publish_steps, 6)
                if line.startswith("-")
            ),
            (
                "- name: Download all distributions",
                "- name: Validate complete publication set",
                "- name: Publish to PyPI with trusted publishing",
            ),
        )
        download_uses = tuple(
            line.removeprefix("uses: ").split(" #", 1)[0]
            for line in self.yaml_direct_lines(publication_download, 8)
            if line.startswith("uses:")
        )
        self.assertEqual(
            download_uses,
            (
                "actions/download-artifact@"
                "d3f86a106a0bac45b974a628896c90dbdf5c8093",
            ),
        )
        download_inputs = self.yaml_block(publication_download, "        with:")
        self.assertEqual(
            self.yaml_direct_lines(download_inputs, 10),
            ("path: dist", "pattern: artifact-*", "merge-multiple: true"),
        )
        validation_lines = self.yaml_run_lines(publication_validation)
        self.assert_no_success_short_circuit(validation_lines)
        self.assertEqual(
            validation_lines,
            (
                "python - <<'PY'",
                "from pathlib import Path",
                'wheels = list(Path("dist").glob("*.whl"))',
                'sdists = list(Path("dist").glob("*.tar.gz"))',
                "assert len(wheels) == 12, wheels",
                "assert len(sdists) == 2, sdists",
                "PY",
            ),
        )
        publication_uses = tuple(
            line.removeprefix("uses: ").split(" #", 1)[0]
            for line in self.yaml_direct_lines(publication_action, 8)
            if line.startswith("uses:")
        )
        self.assertEqual(
            publication_uses,
            (
                "pypa/gh-action-pypi-publish@"
                "cef221092ed1bacb1cc03d23a2d87d1d172e277b",
            ),
        )
        publication_inputs = self.yaml_block(publication_action, "        with:")
        self.assertEqual(
            self.yaml_direct_lines(publication_inputs, 10),
            ("packages-dir: dist",),
        )
        for step in (publication_validation, publication_action):
            self.assertFalse(
                any(
                    line.startswith("if:")
                    for line in self.yaml_direct_lines(step, 8)
                )
            )
        self.assertNotIn("continue-on-error:", publication_validation)
        self.assertNotIn("continue-on-error:", publication_action)

        action_uses = []
        for job_name, job in (
            ("build-wheels", build_wheels),
            ("build-sdists", build_sdists),
            ("publish", publish),
        ):
            for position, step in enumerate(self.yaml_steps(job), start=1):
                identity = self.yaml_step_identity(step, position)
                for action in self.yaml_step_actions(step):
                    action_uses.append((job_name, identity, action))
        self.assertEqual(
            tuple(action_uses),
            (
                (
                    "build-wheels",
                    "step:1",
                    "actions/checkout@de0fac2e4500dabe0009e67214ff5f5447ce83dd",
                ),
                (
                    "build-wheels",
                    "id:cp39",
                    "actions/setup-python@"
                    "a309ff8b426b58ec0e2a45f0f869d46889d02405",
                ),
                (
                    "build-wheels",
                    "Build cp39-abi3 wheel",
                    "PyO3/maturin-action@"
                    "86b9d133d34bc1b40018696f782949dac11bd380",
                ),
                (
                    "build-wheels",
                    "id:cp312",
                    "actions/setup-python@"
                    "a309ff8b426b58ec0e2a45f0f869d46889d02405",
                ),
                (
                    "build-wheels",
                    "Upload wheel",
                    "actions/upload-artifact@"
                    "ea165f8d65b6e75b540449e92b4886f43607fa02",
                ),
                (
                    "build-sdists",
                    "step:1",
                    "actions/checkout@de0fac2e4500dabe0009e67214ff5f5447ce83dd",
                ),
                (
                    "build-sdists",
                    "step:2",
                    "actions/setup-python@"
                    "a309ff8b426b58ec0e2a45f0f869d46889d02405",
                ),
                (
                    "build-sdists",
                    "Build source distribution",
                    "PyO3/maturin-action@"
                    "86b9d133d34bc1b40018696f782949dac11bd380",
                ),
                (
                    "build-sdists",
                    "Upload source distribution",
                    "actions/upload-artifact@"
                    "ea165f8d65b6e75b540449e92b4886f43607fa02",
                ),
                (
                    "publish",
                    "Download all distributions",
                    "actions/download-artifact@"
                    "d3f86a106a0bac45b974a628896c90dbdf5c8093",
                ),
                (
                    "publish",
                    "Publish to PyPI with trusted publishing",
                    "pypa/gh-action-pypi-publish@"
                    "cef221092ed1bacb1cc03d23a2d87d1d172e277b",
                ),
            ),
        )

    def test_wheels_workflow_covers_every_package_target_and_clean_install(
        self,
    ) -> None:
        workflow_bytes = (
            workflow.REPO / ".github/workflows/wheels.yml"
        ).read_bytes()
        self.assert_wheels_workflow_contract(workflow_bytes)

    def test_wheels_workflow_rejects_matrix_and_security_mutations(self) -> None:
        workflow_bytes = (
            workflow.REPO / ".github/workflows/wheels.yml"
        ).read_bytes()
        self.assert_wheels_workflow_contract(workflow_bytes)
        wheels = workflow_bytes.decode("utf-8", errors="strict")
        native_condition = "if: matrix.platform.install == 'native'"
        sdist_start = wheels.index("  build-sdists:\n")
        sdist_head = wheels[:sdist_start]
        sdist_tail = wheels[sdist_start:]
        publish_start = wheels.index("  publish:\n")
        publish_head = wheels[:publish_start]
        publish_tail = wheels[publish_start:]
        publication_download = self.yaml_step(
            publish_tail, "Download all distributions"
        )
        publication_validation = self.yaml_step(
            publish_tail, "Validate complete publication set"
        )
        publication_action = self.yaml_step(
            publish_tail, "Publish to PyPI with trusted publishing"
        )
        cp39_setup = self.yaml_block(wheels, "      - id: cp39")
        cp312_setup = self.yaml_block(wheels, "      - id: cp312")
        critical_run_steps = (
            "Validate wheel metadata",
            "Install and test native wheel",
            "Validate installed typing surface",
            "Install and test musllinux wheel",
            "Validate complete publication set",
        )
        musllinux_step = self.yaml_step(
            wheels, "Install and test musllinux wheel"
        )
        musllinux_parity_start = musllinux_step.index(
            '              if [ "$PACKAGE_DISTRIBUTION" = rdocx ]; then\n'
        )
        musllinux_parity_end = musllinux_step.index(
            "            '\n", musllinux_parity_start
        )
        musllinux_import_only_step = (
            musllinux_step[:musllinux_parity_start]
            + musllinux_step[musllinux_parity_end:]
        )
        early_success_mutations = tuple(
            (
                f"{name.lower().replace(' ', '-')}-{command.replace(' ', '-')}",
                wheels.replace(
                    step,
                    step.replace(
                        "        run: |\n",
                        f"        run: |\n          {command}\n",
                        1,
                    ),
                    1,
                ),
            )
            for name in critical_run_steps
            for command in ("exit 0", "return 0", "true")
            for step in (self.yaml_step(wheels, name),)
        )

        def mutate_run(
            name: str,
            *,
            prefix: str = "",
            suffix: str = "",
            first_line_suffix: str = "",
        ) -> str:
            step = self.yaml_step(wheels, name)
            mutated_step = step
            marker = "        run: |\n"
            if first_line_suffix:
                marker_end = mutated_step.index(marker) + len(marker)
                line_end = mutated_step.index("\n", marker_end)
                first_line = mutated_step[marker_end:line_end]
                mutated_step = (
                    mutated_step[:marker_end]
                    + first_line
                    + first_line_suffix
                    + mutated_step[line_end:]
                )
            if prefix:
                mutated_step = mutated_step.replace(
                    marker, marker + f"          {prefix}\n", 1
                )
            if suffix:
                mutated_step += f"          {suffix}\n"
            return wheels.replace(step, mutated_step, 1)

        control_flow_mutations = tuple(
            (f"{name.lower().replace(' ', '-')}-{label}", mutation)
            for name in critical_run_steps
            for label, mutation in (
                (
                    "if-false-wrapper",
                    mutate_run(name, prefix="if false; then", suffix="fi"),
                ),
                (
                    "if-true-wrapper",
                    mutate_run(name, prefix="if true; then", suffix="fi"),
                ),
                (
                    "set-plus-e-trailing-noop",
                    mutate_run(name, prefix="set +e", suffix=":"),
                ),
                (
                    "or-true",
                    mutate_run(name, first_line_suffix=" || true"),
                ),
                (
                    "or-noop",
                    mutate_run(name, first_line_suffix=" || :"),
                ),
                (
                    "semicolon-noop",
                    mutate_run(name, first_line_suffix="; :"),
                ),
                (
                    "semicolon-true",
                    mutate_run(name, first_line_suffix="; true"),
                ),
                ("trailing-noop", mutate_run(name, suffix=":")),
            )
        )
        duplicate_rdocx_sdist = sdist_head + sdist_tail.replace(
            "          - { distribution: rpptx, crate: rpptx-py, "
            "module: rpptx }",
            "          - { distribution: rdocx, crate: rdocx-py, "
            "module: rdocx }",
            1,
        )
        extra_sdist_axis = sdist_head + sdist_tail.replace(
            "      matrix:\n        package:\n",
            "      matrix:\n        python: [3.9, 3.12]\n        package:\n",
            1,
        )
        mutations = (
            (
                "missing-platform",
                wheels.replace(
                    "label: windows-x86_64", "label: windows-missing", 1
                ),
            ),
            (
                "missing-package",
                wheels.replace("distribution: rpptx", "distribution: absent", 1),
            ),
            (
                "missing-install",
                wheels.replace("-m venv .wheel-venv", "-m pip --version", 1),
            ),
            (
                "native-pytest-collect-only-environment",
                wheels.replace(
                    "      - name: Install and test native wheel\n",
                    "      - name: Install and test native wheel\n"
                    "        env:\n"
                    "          PYTEST_ADDOPTS: --collect-only\n",
                    1,
                ),
            ),
            (
                "typing-mypy-config-environment",
                wheels.replace(
                    "      - name: Validate installed typing surface\n",
                    "      - name: Validate installed typing surface\n"
                    "        env:\n"
                    "          MYPY_CONFIG_FILE: /dev/null\n",
                    1,
                ),
            ),
            (
                "wheel-checkout-ref-input",
                wheels.replace(
                    "      - uses: actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n",
                    "      - uses: actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n"
                    "        with: { ref: main }\n",
                    1,
                ),
            ),
            (
                "wheel-checkout-repository-input",
                wheels.replace(
                    "      - uses: actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n",
                    "      - uses: actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n"
                    "        with:\n"
                    "          repository: example/other\n"
                    "          ref: main\n",
                    1,
                ),
            ),
            (
                "sdist-checkout-ref-input",
                sdist_head
                + sdist_tail.replace(
                    "      - uses: actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n",
                    "      - uses: actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n"
                    "        with: { ref: main }\n",
                    1,
                ),
            ),
            (
                "sdist-checkout-repository-input",
                sdist_head
                + sdist_tail.replace(
                    "      - uses: actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n",
                    "      - uses: actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd # v6.0.2\n"
                    "        with:\n"
                    "          repository: example/other\n"
                    "          ref: main\n",
                    1,
                ),
            ),
            (
                "cp39-wrong-version-with-required-comment",
                wheels.replace(
                    '          python-version: "3.9"\n',
                    '          python-version: "3.12" # '
                    'python-version: "3.9"\n',
                    1,
                ),
            ),
            (
                "cp312-wrong-version-with-required-comment",
                wheels.replace(
                    '          python-version: "3.12"\n',
                    '          python-version: "3.9" # '
                    'python-version: "3.12"\n',
                    1,
                ),
            ),
            (
                "sdist-wrong-python-version",
                sdist_head
                + sdist_tail.replace(
                    '          python-version: "3.9"\n',
                    '          python-version: "3.8"\n',
                    1,
                ),
            ),
            (
                "sdist-wrong-version-with-required-comment",
                sdist_head
                + sdist_tail.replace(
                    '          python-version: "3.9"\n',
                    '          python-version: "3.12" # '
                    'python-version: "3.9"\n',
                    1,
                ),
            ),
            (
                "renamed-cp39-id",
                wheels.replace("      - id: cp39\n", "      - id: py39\n", 1),
            ),
            (
                "renamed-cp312-id",
                wheels.replace(
                    "      - id: cp312\n", "      - id: py312\n", 1
                ),
            ),
            (
                "cp39-package-restricted",
                wheels.replace(
                    cp39_setup,
                    cp39_setup.replace(
                        "      - id: cp39\n",
                        "      - id: cp39\n"
                        "        if: matrix.package.distribution == 'rdocx'\n",
                        1,
                    ),
                    1,
                ),
            ),
            (
                "cp312-unconditional",
                wheels.replace(
                    cp312_setup,
                    cp312_setup.replace(
                        "        if: matrix.platform.install == 'native'\n",
                        "",
                        1,
                    ),
                    1,
                ),
            ),
            (
                "sdist-setup-wrong-position",
                sdist_head
                + sdist_tail.replace(
                    "      - uses: actions/setup-python@",
                    "      - name: Setup prelude\n"
                    "        run: echo setup\n"
                    "      - uses: actions/setup-python@",
                    1,
                ),
            ),
            (
                "cp39-setup-wrong-position",
                wheels.replace(cp39_setup, "", 1).replace(
                    cp312_setup, cp39_setup + cp312_setup, 1
                ),
            ),
            (
                "missing-artifact-dependency",
                wheels.replace(
                    "needs: [build-wheels, build-sdists]",
                    "needs: build-wheels",
                    1,
                ),
            ),
            (
                "missing-sdist-need-preserved-in-comment",
                wheels.replace(
                    "needs: [build-wheels, build-sdists]",
                    "needs: build-wheels # "
                    "needs: [build-wheels, build-sdists]",
                    1,
                ),
            ),
            (
                "missing-wheel-need-preserved-in-comment",
                wheels.replace(
                    "needs: [build-wheels, build-sdists]",
                    "needs: build-sdists # "
                    "needs: [build-wheels, build-sdists]",
                    1,
                ),
            ),
            (
                "extra-publish-need",
                wheels.replace(
                    "needs: [build-wheels, build-sdists]",
                    "needs: [build-wheels, build-sdists, audit]",
                    1,
                ),
            ),
            (
                "reversed-publish-needs",
                wheels.replace(
                    "needs: [build-wheels, build-sdists]",
                    "needs: [build-sdists, build-wheels]",
                    1,
                ),
            ),
            (
                "wrong-tag-prefix",
                wheels.replace(
                    "startsWith(github.ref, 'refs/tags/py-v')",
                    "startsWith(github.ref, 'refs/heads/')",
                    1,
                ),
            ),
            (
                "tag-ignore-with-required-tag-comment",
                wheels.replace(
                    '    tags: ["py-v*"]',
                    '    tags-ignore: ["py-v*"] # tags: ["py-v*"]',
                    1,
                ),
            ),
            (
                "commented-push-trigger",
                wheels.replace("  push:\n", "  # push:\n", 1),
            ),
            (
                "comment-only-tag-filter",
                wheels.replace(
                    '    tags: ["py-v*"]', '    # tags: ["py-v*"]', 1
                ),
            ),
            (
                "extra-schedule-trigger",
                wheels.replace(
                    "  workflow_dispatch:\n",
                    "  workflow_dispatch:\n  schedule: []\n",
                    1,
                ),
            ),
            (
                "commented-workflow-dispatch",
                wheels.replace(
                    "  workflow_dispatch:\n", "  # workflow_dispatch:\n", 1
                ),
            ),
            (
                "extra-matrix-axis",
                wheels.replace(
                    "        platform:\n",
                    "        python: [3.9, 3.12]\n        platform:\n",
                    1,
                ),
            ),
            (
                "rdocx-only-native-gates",
                wheels.replace(
                    native_condition,
                    native_condition
                    + " && matrix.package.distribution == 'rdocx'",
                ),
            ),
            (
                "manual-dispatch-publication",
                wheels.replace(
                    "startsWith(github.ref, 'refs/tags/py-v')",
                    "startsWith(github.ref, 'refs/tags/py-v') || "
                    "github.event_name == 'workflow_dispatch'",
                    1,
                ),
            ),
            (
                "root-write-permission",
                wheels.replace("  contents: read", "  contents: write", 1),
            ),
            (
                "build-write-all",
                wheels.replace(
                    "    name: ${{ matrix.package.distribution }} "
                    "${{ matrix.platform.label }}\n",
                    "    name: ${{ matrix.package.distribution }} "
                    "${{ matrix.platform.label }}\n"
                    "    permissions: write-all\n",
                    1,
                ),
            ),
            (
                "sdist-write-all",
                wheels.replace(
                    "    name: ${{ matrix.package.distribution }} sdist\n",
                    "    name: ${{ matrix.package.distribution }} sdist\n"
                    "    permissions: write-all\n",
                    1,
                ),
            ),
            (
                "publish-contents-write",
                publish_head
                + publish_tail.replace(
                    "      contents: read", "      contents: write", 1
                ),
            ),
            (
                "publish-id-token-read",
                publish_head
                + publish_tail.replace(
                    "      id-token: write", "      id-token: read", 1
                ),
            ),
            (
                "publish-extra-permission",
                publish_head
                + publish_tail.replace(
                    "      contents: read\n",
                    "      contents: read\n      issues: write\n",
                    1,
                ),
            ),
            (
                "publish-staging-environment",
                wheels.replace("    environment: pypi", "    environment: pypi-staging", 1),
            ),
            (
                "native-continue-on-error",
                wheels.replace(
                    "      - name: Install and test native wheel\n",
                    "      - name: Install and test native wheel\n"
                    "        continue-on-error: true\n",
                    1,
                ),
            ),
            (
                "typing-continue-on-error",
                wheels.replace(
                    "      - name: Validate installed typing surface\n",
                    "      - name: Validate installed typing surface\n"
                    "        continue-on-error: true\n",
                    1,
                ),
            ),
            (
                "musllinux-if-false",
                wheels.replace(
                    "if: matrix.platform.install == 'musl'", "if: false", 1
                ),
            ),
            (
                "musllinux-import-only",
                wheels.replace(
                    musllinux_step, musllinux_import_only_step, 1
                ),
            ),
            (
                "musllinux-package-restriction",
                wheels.replace(
                    "if: matrix.platform.install == 'musl'",
                    "if: matrix.platform.install == 'musl' && "
                    "matrix.package.distribution == 'rdocx'",
                    1,
                ),
            ),
            (
                "musllinux-or-native",
                wheels.replace(
                    "if: matrix.platform.install == 'musl'",
                    "if: matrix.platform.install == 'musl' || "
                    "matrix.platform.install == 'native'",
                    1,
                ),
            ),
            (
                "wheel-upload-continue-on-error",
                wheels.replace(
                    "      - name: Upload wheel\n",
                    "      - name: Upload wheel\n"
                    "        continue-on-error: true\n",
                    1,
                ),
            ),
            (
                "sdist-upload-continue-on-error",
                wheels.replace(
                    "      - name: Upload source distribution\n",
                    "      - name: Upload source distribution\n"
                    "        continue-on-error: true\n",
                    1,
                ),
            ),
            (
                "wheel-upload-warn-with-error-comment",
                wheels.replace(
                    "          if-no-files-found: error\n",
                    "          if-no-files-found: warn # "
                    "if-no-files-found: error\n",
                    1,
                ),
            ),
            (
                "sdist-upload-warn-with-error-comment",
                sdist_head
                + sdist_tail.replace(
                    "          if-no-files-found: error\n",
                    "          if-no-files-found: warn # "
                    "if-no-files-found: error\n",
                    1,
                ),
            ),
            (
                "wheel-upload-policy-only-in-comment",
                wheels.replace(
                    "          if-no-files-found: error\n",
                    "          # if-no-files-found: error\n",
                    1,
                ),
            ),
            (
                "sdist-upload-policy-only-in-comment",
                sdist_head
                + sdist_tail.replace(
                    "          if-no-files-found: error\n",
                    "          # if-no-files-found: error\n",
                    1,
                ),
            ),
            (
                "publication-validation-continue-on-error",
                wheels.replace(
                    "      - name: Validate complete publication set\n",
                    "      - name: Validate complete publication set\n"
                    "        continue-on-error: true\n",
                    1,
                ),
            ),
            (
                "publication-action-continue-on-error",
                wheels.replace(
                    "      - name: Publish to PyPI with trusted publishing\n",
                    "      - name: Publish to PyPI with trusted publishing\n"
                    "        continue-on-error: true\n",
                    1,
                ),
            ),
            (
                "publication-validation-if-false",
                wheels.replace(
                    "      - name: Validate complete publication set\n",
                    "      - name: Validate complete publication set\n"
                    "        if: false\n",
                    1,
                ),
            ),
            (
                "publication-validation-if-always",
                wheels.replace(
                    "      - name: Validate complete publication set\n",
                    "      - name: Validate complete publication set\n"
                    "        if: always()\n",
                    1,
                ),
            ),
            (
                "publication-action-if-false",
                wheels.replace(
                    "      - name: Publish to PyPI with trusted publishing\n",
                    "      - name: Publish to PyPI with trusted publishing\n"
                    "        if: false\n",
                    1,
                ),
            ),
            (
                "publication-action-if-always",
                wheels.replace(
                    "      - name: Publish to PyPI with trusted publishing\n",
                    "      - name: Publish to PyPI with trusted publishing\n"
                    "        if: always()\n",
                    1,
                ),
            ),
            (
                "publish-before-validation",
                wheels.replace(
                    publication_validation + publication_action,
                    publication_action + publication_validation,
                    1,
                ),
            ),
            (
                "validation-before-download",
                wheels.replace(
                    publication_download + publication_validation,
                    publication_validation + publication_download,
                    1,
                ),
            ),
            (
                "download-other-path",
                wheels.replace("          path: dist\n", "          path: other\n", 1),
            ),
            (
                "download-specific-pattern",
                wheels.replace(
                    "          pattern: artifact-*\n",
                    "          pattern: artifact-wheel-*\n",
                    1,
                ),
            ),
            (
                "download-no-merge",
                wheels.replace(
                    "          merge-multiple: true\n",
                    "          merge-multiple: false\n",
                    1,
                ),
            ),
            (
                "narrow-wheel-validation-glob",
                wheels.replace(
                    'glob("*.whl")', 'glob("rdocx-*.whl")', 1
                ),
            ),
            (
                "narrow-sdist-validation-glob",
                wheels.replace(
                    'glob("*.tar.gz")', 'glob("rdocx-*.tar.gz")', 1
                ),
            ),
            (
                "wheel-count-preserved-only-in-comment",
                wheels.replace(
                    "          assert len(wheels) == 12, wheels\n",
                    "          assert len(wheels) == 1, wheels\n"
                    "          # assert len(wheels) == 12, wheels\n",
                    1,
                ),
            ),
            (
                "sdist-count-preserved-only-in-comment",
                wheels.replace(
                    "          assert len(sdists) == 2, sdists\n",
                    "          assert len(sdists) == 1, sdists\n"
                    "          # assert len(sdists) == 2, sdists\n",
                    1,
                ),
            ),
            (
                "publish-empty-input",
                wheels.replace(
                    "          packages-dir: dist\n",
                    "          packages-dir: empty\n",
                    1,
                ),
            ),
            (
                "rdocx-only-wheel-upload",
                wheels.replace(
                    "      - name: Upload wheel\n",
                    "      - name: Upload wheel\n"
                    "        if: matrix.package.distribution == 'rdocx'\n",
                    1,
                ),
            ),
            (
                "rdocx-only-sdist-upload",
                wheels.replace(
                    "      - name: Upload source distribution\n",
                    "      - name: Upload source distribution\n"
                    "        if: matrix.package.distribution == 'rdocx'\n",
                    1,
                ),
            ),
            (
                "push-only-wheel-build",
                wheels.replace(
                    "  build-wheels:\n",
                    "  build-wheels:\n    if: github.event_name == 'push'\n",
                    1,
                ),
            ),
            (
                "push-only-sdist-build",
                wheels.replace(
                    "  build-sdists:\n",
                    "  build-sdists:\n    if: github.event_name == 'push'\n",
                    1,
                ),
            ),
            (
                "second-publication-job",
                wheels
                + "\n  publish-copy:\n"
                + "    runs-on: ubuntu-24.04\n"
                + "    steps: []\n",
            ),
            (
                "publication-named-build-job",
                wheels.replace(
                    "    name: ${{ matrix.package.distribution }} "
                    "${{ matrix.platform.label }}",
                    "    name: Publish ${{ matrix.package.distribution }} "
                    "${{ matrix.platform.label }}",
                    1,
                ),
            ),
            ("missing-rpptx-sdist", duplicate_rdocx_sdist),
            ("extra-sdist-axis", extra_sdist_axis),
            (
                "wrong-wheel-artifact-path",
                wheels.replace("path: dist/*.whl", "path: dist/rdocx-*.whl", 1),
            ),
            (
                "wrong-wheel-artifact-name",
                wheels.replace(
                    "name: artifact-wheel-${{ matrix.package.distribution }}-"
                    "${{ matrix.platform.label }}",
                    "name: artifact-wheel-rdocx-${{ matrix.platform.label }}",
                    1,
                ),
            ),
            (
                "wrong-sdist-artifact-name",
                wheels.replace(
                    "name: artifact-sdist-${{ matrix.package.distribution }}",
                    "name: artifact-sdist-rdocx",
                    1,
                ),
            ),
            (
                "wrong-sdist-artifact-path",
                wheels.replace(
                    "path: dist/*.tar.gz", "path: dist/rdocx-*.tar.gz", 1
                ),
            ),
            (
                "wheel-fixed-target-with-expression-comment",
                wheels.replace(
                    "          target: ${{ matrix.platform.target }}\n",
                    "          target: x86_64-unknown-linux-gnu # "
                    "target: ${{ matrix.platform.target }}\n",
                    1,
                ),
            ),
            (
                "wheel-fixed-manylinux-with-expression-comment",
                wheels.replace(
                    "          manylinux: ${{ matrix.platform.manylinux }}\n",
                    "          manylinux: off # "
                    "manylinux: ${{ matrix.platform.manylinux }}\n",
                    1,
                ),
            ),
            (
                "wheel-fixed-package-with-expression-comment",
                wheels.replace(
                    "          working-directory: "
                    "crates/${{ matrix.package.crate }}\n",
                    "          working-directory: crates/rdocx-py # "
                    "working-directory: crates/${{ matrix.package.crate }}\n",
                    1,
                ),
            ),
            (
                "wheel-weakened-args-with-required-comment",
                wheels.replace(
                    "          args: --release --locked --compatibility pypi "
                    "--out ../../dist\n",
                    "          args: --release --out ../../dist # "
                    "--locked --compatibility pypi\n",
                    1,
                ),
            ),
            (
                "wheel-wrong-command-with-build-comment",
                wheels.replace(
                    "          command: build\n",
                    "          command: sdist # command: build\n",
                    1,
                ),
            ),
            (
                "wheel-wrong-maturin-version",
                wheels.replace(
                    "          maturin-version: v1.13.3\n",
                    "          maturin-version: v1.13.2\n",
                    1,
                ),
            ),
            (
                "wheel-package-restricted-build",
                wheels.replace(
                    "      - name: Build cp39-abi3 wheel\n",
                    "      - name: Build cp39-abi3 wheel\n"
                    "        if: matrix.package.distribution == 'rdocx'\n",
                    1,
                ),
            ),
            (
                "sdist-fixed-package",
                sdist_head
                + sdist_tail.replace(
                    "          working-directory: "
                    "crates/${{ matrix.package.crate }}\n",
                    "          working-directory: crates/rdocx-py\n",
                    1,
                ),
            ),
            (
                "sdist-fixed-package-with-expression-comment",
                sdist_head
                + sdist_tail.replace(
                    "          working-directory: "
                    "crates/${{ matrix.package.crate }}\n",
                    "          working-directory: crates/rdocx-py # "
                    "working-directory: crates/${{ matrix.package.crate }}\n",
                    1,
                ),
            ),
            (
                "sdist-wrong-command-with-sdist-comment",
                sdist_head
                + sdist_tail.replace(
                    "          command: sdist\n",
                    "          command: build # command: sdist\n",
                    1,
                ),
            ),
            (
                "sdist-wrong-maturin-version",
                sdist_head
                + sdist_tail.replace(
                    "          maturin-version: v1.13.3\n",
                    "          maturin-version: v1.13.2\n",
                    1,
                ),
            ),
            (
                "sdist-wrong-args-with-required-comment",
                sdist_head
                + sdist_tail.replace(
                    "          args: --out ../../dist\n",
                    "          args: --out dist # args: --out ../../dist\n",
                    1,
                ),
            ),
            (
                "sdist-package-restricted-build",
                wheels.replace(
                    "      - name: Build source distribution\n",
                    "      - name: Build source distribution\n"
                    "        if: matrix.package.distribution == 'rdocx'\n",
                    1,
                ),
            ),
            (
                "wheel-checkout-unreviewed-sha",
                wheels.replace(
                    "actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd",
                    "actions/checkout@1111111111111111111111111111111111111111 "
                    "# actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd",
                    1,
                ),
            ),
            (
                "wheel-setup-unreviewed-sha",
                wheels.replace(
                    "actions/setup-python@"
                    "a309ff8b426b58ec0e2a45f0f869d46889d02405",
                    "actions/setup-python@2222222222222222222222222222222222222222 "
                    "# actions/setup-python@"
                    "a309ff8b426b58ec0e2a45f0f869d46889d02405",
                    1,
                ),
            ),
            (
                "wheel-maturin-unreviewed-sha",
                wheels.replace(
                    "PyO3/maturin-action@"
                    "86b9d133d34bc1b40018696f782949dac11bd380",
                    "PyO3/maturin-action@"
                    "3333333333333333333333333333333333333333 "
                    "# PyO3/maturin-action@"
                    "86b9d133d34bc1b40018696f782949dac11bd380",
                    1,
                ),
            ),
            (
                "wheel-upload-unreviewed-sha",
                wheels.replace(
                    "actions/upload-artifact@"
                    "ea165f8d65b6e75b540449e92b4886f43607fa02",
                    "actions/upload-artifact@"
                    "4444444444444444444444444444444444444444 "
                    "# actions/upload-artifact@"
                    "ea165f8d65b6e75b540449e92b4886f43607fa02",
                    1,
                ),
            ),
            (
                "sdist-checkout-unreviewed-sha",
                sdist_head
                + sdist_tail.replace(
                    "actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd",
                    "actions/checkout@"
                    "6666666666666666666666666666666666666666 "
                    "# actions/checkout@"
                    "de0fac2e4500dabe0009e67214ff5f5447ce83dd",
                    1,
                ),
            ),
            (
                "sdist-maturin-unreviewed-sha",
                sdist_head
                + sdist_tail.replace(
                    "PyO3/maturin-action@"
                    "86b9d133d34bc1b40018696f782949dac11bd380",
                    "PyO3/maturin-action@"
                    "7777777777777777777777777777777777777777 "
                    "# PyO3/maturin-action@"
                    "86b9d133d34bc1b40018696f782949dac11bd380",
                    1,
                ),
            ),
            (
                "extra-unreviewed-action",
                wheels.replace(
                    "      - uses: actions/checkout@",
                    "      - uses: example/unknown@"
                    "5555555555555555555555555555555555555555\n"
                    "      - uses: actions/checkout@",
                    1,
                ),
            ),
        ) + early_success_mutations + control_flow_mutations + (
            (
                "crlf-workflow-bytes",
                workflow_bytes.replace(b"\n", b"\r\n"),
            ),
        )

        for name, mutated in mutations:
            mutated_bytes = (
                mutated if isinstance(mutated, bytes) else mutated.encode("utf-8")
            )
            self.assertNotEqual(mutated_bytes, workflow_bytes, name)
            with self.subTest(name=name):
                with self.assertRaises(AssertionError):
                    self.assert_wheels_workflow_contract(mutated_bytes)

    def assert_publish_preflight_contract(self, publish: str) -> None:
        publishable_crates = (
            "oxml-core",
            "oxml-drawing",
            "oxml-layout",
            "oxml-media",
            "oxml-opc",
            "oxml-pdf",
            "oxml-sml",
            "oxml-cli-support",
            "oxml-chart",
            "rdocx",
            "rdocx-cli",
            "rdocx-html",
            "rdocx-layout",
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx-pdf",
            "rpptx",
            "rpptx-cli",
            "rpptx-chart",
            "rpptx-layout",
            "rpptx-oxml",
            "rpptx-render",
        )
        marker = "      - name: Verify publication archives\n"
        self.assertEqual(publish.count(marker), 1)
        start = publish.index(marker)
        end = publish.index("\n      - name:", start + len(marker))
        block = publish[start:end]

        self.assertEqual(block.count("cargo publish --workspace --dry-run"), 1)
        for package in publishable_crates:
            config = (
                f"--config 'patch.crates-io.{package}.path=\"crates/{package}\"'"
            )
            self.assertEqual(block.count(config), 1, package)
        self.assertEqual(block.count("--config 'patch.crates-io."), 22)
        self.assertNotIn("--no-verify", block)
        self.assertNotIn("continue-on-error", block)

    def assert_publish_workflow_contract(self, publish: str) -> None:
        stable_crates = (
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx-layout",
            "rdocx-html",
            "rdocx-pdf",
            "rdocx",
            "rdocx-cli",
        )
        incubating_crates = (
            "oxml-core",
            "oxml-opc",
            "oxml-media",
            "oxml-layout",
            "oxml-drawing",
            "oxml-pdf",
            "oxml-sml",
            "oxml-cli-support",
            "oxml-chart",
            "rpptx-oxml",
            "rpptx-chart",
            "rpptx-layout",
            "rpptx-render",
            "rpptx",
            "rpptx-cli",
        )

        self.assertIn('tags: ["v*", "rpptx-v*"]', publish)
        for step_name, condition, packages in (
            (
                "Publish stable allowlist",
                "if: startsWith(github.ref_name, 'v')",
                stable_crates,
            ),
            (
                "Publish incubating allowlist",
                "if: startsWith(github.ref_name, 'rpptx-v')",
                incubating_crates,
            ),
        ):
            marker = f"      - name: {step_name}\n"
            self.assertEqual(publish.count(marker), 1)
            start = publish.index(marker)
            block_lines = []
            for line in publish[start:].splitlines():
                if block_lines and (
                    line.startswith("      - ")
                    or (line.strip() and len(line) - len(line.lstrip()) <= 4)
                ):
                    break
                block_lines.append(line)
            block = "\n".join(block_lines)

            conditions = [line.strip() for line in block_lines if "if:" in line]
            self.assertEqual(conditions, [condition])
            self.assertNotIn("continue-on-error", block)
            run_index = next(
                index
                for index, line in enumerate(block_lines)
                if line.strip() == "run: |"
            )
            commands = [
                line.strip()
                for line in block_lines[run_index + 1 :]
                if line.strip()
            ]
            expected_commands = []
            for index, package in enumerate(packages):
                expected_commands.append(f"cargo publish -p {package}")
                if index + 1 < len(packages):
                    expected_commands.append("sleep 60")
            self.assertEqual(commands, expected_commands)

            package_position = {name: index for index, name in enumerate(packages)}
            for name in packages:
                manifest = tomllib.loads(
                    (workflow.REPO / f"crates/{name}/Cargo.toml").read_text(
                        encoding="utf-8"
                    )
                )
                for dependency in manifest.get("dependencies", {}):
                    if dependency in package_position:
                        self.assertLess(
                            package_position[dependency],
                            package_position[name],
                            f"{dependency} must publish before {name}",
                        )

        actual_publish_commands = [
            line.strip()
            for line in publish.splitlines()
            if line.strip().startswith("cargo publish -p ")
        ]
        expected_publish_commands = [
            f"cargo publish -p {package}"
            for package in stable_crates + incubating_crates
        ]
        self.assertEqual(actual_publish_commands, expected_publish_commands)

    def assert_release_notes_publish_contract(self, publish: str) -> None:
        publish_job = self.yaml_block(publish, "  publish:")
        prepublish = self.yaml_step(publish_job, "Verify reviewed release notes")
        validation_command = (
            "run: python3 scripts/sprint_workflow.py release-notes "
            '"${{ github.ref_name }}" --check'
        )
        self.assertEqual(
            self.yaml_direct_lines(prepublish, 8),
            (validation_command,),
        )
        self.assert_no_success_short_circuit(self.operative_lines(prepublish))
        for publish_command in (
            "cargo publish -p rdocx-opc",
            "cargo publish -p oxml-core",
        ):
            self.assertLess(
                publish_job.index(prepublish),
                publish_job.index(publish_command),
                publish_command,
            )

        release = self.yaml_block(publish, "  release:")
        steps = self.yaml_steps(release)
        self.assertEqual(
            tuple(
                self.yaml_step_identity(step, index)
                for index, step in enumerate(steps, 1)
            ),
            ("step:1", "Create GitHub Release from reviewed notes"),
        )
        create = self.yaml_step(release, "Create GitHub Release from reviewed notes")
        self.assertEqual(
            self.yaml_direct_lines(create, 8),
            ("run: |", "env:"),
        )
        self.assertIn(
            "GH_TOKEN: ${{ secrets.GITHUB_TOKEN }}",
            self.yaml_direct_lines(create, 10),
        )
        commands = (
            "python3 scripts/sprint_workflow.py release-notes "
            '"${{ github.ref_name }}" --check',
            "python3 scripts/sprint_workflow.py release-notes "
            '"${{ github.ref_name }}" --render > '
            '"$RUNNER_TEMP/release-notes.md"',
            "python3 scripts/sprint_workflow.py release-notes "
            '"${{ github.ref_name }}" --render | cmp - '
            '"$RUNNER_TEMP/release-notes.md"',
            'gh release create "${{ github.ref_name }}" --notes-file '
            '"$RUNNER_TEMP/release-notes.md"',
        )
        self.assertEqual(self.yaml_run_lines(create), commands)
        self.assertEqual(create.count("python3 scripts/sprint_workflow.py"), 3)
        self.assertEqual(create.count('"${{ github.ref_name }}" --check'), 1)
        self.assertEqual(create.count("--render"), 2)
        self.assertEqual(create.count("cmp -"), 1)
        self.assertEqual(create.count("--notes-file"), 1)
        self.assertNotIn("--generate-notes", publish)
        self.assert_no_success_short_circuit(self.operative_lines(create))

    def test_preset_geometry_provenance_is_recorded(self) -> None:
        rendering = (workflow.REPO / "docs/hld/08-rendering-spec.md").read_text(
            encoding="utf-8"
        )
        risks = (
            workflow.REPO / "docs/hld/13-risks-and-open-questions.md"
        ).read_text(encoding="utf-8")
        decision = rendering + risks

        self.assertIn("ECMA-376-1_5th_edition_december_2016.zip", decision)
        self.assertIn(
            "OfficeOpenXML-DrawingMLGeometries.zip/presetShapeDefinitions.xml",
            decision,
        )
        self.assertIn("187 preset shape definitions", decision)
        self.assertIn(
            "2f7c868d857c1e3c4b5a6068759fe0e07d77ad58377a6618d1b02ba3507b6939",
            decision,
        )
        self.assertIn("Ecma software policy", decision)
        self.assertIn("three-clause BSD", decision)
        self.assertIn("retain the Ecma copyright notice", decision)

    def test_libreoffice_preset_table_remains_rejected(self) -> None:
        rendering = (workflow.REPO / "docs/hld/08-rendering-spec.md").read_text(
            encoding="utf-8"
        )
        risks = (
            workflow.REPO / "docs/hld/13-risks-and-open-questions.md"
        ).read_text(encoding="utf-8")
        decision = rendering + risks

        self.assertIn("LibreOffice's preset table must not be used", decision)
        self.assertIn("MPL-2.0 file-level copyleft", decision)

    def test_validation_only_sprint_initialises_without_wave_rows(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            current = root / "CURRENT_SPRINT.md"
            scratch = root / "scratch"
            current.write_text(
                "# Current Sprint, S11\n\n"
                "**Validation-only**: yes\n\n"
                "## The wave\n\n"
                "| F-ID | Title | Size | Status | Owner |\n"
                "|---|---|---|---|---|\n",
                encoding="utf-8",
            )
            args = argparse.Namespace(
                sprint="S11",
                resume=False,
                force=False,
                max_review_passes=3,
                max_workers=None,
            )

            with (
                patch.object(workflow, "CURRENT_SPRINT", current),
                patch.object(workflow, "SCRATCH", scratch),
            ):
                workflow.cmd_init(args)

            saved = json.loads((scratch / "S11-run.json").read_text(encoding="utf-8"))
            self.assertEqual(saved["features"], {})
            self.assertEqual(saved["phase"], "design")

    def test_init_resume_refreshes_feature_metadata_without_losing_progress(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            current = root / "CURRENT_SPRINT.md"
            scratch = root / "scratch"
            current.write_text(
                "# Current Sprint, S11\n\n"
                "| F-ID | Title | Size | Status | Owner |\n"
                "|---|---|---|---|---|\n"
                "| F-001 | Refreshed title | M | in-progress | claude |\n"
                "| F-002 | Newly added | S | pending | - |\n",
                encoding="utf-8",
            )
            existing = {
                "schema_version": workflow.SCHEMA_VERSION,
                "sprint": "S11",
                "phase": "implementation",
                "max_review_passes": 3,
                "max_workers": 2,
                "features": {
                    "F-001": {
                        "state": "reviewed",
                        "size": "S",
                        "title": "Original title",
                        "owner": "codex",
                        "wave": 2,
                        "branch": "work/f-001-codex",
                        "worktree": "/private/tmp/f-001",
                        "head": "abc123",
                        "handoff": "consumed",
                        "integration_commit": "def456",
                    }
                },
                "reviews": [{"pass": 1, "blocking": 0, "head": "def456"}],
                "verifications": [
                    {"scope": "full", "passed": True, "head": "def456"}
                ],
            }
            scratch.mkdir()
            (scratch / "S11-run.json").write_text(
                json.dumps(existing), encoding="utf-8"
            )
            args = argparse.Namespace(
                sprint="S11",
                resume=True,
                force=False,
                max_review_passes=4,
                max_workers=3,
            )

            with (
                patch.object(workflow, "CURRENT_SPRINT", current),
                patch.object(workflow, "SCRATCH", scratch),
            ):
                workflow.cmd_init(args)

            saved = json.loads((scratch / "S11-run.json").read_text(encoding="utf-8"))
            refreshed = saved["features"]["F-001"]
            self.assertEqual(refreshed["title"], "Refreshed title")
            self.assertEqual(refreshed["size"], "M")
            for field in (
                "state",
                "owner",
                "wave",
                "branch",
                "worktree",
                "head",
                "handoff",
                "integration_commit",
            ):
                self.assertEqual(refreshed[field], existing["features"]["F-001"][field])
            self.assertEqual(saved["reviews"], existing["reviews"])
            self.assertEqual(saved["verifications"], existing["verifications"])
            self.assertEqual(saved["phase"], "implementation")
            self.assertEqual(saved["max_review_passes"], 4)
            self.assertEqual(saved["max_workers"], 3)
            self.assertEqual(
                saved["features"]["F-002"],
                {
                    "state": "pending",
                    "size": "S",
                    "title": "Newly added",
                    "owner": None,
                },
            )

    def test_empty_sprint_without_validation_marker_is_refused(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            current = Path(directory) / "CURRENT_SPRINT.md"
            current.write_text(
                "# Current Sprint, S11\n\n"
                "| F-ID | Title | Size | Status | Owner |\n"
                "|---|---|---|---|---|\n",
                encoding="utf-8",
            )

            with patch.object(workflow, "CURRENT_SPRINT", current):
                with self.assertRaises(SystemExit):
                    workflow.parse_current_sprint()

    def test_workspace_release_versions_move_in_lockstep(self) -> None:
        root = tomllib.loads((workflow.REPO / "Cargo.toml").read_text(encoding="utf-8"))
        workspace = root["workspace"]
        version = workspace["package"]["version"]

        for name in (
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx",
            "rdocx-layout",
            "rdocx-pdf",
            "rdocx-html",
        ):
            self.assertEqual(workspace["dependencies"][name]["version"], version)

        wasm = tomllib.loads(
            (workflow.REPO / "crates/rdocx-wasm/Cargo.toml").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(wasm["package"]["version"], {"workspace": True})
        self.assertFalse(wasm["package"]["publish"])

    def assert_stable_collaboration_release_notes_contract(
        self, changelog: str
    ) -> None:
        notes = workflow.render_release_notes(changelog, "v0.8.0")
        added = notes.split("### Added\n\n", 1)[1].split("\n### Fixed", 1)[0]
        for claim in (
            "comments and threaded conversations",
            "content controls to namespace-aware custom XML",
            "bookmarks and resolve `REF` and `PAGEREF` cross-references",
            "Inspect tracked revisions",
            "accept or reject all or a filtered selection",
            "Render accepted or tracked revision views",
            "document-protection intent",
        ):
            self.assertIn(claim, added, claim)

        contributors = notes.split("### Contributors\n\n", 1)[1]
        for credit in (
            "@emptinessform",
            "Issue 37 complete-layout report",
            "Issue 39 relayout measurements",
        ):
            self.assertIn(credit, contributors, credit)

    def test_v0_8_0_notes_cover_native_collaboration_tranche(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        self.assert_stable_collaboration_release_notes_contract(changelog)

    def test_v0_8_0_notes_reject_an_omitted_collaboration_capability(
        self,
    ) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        for claim in (
            "comments and threaded conversations",
            "content controls to namespace-aware custom XML",
            "bookmarks and resolve `REF` and `PAGEREF` cross-references",
            "Inspect tracked revisions",
            "accept or reject all or a filtered selection",
            "Render accepted or tracked revision views",
            "document-protection intent",
        ):
            mutated = changelog.replace(claim, "native collaboration capability", 1)
            self.assertNotEqual(mutated, changelog, claim)
            with self.subTest(claim=claim), self.assertRaises(AssertionError):
                self.assert_stable_collaboration_release_notes_contract(mutated)

    def test_v0_8_0_notes_reject_an_omitted_issue_reporter_credit(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        stable_notes = workflow.render_release_notes(changelog, "v0.8.0")
        for credit in (
            "@emptinessform",
            "Issue 37 complete-layout report",
            "Issue 39 relayout measurements",
        ):
            mutated_notes = stable_notes.replace(credit, "external report", 1)
            self.assertNotEqual(mutated_notes, stable_notes, credit)
            mutated = changelog.replace(stable_notes, mutated_notes, 1)
            with self.subTest(credit=credit), self.assertRaises(AssertionError):
                self.assert_stable_collaboration_release_notes_contract(mutated)

    def assert_v0_10_1_release_notes_truth_contract(self, changelog: str) -> str:
        notes = workflow.render_release_notes(changelog, "v0.10.1")
        for record in (
            "https://github.com/tensorbee/rdocx/issues/44",
            "https://github.com/tensorbee/rdocx/pull/45",
            "https://github.com/tensorbee/rdocx/issues/46",
            "https://github.com/tensorbee/rdocx/pull/47",
            "https://github.com/tensorbee/rdocx/pull/48",
            "https://github.com/tensorbee/rdocx/pull/49",
            "https://github.com/tensorbee/rdocx/pull/50",
            "https://github.com/tensorbee/rdocx/pull/51",
            "https://github.com/tensorbee/rdocx/pull/52",
        ):
            self.assertEqual(notes.count(record), 2, record)
        for contributor in ("@emptinessform", "@pedroassumpcao"):
            self.assertGreaterEqual(notes.count(contributor), 1, contributor)

        highlights = notes.split("### Highlights\n\n", 1)[1].split(
            "\n### Added", 1
        )[0]
        highlight_paragraphs = [
            " ".join(paragraph.split()) for paragraph in highlights.split("\n\n")
        ]
        self.assertEqual(
            highlight_paragraphs[1],
            "This patch release recovers the complete stable family after v0.10.0 "
            "published only `rdocx-opc` and `rdocx-oxml`, then stopped during "
            "`rdocx-layout` package verification. Version 0.10.1 is the first "
            "complete stable family carrying the S56 outcome.",
        )

        compatibility = notes.split("### Compatibility\n\n", 1)[1].split(
            "\n### Contributors", 1
        )[0]
        compatibility_paragraphs = [
            " ".join(paragraph.split())
            for paragraph in compatibility.split("\n\n")
        ]
        self.assertEqual(
            compatibility_paragraphs[1],
            "The immutable v0.10.0 attempt contains only `rdocx-opc` and "
            "`rdocx-oxml`. Callers should select 0.10.1 for a coherent "
            "seven-package stable graph. No v0.10.0 tag or registry entry was "
            "moved, replaced, or reused.",
        )

        contributor_paragraphs = [
            " ".join(paragraph.split())
            for paragraph in notes.split("### Contributors\n\n", 1)[1].split(
                "\n\n"
            )
        ]
        self.assertEqual(
            contributor_paragraphs[0],
            "Atul Sharma maintained the release. Thanks to `@emptinessform` "
            "for the caller font-alias report and "
            "reference implementation in [Issue "
            "44](https://github.com/tensorbee/rdocx/issues/44) and [PR "
            "45](https://github.com/tensorbee/rdocx/pull/45), and for the editor "
            "performance measurements and migration evidence in [Issue "
            "46](https://github.com/tensorbee/rdocx/issues/46).",
        )
        self.assertEqual(
            contributor_paragraphs[1],
            "Thanks to `@pedroassumpcao` for the ordered cell, run, paragraph, "
            "and hyperlink reader designs in [PR "
            "47](https://github.com/tensorbee/rdocx/pull/47), [PR "
            "48](https://github.com/tensorbee/rdocx/pull/48), and [PR "
            "49](https://github.com/tensorbee/rdocx/pull/49), the unsupported "
            "XML facts in [PR 50](https://github.com/tensorbee/rdocx/pull/50), "
            "producer-defined numbering preservation in [PR "
            "51](https://github.com/tensorbee/rdocx/pull/51), and fail-closed "
            "text decoding in [PR 52](https://github.com/tensorbee/rdocx/pull/52).",
        )
        direct_classification = (
            "No named external patch landed directly. Each named report or "
            "proposal landed through the hardened equivalent described above so "
            "that current namespace, non-exhaustive API, bounded-allocation, "
            "diagnostic, and compatibility contracts remain intact."
        )
        self.assertEqual(contributor_paragraphs[2], direct_classification)
        self.assertEqual(" ".join(contributor_paragraphs).count("landed directly"), 1)
        return notes

    def assert_v0_11_1_release_notes_truth_contract(self, changelog: str) -> str:
        notes = workflow.render_release_notes(changelog, "v0.11.1")
        normalized_notes = " ".join(notes.split())
        records = (
            "https://github.com/tensorbee/rdocx/issues/53",
            "https://github.com/tensorbee/rdocx/issues/54",
            "https://github.com/tensorbee/rdocx/pull/55",
            "https://github.com/tensorbee/rdocx/pull/56",
            "https://github.com/tensorbee/rdocx/pull/57",
            "https://github.com/tensorbee/rdocx/pull/58",
        )
        for record in records:
            self.assertEqual(notes.count(record), 2, record)
        for contributor in ("@emptinessform", "@pedroassumpcao"):
            self.assertGreaterEqual(notes.count(contributor), 1, contributor)
        for source_sha in (
            "056d48fdf23f35e3538ef3d6ff78cf9e3863e3a5",
            "8b79c4cd0452defafe0a58e86b332c98e7fe52d7",
            "44498f042a2290ef40c7a6c26025f38e38e9ce2a",
            "c8fed1d1268fd765d602bac2da6524900c1c1cfd",
        ):
            self.assertEqual(notes.count(source_sha), 1, source_sha)
        for claim in (
            "700-paragraph note and header or footer workloads",
            "22 MiB caller-font workload",
            "whole-valued decimal table measurements",
            "tracked table-grid history",
            "legacy VML reader classification",
            "locked Word fidelity dependency preparation",
            "No named external patch landed directly",
            "Both issues remain open after their release-bound thank-yous",
            "All four pull requests remain open after their release-bound thank-yous",
        ):
            self.assertIn(claim, normalized_notes, claim)
        self.assertEqual(notes.count("`rdocx-opc@0.11.0`"), 2)
        self.assertEqual(notes.count("`rdocx-oxml@0.11.0`"), 2)
        for absent in (
            "rdocx-layout@0.11.0",
            "rdocx-html@0.11.0",
            "rdocx-pdf@0.11.0",
            "rdocx@0.11.0",
            "rdocx-cli@0.11.0",
        ):
            self.assertNotIn(absent, notes)
        self.assertNotIn("renders legacy VML horizontal rules", notes)
        self.assertIn(
            "callers constructing full `TextSegment` literals must initialize "
            "the `direction` field",
            normalized_notes,
        )
        return notes

    def assert_v0_12_0_release_notes_truth_contract(self, changelog: str) -> str:
        notes = workflow.render_release_notes(changelog, "v0.12.0")
        normalized_notes = " ".join(notes.split())
        records = (
            "https://github.com/tensorbee/rdocx/pull/61",
            "https://github.com/tensorbee/rdocx/pull/62",
            "https://github.com/tensorbee/rdocx/pull/63",
            "https://github.com/tensorbee/rdocx/pull/64",
            "https://github.com/tensorbee/rdocx/issues/65",
            "https://github.com/tensorbee/rdocx/issues/66",
            "https://github.com/tensorbee/rdocx/issues/67",
        )
        for record in records:
            self.assertEqual(notes.count(record), 2, record)
        for contributor in ("@pedroassumpcao", "@emptinessform"):
            self.assertGreaterEqual(notes.count(contributor), 1, contributor)
        for claim in (
            "relationship-safe hyperlink and drawing reader design",
            "document and table completeness design",
            "numbering and effective formatting design",
            "tracked insertion and field safety design",
            "note-reference cache report",
            "ordinary-prose restart report",
            "page-spanning paragraph regression report",
            "No named external patch landed directly",
            "Each contribution landed through a reviewed hardened equivalent",
            "The four pull requests and Issues 65 and 66 remain closed",
            "Issue 67 remains open",
        ):
            self.assertIn(claim, normalized_notes, claim)
        for package in (
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx-layout",
            "rdocx-html",
            "rdocx-pdf",
            "rdocx",
            "rdocx-cli",
        ):
            self.assertIn(f"`{package}`", notes, package)
        self.assertIn("shared OOXML 0.9.0 family", normalized_notes)
        self.assertIn("`rdocx-wasm@0.12.0` is not a crates.io package", normalized_notes)
        self.assertNotIn("rpptx-v0.9.0", notes)
        return notes

    def assert_v0_13_0_release_notes_truth_contract(self, changelog: str) -> str:
        notes = workflow.render_release_notes(changelog, "v0.13.0")
        normalized_notes = " ".join(notes.split())
        for claim in (
            "M22 Word-depth boundary",
            "OfficeMath",
            "MathML and LaTeX",
            "dynamic table-of-contents rebuilding",
            "sectioned mail merge",
            "VBA, OLE, and ActiveX",
            "DOCX, DOCM, DOTX, and DOTM",
            "Flat OPC",
            "MHTML",
            "shared `oxml-core` validator",
            "separately published shared OOXML 0.10.0 family",
            "additive pre-1.0 surfaces",
            "does not remove macros",
            "No external issue or pull request belongs to the selected stable-family",
        ):
            self.assertIn(claim, normalized_notes, claim)
        for package in (
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx-layout",
            "rdocx-html",
            "rdocx-pdf",
            "rdocx",
            "rdocx-cli",
        ):
            self.assertIn(f"`{package}`", notes, package)
        self.assertIn("`rdocx-wasm@0.13.0` is not a crates.io package", normalized_notes)
        self.assertIn("Atul Sharma maintained the release", normalized_notes)
        self.assertNotIn("github.com/tensorbee/rdocx/issues/", notes)
        self.assertNotIn("github.com/tensorbee/rdocx/pull/", notes)
        self.assertNotIn("rpptx-v0.10.0", notes)
        return notes

    def test_stable_release_family_is_prepared_at_0_13_0(self) -> None:
        expected_version = "0.13.0"
        stable_members = (
            "oxml-py-support",
            "rpptx-py",
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx",
            "rdocx-layout",
            "rdocx-pdf",
            "rdocx-html",
            "rdocx-py",
            "rdocx-cli",
            "rdocx-wasm",
        )
        stable_pins = (
            "oxml-py-support",
            "rpptx-py",
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx",
            "rdocx-layout",
            "rdocx-pdf",
            "rdocx-html",
            "rdocx-py",
        )
        stable_publishable = {
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx-layout",
            "rdocx-html",
            "rdocx-pdf",
            "rdocx",
            "rdocx-cli",
        }
        incubating_members = (
            "oxml-core",
            "oxml-drawing",
            "oxml-layout",
            "oxml-media",
            "oxml-opc",
            "oxml-pdf",
            "oxml-sml",
            "oxml-cli-support",
            "oxml-chart",
            "rpptx",
            "rpptx-cli",
            "rpptx-chart",
            "rpptx-layout",
            "rpptx-oxml",
            "rpptx-render",
            "rpptx-wasm",
        )

        root_text = (workflow.REPO / "Cargo.toml").read_text(encoding="utf-8")
        root = tomllib.loads(root_text)
        workspace = root["workspace"]
        self.assertEqual(workspace["package"]["version"], expected_version)
        dependencies = workspace["dependencies"]
        for name in stable_pins:
            self.assertEqual(dependencies[name]["version"], expected_version, name)

        lock = tomllib.loads((workflow.REPO / "Cargo.lock").read_text(encoding="utf-8"))
        lock_versions = {
            package["name"]: package["version"]
            for package in lock["package"]
            if package["name"] in stable_members
        }
        self.assertEqual(set(lock_versions), set(stable_members))
        for name in stable_members:
            self.assertEqual(lock_versions[name], expected_version, name)

        publishable = set()
        for name in stable_members:
            manifest = tomllib.loads(
                (workflow.REPO / f"crates/{name}/Cargo.toml").read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual(manifest["package"]["version"], {"workspace": True})
            if manifest["package"].get("publish", True):
                publishable.add(name)
        self.assertEqual(publishable, stable_publishable)

        for name in ("rdocx-py", "rpptx-py"):
            pyproject = tomllib.loads(
                (workflow.REPO / f"crates/{name}/pyproject.toml").read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual(pyproject["project"]["version"], expected_version, name)

        migration = (
            workflow.REPO / "docs/hld/11-migration-plan.md"
        ).read_text(encoding="utf-8")
        self.assertNotIn("then stop publishing", migration)
        self.assertIn(
            "Both deprecated shims continue to publish with each coherent "
            "stable train",
            migration,
        )

        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(encoding="utf-8")
        self.assertEqual(
            ci.count(
                'verify_package "$package_root/rdocx-wasm" '
                f'"@tensorbee/rdocx-wasm" "{expected_version}" "rdocx_wasm"'
            ),
            1,
        )
        wasm_source = (workflow.REPO / "crates/rdocx-wasm/src/lib.rs").read_text(
            encoding="utf-8"
        )
        for dependency in ("rdocx", "rdocx-layout"):
            self.assertEqual(
                wasm_source.count(
                    f'{dependency} = {{ path = \\"crates/{dependency}\\", '
                    f'version = \\"{expected_version}\\", '
                    'default-features = false }'
                ),
                1,
                dependency,
            )

        readme_requirements = {
            "README.md": ('rdocx = "0.13.0"', 'version = "0.13.0"'),
            "crates/rdocx-cli/README.md": ("--version '^0.13.0'",),
            "crates/rdocx-html/README.md": ('rdocx-html = "0.13.0"',),
            "crates/rdocx-layout/README.md": ('rdocx-layout = "0.13.0"',),
            "crates/rdocx-opc/README.md": ('rdocx-opc = "0.13.0"',),
            "crates/rdocx-oxml/README.md": ('rdocx-oxml = "0.13.0"',),
            "crates/rdocx-pdf/README.md": ('rdocx-pdf = "0.13.0"',),
        }
        for path, requirements in readme_requirements.items():
            text = (workflow.REPO / path).read_text(encoding="utf-8")
            for requirement in requirements:
                self.assertIn(requirement, text, path)

        readme_gate = (workflow.REPO / "scripts/readme_doctests.py").read_text(
            encoding="utf-8"
        )
        for path, requirements in readme_requirements.items():
            for requirement in requirements:
                self.assertEqual(readme_gate.count(requirement), 1, path)

        for name in incubating_members:
            manifest = tomllib.loads(
                (workflow.REPO / f"crates/{name}/Cargo.toml").read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual(manifest["package"]["version"], "0.11.0", name)
            self.assertIs(
                manifest["package"].get("publish", True),
                name != "rpptx-wasm",
                name,
            )

        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        self.assert_publish_workflow_contract(publish)
        self.assertEqual(
            publish.count(
                "scripts.test_sprint_workflow.SprintWorkflowTests."
                "test_stable_release_family_is_prepared_at_0_13_0"
            ),
            1,
        )

        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = self.assert_v0_13_0_release_notes_truth_contract(changelog)
        self.assertIn("shared OOXML 0.10.0 family", " ".join(notes.split()))

    def test_release_notes_v0_13_0_cover_reviewed_word_outcomes(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        self.assert_v0_13_0_release_notes_truth_contract(changelog)

    @unittest.skipUnless(
        os.environ.get("RDOCX_VERIFY_PUBLISHED_SHARED") == "1",
        "requires the separately published shared 0.10.0 family",
    )
    def test_prepared_rdocx_layout_0_13_0_requires_published_oxml_layout_0_10_0(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory(prefix="rdocx-layout-package-") as temp:
            target = Path(temp) / "target"
            packaged = subprocess.run(
                (
                    "cargo",
                    "package",
                    "--locked",
                    "--allow-dirty",
                    "--config",
                    'patch.crates-io.rdocx-oxml.path="crates/rdocx-oxml"',
                    "--package",
                    "rdocx-layout",
                    "--target-dir",
                    str(target),
                ),
                cwd=workflow.REPO,
                check=False,
                capture_output=True,
                text=True,
            )
            self.assertEqual(
                packaged.returncode,
                0,
                packaged.stdout + packaged.stderr,
            )
            archive = target / "package" / "rdocx-layout-0.13.0.crate"
            self.assertTrue(archive.is_file(), archive)
            with tarfile.open(archive, mode="r:gz") as package:
                normalized = package.extractfile(
                    "rdocx-layout-0.13.0/Cargo.toml"
                )
                self.assertIsNotNone(normalized)
                manifest = tomllib.loads(normalized.read().decode("utf-8"))
            dependency = manifest["dependencies"]["oxml-layout"]
            self.assertEqual(dependency["version"], "0.10.0")
            self.assertNotIn("path", dependency)

            verified_manifest = target / "package" / "rdocx-layout-0.13.0" / "Cargo.toml"
            self.assertTrue(verified_manifest.is_file(), verified_manifest)
            resolved = subprocess.run(
                (
                    "cargo",
                    "tree",
                    "--manifest-path",
                    str(verified_manifest),
                    "--config",
                    f'patch.crates-io.rdocx-oxml.path="{workflow.REPO / "crates/rdocx-oxml"}"',
                    "--edges",
                    "normal",
                    "--prefix",
                    "none",
                ),
                cwd=verified_manifest.parent,
                check=False,
                capture_output=True,
                text=True,
            )
        self.assertEqual(resolved.returncode, 0, resolved.stdout + resolved.stderr)
        self.assertIn("rdocx-layout v0.13.0", resolved.stdout + resolved.stderr)
        self.assertIn("oxml-layout v0.10.0", resolved.stdout + resolved.stderr)

    def test_immutable_rdocx_layout_0_10_1_registry_graph_remains_at_oxml_layout_0_6_0(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory(prefix="rdocx-layout-registry-") as temp:
            consumer = Path(temp) / "published-graph"
            (consumer / "src").mkdir(parents=True)
            (consumer / "Cargo.toml").write_text(
                """[package]
name = "published-rdocx-layout-graph"
version = "0.0.0"
edition = "2024"

[dependencies]
rdocx-layout = "=0.10.1"
""",
                encoding="utf-8",
            )
            (consumer / "src/lib.rs").write_text("", encoding="utf-8")
            env = os.environ.copy()
            env["CARGO_HOME"] = str(Path(temp) / "cargo-home")
            completed = subprocess.run(
                (
                    "cargo",
                    "tree",
                    "--manifest-path",
                    str(consumer / "Cargo.toml"),
                    "--edges",
                    "normal",
                    "--prefix",
                    "none",
                    "--package",
                    "rdocx-layout@0.10.1",
                ),
                cwd=consumer,
                env=env,
                check=False,
                capture_output=True,
                text=True,
            )
        self.assertEqual(
            completed.returncode,
            0,
            completed.stdout + completed.stderr,
        )
        self.assertIn(
            "rdocx-layout v0.10.1",
            completed.stdout + completed.stderr,
        )
        self.assertIn(
            "oxml-layout v0.6.0",
            completed.stdout + completed.stderr,
        )
        self.assertNotIn(
            "oxml-layout v0.7.0",
            completed.stdout + completed.stderr,
        )

    def test_release_notes_v0_10_1_reconcile_release_inventory(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = self.assert_v0_10_1_release_notes_truth_contract(changelog)
        normalized_notes = " ".join(notes.split())
        for addition in (
            "exact selected pages",
            "opaque or transparent PNG",
            "quality-controlled JPEG",
            "deterministic multi-page TIFF",
            "native Word, Python",
            "Word and PowerPoint CLI paths",
        ):
            self.assertIn(addition, normalized_notes, addition)

    def test_release_notes_v0_11_1_reconcile_release_inventory(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = self.assert_v0_11_1_release_notes_truth_contract(changelog)
        normalized = " ".join(notes.split())
        for claim in (
            "conditional hyphenation",
            "Arabic, Devanagari, Thai, and Simplified Chinese",
            "paragraph and run direction",
            "bounded restart pagination",
            "published shared 0.8.0 family",
            "Python, WASM, npm, and PyPI publication authority is unchanged",
        ):
            self.assertIn(claim, normalized, claim)

    def test_release_notes_v0_11_1_reject_partial_or_notification_drift(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = self.assert_v0_11_1_release_notes_truth_contract(changelog)
        mutations = (
            notes.replace(
                "`rdocx-oxml@0.11.0`.",
                "`rdocx-oxml@0.11.0` and `rdocx-layout@0.11.0`.",
                1,
            ),
            notes.replace(
                "Both issues remain open",
                "Both issues are closed",
                1,
            ),
            notes.replace(
                "All four pull requests remain open",
                "All four pull requests are closed",
                1,
            ),
        )
        for mutated_notes in mutations:
            self.assertNotEqual(mutated_notes, notes)
            mutated = changelog.replace(notes, mutated_notes, 1)
            with self.assertRaises(AssertionError):
                self.assert_v0_11_1_release_notes_truth_contract(mutated)

    def test_release_notes_v0_10_1_reject_reversed_landing_truth(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        stable_notes = workflow.render_release_notes(changelog, "v0.10.1")
        mutated_notes = stable_notes.replace(
            "No named external patch landed directly.",
            "A named external patch landed directly.",
            1,
        )
        self.assertNotEqual(mutated_notes, stable_notes)
        mutated = changelog.replace(stable_notes, mutated_notes, 1)
        with self.assertRaises(AssertionError):
            self.assert_v0_10_1_release_notes_truth_contract(mutated)

    def test_release_notes_v0_10_1_reject_expanded_partial_inventory(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        stable_notes = workflow.render_release_notes(changelog, "v0.10.1")
        mutated_notes = stable_notes.replace(
            "contains only `rdocx-opc` and `rdocx-oxml`.",
            "contains only `rdocx-opc`, `rdocx-oxml`, and `rdocx-layout`.",
            1,
        )
        self.assertNotEqual(mutated_notes, stable_notes)
        mutated = changelog.replace(stable_notes, mutated_notes, 1)
        with self.assertRaises(AssertionError):
            self.assert_v0_10_1_release_notes_truth_contract(mutated)

    def test_release_notes_v0_10_1_reject_swapped_contributor_credits(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        stable_notes = workflow.render_release_notes(changelog, "v0.10.1")
        mutated_notes = stable_notes.replace(
            "@emptinessform", "@credit-swap", 1
        ).replace(
            "@pedroassumpcao", "@emptinessform", 1
        ).replace(
            "@credit-swap", "@pedroassumpcao", 1
        )
        self.assertNotEqual(mutated_notes, stable_notes)
        mutated = changelog.replace(stable_notes, mutated_notes, 1)
        with self.assertRaises(AssertionError):
            self.assert_v0_10_1_release_notes_truth_contract(mutated)

    def test_release_notes_v0_9_0_include_marked_content_migration(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = workflow.render_release_notes(changelog, "v0.9.0")
        compatibility = notes.split("### Compatibility\n\n", 1)[1].split(
            "\n### Contributors", 1
        )[0]
        for claim in (
            "`PositionedElement` remains non-exhaustive",
            "`PositionedElement::MarkedContent`",
            "`MarkedContent::children`",
            "`oxml_layout::walk`",
            "`PageFrame::elements`",
        ):
            self.assertIn(claim, compatibility, claim)

    def test_oxml_layout_readme_names_recursive_page_traversal(self) -> None:
        readme = (workflow.REPO / "crates/oxml-layout/README.md").read_text(
            encoding="utf-8"
        )
        for claim in (
            "`PageFrame::elements`",
            "`PositionedElement::MarkedContent`",
            "`MarkedContent::children`",
            "`oxml_layout::walk`",
        ):
            self.assertIn(claim, readme, claim)

    def test_readme_archive_gate_rejects_a_missing_local_patch(self) -> None:
        metadata = readme_doctests.cargo_metadata()
        self.assertIsNotNone(metadata)
        packages = metadata["packages"]
        self.assertTrue(readme_doctests.validate_local_patches(packages))

        without_oxml_core = tuple(
            patch_entry
            for patch_entry in readme_doctests.LOCAL_PATCHES
            if patch_entry[0] != "oxml-core"
        )
        errors = io.StringIO()
        with (
            patch.object(readme_doctests, "LOCAL_PATCHES", without_oxml_core),
            contextlib.redirect_stderr(errors),
        ):
            self.assertFalse(readme_doctests.validate_inventory())
        self.assertIn("('oxml-core', 'crates/oxml-core')", errors.getvalue())

    def test_stable_release_family_has_lockstep_preparation_metadata(self) -> None:
        stable_packages = (
            "rdocx-opc",
            "rdocx-oxml",
            "rdocx-layout",
            "rdocx-html",
            "rdocx-pdf",
            "rdocx",
            "rdocx-cli",
            "rdocx-wasm",
        )

        for name in stable_packages:
            binary = os.environ.get("CARGO_RELEASE_BIN")
            command = [binary] if binary else ["cargo"]
            command.extend(
                (
                    "release",
                    "config",
                    "--manifest-path",
                    str(workflow.REPO / f"crates/{name}/Cargo.toml"),
                )
            )
            result = subprocess.run(
                command,
                check=True,
                capture_output=True,
                text=True,
            )
            release = tomllib.loads(result.stdout)
            self.assertEqual(release["shared-version"], "workspace")
            self.assertEqual(release["tag-name"], "v{{version}}")

    def test_incubating_release_family_has_lockstep_preparation_metadata(self) -> None:
        incubating_packages = (
            "oxml-core",
            "oxml-drawing",
            "oxml-layout",
            "oxml-media",
            "oxml-opc",
            "oxml-pdf",
            "oxml-sml",
            "oxml-cli-support",
            "oxml-chart",
            "rpptx-oxml",
            "rpptx-layout",
            "rpptx-render",
            "rpptx-chart",
            "rpptx",
            "rpptx-cli",
        )

        for name in incubating_packages:
            manifest = tomllib.loads(
                (workflow.REPO / f"crates/{name}/Cargo.toml").read_text(
                    encoding="utf-8"
                )
            )
            release = manifest["package"]["metadata"]["release"]
            self.assertEqual(release["shared-version"], "incubating")
            self.assertEqual(release["tag-name"], "rpptx-v{{version}}")

    def test_incubating_release_family_is_prepared_at_0_11_0(self) -> None:
        incubating_packages = (
            "oxml-core",
            "oxml-opc",
            "oxml-media",
            "oxml-layout",
            "oxml-drawing",
            "oxml-pdf",
            "oxml-sml",
            "oxml-cli-support",
            "oxml-chart",
            "rpptx-oxml",
            "rpptx-chart",
            "rpptx-layout",
            "rpptx-render",
            "rpptx",
            "rpptx-cli",
        )
        preparation_packages = (*incubating_packages, "rpptx-wasm")
        expected_version = "0.11.0"
        root = tomllib.loads((workflow.REPO / "Cargo.toml").read_text(encoding="utf-8"))
        self.assertEqual(root["workspace"]["package"]["version"], "0.13.0")
        dependencies = root["workspace"]["dependencies"]
        lock = tomllib.loads((workflow.REPO / "Cargo.lock").read_text(encoding="utf-8"))
        lock_versions = {
            package["name"]: package["version"]
            for package in lock["package"]
            if package["name"] in preparation_packages
        }

        self.assertEqual(set(lock_versions), set(preparation_packages))
        for name in incubating_packages:
            manifest = tomllib.loads(
                (workflow.REPO / f"crates/{name}/Cargo.toml").read_text(
                    encoding="utf-8"
                )
            )
            self.assertEqual(manifest["package"]["version"], expected_version, name)
            self.assertTrue(manifest["package"].get("description", "").strip(), name)
            self.assertTrue(manifest["package"].get("publish", True), name)
            self.assertEqual(dependencies[name]["version"], expected_version, name)
            self.assertEqual(lock_versions[name], expected_version, name)

        wasm = tomllib.loads(
            (workflow.REPO / "crates/rpptx-wasm/Cargo.toml").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(wasm["package"]["version"], expected_version)
        self.assertFalse(wasm["package"]["publish"])
        self.assertNotIn("rpptx-wasm", dependencies)
        self.assertEqual(lock_versions["rpptx-wasm"], expected_version)

        readme_requirements = {
            "crates/oxml-core/README.md": ('oxml-core = "0.11.0"',),
            "crates/oxml-drawing/README.md": ('oxml-drawing = "0.11.0"',),
            "crates/oxml-layout/README.md": ('version = "0.11.0"',),
            "crates/oxml-media/README.md": ('oxml-media = "0.11.0"',),
            "crates/oxml-opc/README.md": ('oxml-opc = "0.11.0"',),
            "crates/oxml-pdf/README.md": (
                'oxml-pdf = "0.11.0"',
                'oxml-layout = "0.11.0"',
            ),
            "crates/oxml-chart/README.md": ('oxml-chart = "0.11.0"',),
            "crates/rpptx-chart/README.md": ('rpptx-chart = "0.11.0"',),
            "crates/rpptx-cli/README.md": ("--version '^0.11.0'",),
            "crates/rpptx-layout/README.md": ('rpptx-layout = "0.11.0"',),
            "crates/rpptx-oxml/README.md": ('rpptx-oxml = "0.11.0"',),
            "crates/rpptx-render/README.md": ('rpptx-render = "0.11.0"',),
        }
        for path, requirements in readme_requirements.items():
            text = (workflow.REPO / path).read_text(encoding="utf-8")
            for requirement in requirements:
                self.assertIn(requirement, text, path)

        source_requirements = {
            "crates/oxml-chart/src/lib.rs": (
                'manifest.contains("version = \\"0.11.0\\"")',
            ),
            "crates/oxml-drawing/src/lib.rs": (
                'manifest.contains("version = \\"0.11.0\\"")',
            ),
            "crates/rdocx-wasm/src/lib.rs": (
                'oxml-layout = { path = \\"crates/oxml-layout\\", '
                'version = \\"0.11.0\\", default-features = false }',
            ),
            "crates/rpptx-oxml/tests/integration.rs": (
                'manifest.contains("version = \\"0.11.0\\"")',
            ),
            "crates/rpptx-render/src/lib.rs": (
                'manifest.contains("version = \\"0.11.0\\"")',
            ),
            "crates/rpptx-wasm/src/lib.rs": (
                'rpptx = { path = \\"crates/rpptx\\", version = \\"0.11.0\\", '
                'default-features = false }',
            ),
            "crates/rpptx/tests/integration.rs": (
                'rpptx = { path = \\"crates/rpptx\\", version = \\"0.11.0\\", '
                'default-features = false }',
                'manifest.contains("version = \\"0.11.0\\"")',
            ),
        }
        for path, requirements in source_requirements.items():
            text = (workflow.REPO / path).read_text(encoding="utf-8")
            for requirement in requirements:
                self.assertIn(requirement, text, path)

        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assertEqual(
            ci.count(
                'verify_package "$package_root/rpptx-wasm" '
                '"@tensorbee/rpptx-wasm" "0.11.0" "rpptx_wasm"'
            ),
            1,
        )

        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        self.assert_publish_workflow_contract(publish)
        self.assertEqual(
            publish.count(
                "scripts.test_sprint_workflow.SprintWorkflowTests."
                "test_incubating_release_family_is_prepared_at_0_11_0"
            ),
            1,
        )

    def test_release_notes_rpptx_v0_9_0_cover_presentation_depth_boundary(
        self,
    ) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = workflow.render_release_notes(changelog, "rpptx-v0.9.0")
        normalized = " ".join(notes.split())
        for claim in (
            "collaboration",
            "timing",
            "media",
            "SmartArt",
            "ODP",
            "HTML",
            "PDF",
            "0.9.0",
        ):
            self.assertIn(claim, normalized, claim)
        self.assertIn("stable Word family remains at 0.11.1", normalized)
        self.assertNotIn("stable Word family moves", normalized)
        self.assertNotIn("github.com/tensorbee/rdocx/issues/", notes)
        self.assertNotIn("github.com/tensorbee/rdocx/pull/", notes)

    def test_release_notes_rpptx_v0_10_0_cover_selected_shared_changes(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = workflow.render_release_notes(changelog, "rpptx-v0.10.0")
        normalized = " ".join(notes.split())
        for claim in (
            "strict XML 1.0 lexical validator",
            "baseline-aware inline groups",
            "Word glossary",
            "0.10.0",
            "stable Word family remains at 0.12.0",
            "No external issue or pull request belongs to the selected family",
        ):
            self.assertIn(claim, normalized, claim)
        self.assertIn("Atul Sharma maintained the release", normalized)
        self.assertNotIn("github.com/tensorbee/rdocx/issues/", notes)
        self.assertNotIn("github.com/tensorbee/rdocx/pull/", notes)

    def test_release_notes_rpptx_v0_11_0_cover_word_package_constants(self) -> None:
        changelog = (workflow.REPO / "CHANGELOG.md").read_text(encoding="utf-8")
        notes = workflow.render_release_notes(changelog, "rpptx-v0.11.0")
        normalized = " ".join(notes.split())
        for claim in (
            "Word main content-type constants",
            "DOCX, DOCM, DOTX, and DOTM",
            "0.11.0",
            "stable Word family remains at 0.13.0",
            "No external issue or pull request belongs to the selected family",
        ):
            self.assertIn(claim, normalized, claim)
        self.assertIn("Atul Sharma maintained the release", normalized)
        self.assertNotIn("github.com/tensorbee/rdocx/issues/", notes)
        self.assertNotIn("github.com/tensorbee/rdocx/pull/", notes)

    def assert_release_preparation_metadata_contract(
        self, manifest_overrides: dict[str, str] | None = None
    ) -> None:
        manifest_overrides = manifest_overrides or {}
        root = tomllib.loads((workflow.REPO / "Cargo.toml").read_text(encoding="utf-8"))
        release = root["workspace"]["metadata"]["release"]

        self.assertTrue(release["consolidate-commits"])
        self.assertEqual(release["dependent-version"], "upgrade")
        self.assertTrue(release["verify"])
        self.assertFalse(release["publish"])
        self.assertFalse(release["tag"])
        self.assertFalse(release["push"])
        self.assertNotIn("pre-release-replacements", release)

        family_members = {"workspace": [], "incubating": []}
        manifests = {}
        for member in root["workspace"]["members"]:
            manifest_text = manifest_overrides.get(
                member,
                (workflow.REPO / member / "Cargo.toml").read_text(encoding="utf-8"),
            )
            manifest = tomllib.loads(manifest_text)
            manifests[member] = manifest
            family = manifest["package"]["metadata"]["release"]["shared-version"]
            self.assertIn(family, family_members)
            family_members[family].append(member)

        self.assertEqual(
            tuple(family_members["workspace"]),
            (
                "crates/oxml-py-support",
                "crates/rpptx-py",
                "crates/rdocx-opc",
                "crates/rdocx-oxml",
                "crates/rdocx",
                "crates/rdocx-layout",
                "crates/rdocx-pdf",
                "crates/rdocx-html",
                "crates/rdocx-py",
                "crates/rdocx-cli",
                "crates/rdocx-wasm",
            ),
        )
        self.assertEqual(
            tuple(family_members["incubating"]),
            (
                "crates/oxml-core",
                "crates/oxml-drawing",
                "crates/oxml-layout",
                "crates/oxml-media",
                "crates/oxml-opc",
                "crates/oxml-pdf",
                "crates/oxml-sml",
            "crates/oxml-cli-support",
            "crates/oxml-chart",
            "crates/rpptx",
                "crates/rpptx-cli",
                "crates/rpptx-chart",
                "crates/rpptx-layout",
                "crates/rpptx-oxml",
                "crates/rpptx-render",
                "crates/rpptx-wasm",
            ),
        )

        family_counts = {
            family: len(members) for family, members in family_members.items()
        }
        self.assertEqual(family_counts, {"workspace": 11, "incubating": 16})

        wasm_package = manifests["crates/rpptx-wasm"]["package"]
        self.assertEqual(wasm_package["name"], "rpptx-wasm")
        self.assertEqual(wasm_package["version"], "0.11.0")
        self.assertTrue(wasm_package.get("description", "").strip())
        self.assertFalse(wasm_package["publish"])
        self.assertEqual(
            wasm_package["metadata"]["release"],
            {
                "shared-version": "incubating",
                "tag-name": "rpptx-v{{version}}",
            },
        )

        dependencies = root["workspace"]["dependencies"]
        self.assertNotIn("rpptx-wasm", dependencies)
        lock = tomllib.loads((workflow.REPO / "Cargo.lock").read_text(encoding="utf-8"))
        wasm_lock_versions = tuple(
            package["version"]
            for package in lock["package"]
            if package["name"] == "rpptx-wasm"
        )
        self.assertEqual(wasm_lock_versions, ("0.11.0",))

    def test_release_preparation_metadata_cannot_mutate_external_state(self) -> None:
        self.assert_release_preparation_metadata_contract()

    def test_release_preparation_metadata_rejects_a_wasm_family_mutation(
        self,
    ) -> None:
        member = "crates/rpptx-wasm"
        manifest = (workflow.REPO / member / "Cargo.toml").read_text(
            encoding="utf-8"
        )
        mutated = manifest.replace(
            'shared-version = "incubating"',
            'shared-version = "workspace"',
            1,
        )
        self.assertNotEqual(mutated, manifest)
        with self.assertRaises(AssertionError):
            self.assert_release_preparation_metadata_contract({member: mutated})

    def test_release_preparation_metadata_rejects_wasm_tag_and_version_mutations(
        self,
    ) -> None:
        member = "crates/rpptx-wasm"
        manifest = (workflow.REPO / member / "Cargo.toml").read_text(
            encoding="utf-8"
        )
        mutations = {
            "stable-tag-template": manifest.replace(
                'tag-name = "rpptx-v{{version}}"',
                'tag-name = "v{{version}}"',
                1,
            ),
            "workspace-version": manifest.replace(
                'version = "0.11.0"',
                "version.workspace = true",
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, manifest, name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_release_preparation_metadata_contract({member: mutated})

    def assert_partial_v0_11_0_cleanup_contract(self, plan: str) -> None:
        approach = plan[
            plan.index("## Approach") : plan.index("## Rejected alternatives")
        ]
        normalized = " ".join(approach.split())
        bash_blocks = re.findall(r"```bash\n(.*?)```", plan, flags=re.DOTALL)
        self.assertEqual(len(bash_blocks), 1)
        fenced_commands = tuple(
            line.strip() for line in bash_blocks[0].splitlines() if line.strip()
        )
        authorized_commands = (
            "cargo yank --registry crates-io --version 0.11.0 rdocx-opc",
            "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml",
        )
        self.assertEqual(fenced_commands, authorized_commands)

        authorized_block = (
            "```bash\n" + "\n".join(authorized_commands) + "\n```"
        )
        self.assertEqual(plan.count(authorized_block), 1)
        plan_without_authorized_commands = plan.replace(authorized_block, "", 1)
        command_token = re.compile(
            r"\b(?:env|command|sudo|cargo|git|gh|curl|wget|python|python3|npm|"
            r"npx|bash|sh|zsh)\s+"
        )
        self.assertNotRegex(plan_without_authorized_commands, command_token)
        for contract in (
            "Complete coherent stable releases remain available.",
            "separate final approval immediately before the first yank",
            "No other external mutation is authorized.",
            "Do not delete or move tags, create a release, post comments, close external issues or pull requests, or alter any other version.",
        ):
            self.assertIn(contract, normalized)

    def test_partial_v0_11_0_cleanup_contract(self) -> None:
        plan = (workflow.REPO / ".claude/plans/F-X070-design.md").read_text(
            encoding="utf-8"
        )
        self.assert_partial_v0_11_0_cleanup_contract(plan)
        self.assertEqual(plan.count("`docs/hld/11-migration-plan.md`"), 2)

        migration = (
            workflow.REPO / "docs/hld/11-migration-plan.md"
        ).read_text(encoding="utf-8")
        normalized_migration = " ".join(migration.split())
        self.assertIn("Do not yank a complete coherent release.", migration)
        self.assertIn(
            "Both incomplete 0.11.0 entries are yanked after the complete "
            "0.11.1 family verified and a separate immediate approval was granted.",
            normalized_migration,
        )
        self.assertIn(
            "The complete seven-package rdocx family is published at stable "
            "0.11.1",
            normalized_migration,
        )
        self.assertIn(
            "The complete 15-package shared and PowerPoint family is published "
            "at 0.8.0",
            normalized_migration,
        )
        self.assertIn(
            "The immutable 0.6.0 stable and 0.1.3 shared family releases remain "
            "available as historical boundaries.",
            normalized_migration,
        )
        self.assertIn(
            "Normal local sprint ledgers, progress notes, review artifacts, and "
            "handoff records still advance through the feature workflow.",
            normalized_migration,
        )
        self.assertIn(
            "Both incomplete 0.11.0 entries are yanked after the complete "
            "0.11.1 family verified and a separate immediate approval was granted.",
            normalized_migration,
        )

        backlog = (
            workflow.REPO / "docs/hld/14-development-backlog.md"
        ).read_text(encoding="utf-8")
        normalized_backlog = " ".join(backlog.split())
        local_record_contract = (
            "Normal local sprint ledgers, progress notes, review artifacts, and "
            "handoff records still advance through the feature workflow."
        )
        self.assertIn(local_record_contract, normalized_backlog)
        self.assertIn(local_record_contract, " ".join(plan.split()))

        normalized_plan = " ".join(plan.split())
        for evidence in (
            "`rdocx-opc@0.11.0` and `rdocx-oxml@0.11.0` read back with "
            "`yanked=true`",
            "all seven 0.11.1 packages read back with `yanked=false` under sole "
            "owner `mantissaman (Atul Sharma)`",
            "the other five 0.11.0 package endpoints return 404",
            "The remote annotated `v0.11.0` tag still peels to "
            "`25350d000ed7ed96bf4f6e371f01f8fbc8e2cec4`",
            "the v0.11.0 GitHub release lookup returns 404",
        ):
            self.assertIn(evidence, normalized_plan)
        for completed_item in (
            "- [x] Stop for separate final approval immediately before the first yank.",
            "- [x] Yank exactly `rdocx-opc@0.11.0` and `rdocx-oxml@0.11.0`.",
            "- [x] Verify yanked flags, complete live 0.11.1 family, immutable tag, and absent v0.11.0 release.",
        ):
            self.assertIn(completed_item, plan)
        self.assertIn(
            "- [x] Complete the delivery records without any unrelated external mutation.",
            plan,
        )

        mutations = {
            "missing-package": plan.replace(
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n",
                "",
                1,
            ),
            "other-version": plan.replace(
                "--version 0.11.0 rdocx-opc",
                "--version 0.11.1 rdocx-opc",
                1,
            ),
            "extra-package": plan.replace(
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n",
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n"
                "cargo yank --registry crates-io --version 0.11.0 rdocx-layout\n",
                1,
            ),
            "tag-deletion": plan.replace(
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n",
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n"
                "git push origin --delete refs/tags/v0.11.0\n",
                1,
            ),
            "github-release-creation": plan.replace(
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n",
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n"
                "gh release create v0.11.0\n",
                1,
            ),
            "issue-closure": plan.replace(
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n",
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n"
                "gh issue close 53\n",
                1,
            ),
            "pr-closure": plan.replace(
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n",
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n"
                "gh pr close 54\n",
                1,
            ),
            "unrelated-registry-mutation": plan.replace(
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n",
                "cargo yank --registry crates-io --version 0.11.0 rdocx-oxml\n"
                "cargo publish --registry crates-io -p rdocx\n",
                1,
            ),
            "out-of-block-mutation": plan.replace(
                "```\n\nNo other external mutation is authorized.",
                "```\ncargo publish --registry crates-io -p rdocx\n\n"
                "No other external mutation is authorized.",
                1,
            ),
            "outside-approach-release": plan.replace(
                "## Rejected alternatives\n",
                "## Rejected alternatives\n\ngh release create v0.11.0\n",
                1,
            ),
            "env-prefixed-release": plan.replace(
                "```\n\nNo other external mutation is authorized.",
                "```\nenv gh release create v0.11.0\n\n"
                "No other external mutation is authorized.",
                1,
            ),
            "env-prefixed-tag-deletion": plan.replace(
                "```\n\nNo other external mutation is authorized.",
                "```\nenv git push origin --delete refs/tags/v0.11.0\n\n"
                "No other external mutation is authorized.",
                1,
            ),
            "env-prefixed-registry-mutation": plan.replace(
                "```\n\nNo other external mutation is authorized.",
                "```\nenv cargo publish --registry crates-io -p rdocx\n\n"
                "No other external mutation is authorized.",
                1,
            ),
            "shell-wrapped-pr-closure": plan.replace(
                "```\n\nNo other external mutation is authorized.",
                "```\ncommand sh -c 'gh pr close 54'\n\n"
                "No other external mutation is authorized.",
                1,
            ),
            "missing-immediate-approval": plan.replace(
                "separate final approval immediately before the first yank",
                "general sprint approval",
                1,
            ),
            "release-mutation": plan.replace(
                "No other external mutation is authorized.",
                "Create a release after the cleanup.",
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, plan, name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_partial_v0_11_0_cleanup_contract(mutated)

    def test_release_command_is_the_only_release_tag_authority(self) -> None:
        release = (workflow.REPO / ".claude/commands/release.md").read_text(
            encoding="utf-8"
        )
        run_sprint = (workflow.REPO / ".claude/commands/run-sprint.md").read_text(
            encoding="utf-8"
        )
        close_sprint = (workflow.REPO / ".claude/commands/close-sprint.md").read_text(
            encoding="utf-8"
        )
        complete_feature = (
            workflow.REPO / ".claude/commands/complete-feature.md"
        ).read_text(encoding="utf-8")
        normalized_release = " ".join(release.split())

        self.assertIn("only command", release)
        self.assertIn("# /release {vX.Y.Z | rpptx-vX.Y.Z}", release)
        self.assertIn(
            "The exact seven-package stable set is `rdocx-opc`, `rdocx-oxml`, "
            "`rdocx-layout`, `rdocx-html`, `rdocx-pdf`, `rdocx`, and "
            "`rdocx-cli`,",
            normalized_release,
        )
        self.assertIn(
            "The exact 15-package incubating set is `oxml-core`, `oxml-opc`, "
            "`oxml-media`, `oxml-layout`, `oxml-drawing`, `oxml-pdf`, "
            "`oxml-sml`, `oxml-cli-support`, `oxml-chart`, `rpptx-oxml`, `rpptx-chart`, `rpptx-layout`, "
            "`rpptx-render`, `rpptx`, and `rpptx-cli`.",
            normalized_release,
        )
        self.assertIn("go or no-go immediately", normalized_release)
        self.assertIn("The 22 patches keep packaged", normalized_release)
        self.assertIn(
            "The `oxml-layout` archive must contain its complete bundled TTF "
            "and legal-file inventory. The `rdocx-layout` archive must not "
            "duplicate those assets.",
            normalized_release,
        )
        self.assertNotIn(
            "The `rdocx-layout` and `oxml-layout` archives must contain",
            normalized_release,
        )
        self.assertIn(
            "Create one annotated tag for the requested argument",
            normalized_release,
        )
        self.assertIn("Push only that requested tag", normalized_release)
        self.assertIn("/release", run_sprint)
        self.assertIn("Leave it\n`reviewed`", run_sprint)
        self.assertIn("/release", close_sprint)
        self.assertIn("deferred to /release", complete_feature)

        mutated = release.replace(
            "The 22 patches keep packaged", "The 21 patches keep packaged", 1
        )
        self.assertNotEqual(mutated, release)
        with self.assertRaises(AssertionError):
            self.assertIn("The 22 patches keep packaged", " ".join(mutated.split()))

    def assert_run_sprint_dependency_release_checkpoint_contract(
        self, run_sprint: str
    ) -> None:
        normalized = " ".join(run_sprint.split())
        checkpoint = run_sprint[
            run_sprint.index("#### Release dependency extension") :
            run_sprint.index("## 5. Integrate")
        ]

        self.assertIn(
            "a release F-ID is a dependency of any unfinished story in the same sprint",
            normalized,
        )
        ordered_steps = (
            "Prepare and integrate the release F-ID",
            "Run `/verify --full` and `/sprint-review`",
            "Follow `/release <tag>`",
            "Finalise the release F-ID's delivery records",
            "Return the phase to `implementation` again",
        )
        positions = tuple(checkpoint.index(step) for step in ordered_steps)
        self.assertEqual(positions, tuple(sorted(positions)))
        self.assertIn("separate final approval", checkpoint)
        self.assertGreaterEqual(checkpoint.count("current HEAD"), 2)
        self.assertIn("Never use checkpoint evidence for final closure", checkpoint)
        self.assertIn("close-preflight SNN", run_sprint)

    def assert_run_sprint_dependency_prefix_checkpoint_contract(
        self, run_sprint: str
    ) -> None:
        normalized = " ".join(run_sprint.split())
        checkpoint = run_sprint[
            run_sprint.index("### Dependency-prefix checkpoint") :
            run_sprint.index("#### Release dependency extension")
        ]

        self.assertIn(
            "a formal dependency is integrated and `reviewed` but not `completed`",
            normalized,
        )
        ordered_steps = (
            "Integrate the prepared dependency prefix",
            "Run `/verify --full`",
            "Finalise the reviewed non-release prefix",
            "Commit the clean review file",
            "record the clean review at the resulting HEAD",
            "Rerun `/verify --full` because the review commit changed HEAD",
            "Return the phase to `implementation`",
        )
        positions = tuple(checkpoint.index(step) for step in ordered_steps)
        self.assertEqual(positions, tuple(sorted(positions)))
        self.assertIn("Do not run a confirmation review", checkpoint)
        self.assertIn("Do not claim the dependent wave", checkpoint)
        self.assertIn("Pass numbering remains global", checkpoint)
        self.assertIn("at most the configured review-pass bound", checkpoint)
        self.assertIn("scheduled dependency-prefix boundary", checkpoint)

    def test_run_sprint_requires_ordinary_dependency_prefix_checkpoints(self) -> None:
        run_sprint = (workflow.REPO / ".claude/commands/run-sprint.md").read_text(
            encoding="utf-8"
        )
        self.assert_run_sprint_dependency_prefix_checkpoint_contract(run_sprint)

        mutations = {
            "missing-trigger": run_sprint.replace(
                "dependency is integrated and `reviewed` but not `completed`",
                "a formal dependency exists",
                1,
            ),
            "missing-review-commit": run_sprint.replace(
                "Commit the clean review file", "Leave the review file uncommitted", 1
            ),
            "confirmation-review": run_sprint.replace(
                "Do not run a confirmation review",
                "Run a confirmation review",
                1,
            ),
            "missing-checkpoint-review-bound": run_sprint.replace(
                "at most the configured review-pass bound",
                "an unlimited number of review passes",
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, run_sprint, name)
            with self.subTest(name=name), self.assertRaises(
                (AssertionError, ValueError)
            ):
                self.assert_run_sprint_dependency_prefix_checkpoint_contract(mutated)

    def test_run_sprint_requires_dependency_release_checkpoints(self) -> None:
        run_sprint = (workflow.REPO / ".claude/commands/run-sprint.md").read_text(
            encoding="utf-8"
        )
        self.assert_run_sprint_dependency_release_checkpoint_contract(run_sprint)

        mutations = {
            "missing-trigger": run_sprint.replace(
                "a release F-ID is a dependency of any unfinished story in the same sprint",
                "a release F-ID is ready",
                1,
            ),
            "missing-approval": run_sprint.replace(
                "separate final approval", "release approval", 1
            ),
            "missing-final-evidence-boundary": run_sprint.replace(
                "Never use checkpoint evidence for final closure",
                "Checkpoint evidence is reusable for final closure",
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, run_sprint, name)
            with self.subTest(name=name), self.assertRaises(
                (AssertionError, ValueError)
            ):
                self.assert_run_sprint_dependency_release_checkpoint_contract(mutated)

    def assert_release_command_notes_contract(self, release: str) -> None:
        preflight = release[release.index("## Preconditions") : release.index("## Final approval")]
        approval = release[release.index("## Final approval") : release.index("## Release")]
        publication = release[release.index("## Release") : release.index("## Finalise the release F-ID")]

        self.assertIn(
            "python3 scripts/sprint_workflow.py release-notes <requested-tag>\n"
            "   --check",
            preflight,
        )
        self.assertIn(
            "python3 scripts/sprint_workflow.py release-notes\n"
            "   <requested-tag> --render",
            preflight,
        )
        self.assertIn("committed `CHANGELOG.md`", preflight)
        self.assertIn("contribution inventory", preflight)
        self.assertIn("notes source as\n`CHANGELOG.md`", approval)
        self.assertIn("Include the rendered notes", approval)
        self.assertIn("included issue and pull-request URLs", approval)
        self.assertIn("comment that will be posted to each record", approval)
        self.assertIn("compare it byte\n   for byte", publication)
        self.assertIn(
            "python3 scripts/sprint_workflow.py release-notes\n"
            "   <requested-tag> --render",
            publication,
        )
        self.assertIn("notify every issue", publication)
        self.assertIn("Record every resulting\n   comment URL", publication)

    def test_release_notes_are_checked_before_approval_and_after_publication(
        self,
    ) -> None:
        release = (workflow.REPO / ".claude/commands/release.md").read_text(
            encoding="utf-8"
        )
        self.assert_release_command_notes_contract(release)

        mutations = {
            "missing-check": release.replace(
                "python3 scripts/sprint_workflow.py release-notes "
                "<requested-tag>\n   --check",
                "release-notes omitted",
                1,
            ),
            "wrong-validator-executable": release.replace(
                "python3 scripts/sprint_workflow.py release-notes "
                "<requested-tag>\n   --check",
                "echo release-notes <requested-tag>\n   --check",
                1,
            ),
            "missing-render-review": release.replace(
                "Include the rendered notes", "Summarise the notes", 1
            ),
            "missing-source": release.replace(
                "notes source as\n`CHANGELOG.md`", "notes source elsewhere", 1
            ),
            "missing-published-body-comparison": release.replace(
                "compare it byte\n   for byte", "inspect it", 1
            ),
            "missing-contribution-inventory": release.replace(
                "contribution inventory", "contributor summary", 1
            ),
            "missing-record-notification": release.replace(
                "notify every issue", "summarize every issue", 1
            ),
            "missing-comment-evidence": release.replace(
                "Record every resulting\n   comment URL",
                "Record a notification summary",
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, release, name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_release_command_notes_contract(mutated)

    def test_release_notes_command_is_a_generated_agent_ceremony(self) -> None:
        command_path = workflow.REPO / ".claude/commands/release-notes.md"
        skill_path = workflow.REPO / ".agents/skills/release-notes/SKILL.md"
        interface_path = (
            workflow.REPO / ".agents/skills/release-notes/agents/openai.yaml"
        )
        command = command_path.read_text(encoding="utf-8")
        skill = skill_path.read_text(encoding="utf-8")
        interface = interface_path.read_text(encoding="utf-8")
        digest = hashlib.sha256(command_path.read_bytes()).hexdigest()
        normalized = " ".join(command.split())

        for evidence in (
            "design plan",
            "AS_BUILT.md",
            "reviewed commits",
            "merged pull requests",
            "contribution inventory",
            "direct Markdown link",
            "notification list",
            "contributor",
            "compatibility",
        ):
            self.assertIn(evidence, normalized)
        self.assertIn("# /release-notes {vX.Y.Z | rpptx-vX.Y.Z}", command)
        self.assertIn("Canonical source: `.claude/commands/release-notes.md`.", skill)
        self.assertIn(f"Source SHA-256: `{digest}`.", skill)
        self.assertIn("allow_implicit_invocation: false", interface)

    def test_release_notes_require_complete_reviewed_changelog_sections(self) -> None:
        complete = """# Changelog

## v0.8.0

### Highlights

Stable layout access is now available.

### Added

Document layout accessors.

### Fixed

Source positions survive layout.

### Compatibility

Review the layout source changes before upgrading.

### Contributors

Pedro Assumpcao and the rdocx maintainers.

## rpptx-v0.4.0

### Highlights

The incubating family is now publishable.

### Added

The complete PowerPoint package family.

### Fixed

Publication metadata for the incubating crates.

### Compatibility

This family remains incubating.

### Contributors

The rdocx maintainers.
"""

        stable = workflow.render_release_notes(complete, "v0.8.0")
        incubating = workflow.render_release_notes(complete, "rpptx-v0.4.0")
        expected_stable = """### Highlights

Stable layout access is now available.

### Added

Document layout accessors.

### Fixed

Source positions survive layout.

### Compatibility

Review the layout source changes before upgrading.

### Contributors

Pedro Assumpcao and the rdocx maintainers.
"""
        self.assertEqual(stable, expected_stable)
        self.assertNotIn("## v0.8.0", stable)
        self.assertNotIn("rpptx-v0.4.0", stable)
        self.assertIn("The incubating family is now publishable.", incubating)

        raw_html_with_markdown = (
            "<div hidden>Ignored raw HTML text.</div>\n\n"
            "Stable layout access is now available."
        )
        valid_html = complete.replace(
            "Stable layout access is now available.",
            raw_html_with_markdown,
            1,
        )
        self.assertEqual(
            workflow.render_release_notes(valid_html, "v0.8.0"),
            expected_stable.replace(
                "Stable layout access is now available.",
                raw_html_with_markdown,
                1,
            ),
        )

        section_text = {
            "Highlights": "Stable layout access is now available.",
            "Added": "Document layout accessors.",
            "Fixed": "Source positions survive layout.",
            "Compatibility": (
                "Review the layout source changes before upgrading."
            ),
            "Contributors": "Pedro Assumpcao and the rdocx maintainers.",
        }
        empty_html_forms = {
            "empty-container": "<div></div>",
            "adjacent-empty-forms": (
                "<span></span><br><release-note></release-note><!-- hidden -->"
                "<?empty?><!EMPTY><![CDATA[hidden cdata text]]>"
                "&nbsp;&#32;&#x20;&ZeroWidthSpace;"
            ),
            "non-visible-containers": (
                "<script>hidden script text</script>"
                "<style>hidden style text</style>"
                "<template>hidden template text</template>"
                "<head>hidden head text</head>"
                "<title>hidden title text</title>"
            ),
            "ordinary-container-content": (
                "<div>Visible-looking raw HTML text</div>"
            ),
            "hidden-attribute": "<div hidden>Invisible release notes</div>",
            "hidden-css": (
                '<section style="display: none">Invisible release notes</section>'
            ),
            "iframe-fallback": "<iframe>Fallback release notes</iframe>",
            "noscript-content": "<noscript>Fallback release notes</noscript>",
            "custom-container": (
                "<release-note>Raw custom-element text</release-note>"
            ),
        }
        for heading, original in section_text.items():
            for form, empty_text in empty_html_forms.items():
                mutated = complete.replace(
                    f"### {heading}\n\n{original}",
                    f"### {heading}\n\n{empty_text}",
                    1,
                )
                self.assertNotEqual(mutated, complete)
                with self.subTest(empty_heading=heading, empty_form=form):
                    with self.assertRaisesRegex(
                        ValueError,
                        rf"section `### {re.escape(heading)}` is empty",
                    ):
                        workflow.render_release_notes(mutated, "v0.8.0")

        empty_markdown_forms = {
            "empty-inline-link": "[](https://example.com/release-notes)",
            "reference-definition": (
                "[release-notes]: https://example.com/release-notes"
            ),
            "multiline-reference-definition": (
                "[release-notes]:\n"
                "  <https://example.com/release-notes>\n"
                '  "hidden reference title"'
            ),
            "empty-fence-with-info": "```python\n\n```",
            "adjacent-empty-forms": (
                "[](https://example.com)![](release.png)[][release-notes]<>\n"
                "[release-notes]: https://example.com/release-notes"
            ),
            "empty-list-item": "1. [](https://example.com/release-notes)",
            "escaped-asterisk": r"\*",
            "escaped-backslash": r"\\",
            "escaped-open-bracket": r"\[",
            "escaped-exclamation": r"\!",
            "escape-guard-collision": "\u2060\\*\u2060",
        }
        for heading, original in section_text.items():
            for form, empty_text in empty_markdown_forms.items():
                mutated = complete.replace(
                    f"### {heading}\n\n{original}",
                    f"### {heading}\n\n{empty_text}",
                    1,
                )
                self.assertNotEqual(mutated, complete)
                with self.subTest(empty_markdown_heading=heading, empty_form=form):
                    with self.assertRaisesRegex(
                        ValueError,
                        rf"section `### {re.escape(heading)}` is empty",
                    ):
                        workflow.render_release_notes(mutated, "v0.8.0")

        invisible_code_payloads = {
            "zero-width-space": "\u200b",
            "word-joiner": "\u2060",
            "directional-marks": "\u200e\u200f",
            "directional-isolates": "\u2066\u2069",
            "unicode-spacing": "\u00a0\u202f",
            "soft-hyphen-and-bom": "\u00ad\ufeff",
            "combining-and-variation": "\u0301\ufe0f",
        }
        for heading, original in section_text.items():
            for payload_name, payload in invisible_code_payloads.items():
                for code_form, empty_text in (
                    ("inline", f"`{payload}`"),
                    ("fenced", f"```text\n{payload}\n```"),
                ):
                    mutated = complete.replace(
                        f"### {heading}\n\n{original}",
                        f"### {heading}\n\n{empty_text}",
                        1,
                    )
                    self.assertNotEqual(mutated, complete)
                    with self.subTest(
                        invisible_code_heading=heading,
                        payload=payload_name,
                        code_form=code_form,
                    ):
                        with self.assertRaisesRegex(
                            ValueError,
                            rf"section `### {re.escape(heading)}` is empty",
                        ):
                            workflow.render_release_notes(mutated, "v0.8.0")

        meaningful_markdown_forms = {
            "link-label": (
                "[Stable layout access](https://example.com/release-notes)"
            ),
            "fenced-code": "```html\n<div></div>\n```",
            "inline-code": "`<release-note></release-note>`",
            "inline-symbol": "`✨`",
            "fenced-symbol": "```text\n✨\n```",
            "symbol-with-formatting": "`\u200b✨\u2060`",
            "autolink": "<https://example.com/release-notes>",
            "image-label": "![Stable layout diagram](release.png)",
            "escaped-element": r"\<w:document\>",
            "escaped-custom-element": (
                r"\<release-note\>Visible release words\</release-note\>"
            ),
            "escaped-comment": r"\<!-- Visible release words --\>",
            "escaped-processing-instruction": r"\<?release-notes?\>",
            "escaped-declaration": r"\<!RELEASE-NOTES\>",
            "escaped-cdata": r"\<![CDATA[Visible release words]]\>",
            "escaped-entity": r"\&nbsp;",
            "escaped-brackets": r"\[release\]",
            "escaped-punctuation-before-letters": r"\*release",
            "escaped-punctuation-after-letters": r"release\!",
            "backslash-before-letter": r"\release",
        }
        for form, meaningful_text in meaningful_markdown_forms.items():
            valid_markdown = complete.replace(
                "Stable layout access is now available.",
                meaningful_text,
                1,
            )
            with self.subTest(meaningful_markdown=form):
                self.assertEqual(
                    workflow.render_release_notes(valid_markdown, "v0.8.0"),
                    expected_stable.replace(
                        "Stable layout access is now available.",
                        meaningful_text,
                        1,
                    ),
                )

        commented_duplicate = complete.replace(
            "### Added\n\nDocument layout accessors.",
            "<!--\n### Added\n\nHidden comment text.\n-->\n\n"
            "### Added\n\nDocument layout accessors.",
            1,
        )
        self.assertIn(
            "Document layout accessors.",
            workflow.render_release_notes(commented_duplicate, "v0.8.0"),
        )
        fenced_duplicate = complete.replace(
            "### Added\n\nDocument layout accessors.",
            "````markdown\n```\n### Added\n\nHidden code text.\n````\n\n"
            "### Added\n\nDocument layout accessors.",
            1,
        )
        self.assertIn(
            "Document layout accessors.",
            workflow.render_release_notes(fenced_duplicate, "v0.8.0"),
        )

        for tag in ("script", "pre", "style", "textarea"):
            opener = f'<{tag.upper()} data-release="hidden">'
            raw_block = (
                f"{opener}\n### Added\n\nHidden raw HTML text.\n</{tag}>\n\n"
            )
            raw_duplicate = complete.replace(
                "### Added\n\nDocument layout accessors.",
                raw_block + "### Added\n\nDocument layout accessors.",
                1,
            )
            expected = stable.replace(
                "### Added\n\nDocument layout accessors.",
                raw_block + "### Added\n\nDocument layout accessors.",
                1,
            )
            with self.subTest(raw_html_duplicate=tag):
                self.assertEqual(
                    workflow.render_release_notes(raw_duplicate, "v0.8.0"),
                    expected,
                )

            hidden = f"# Changelog\n\n{opener}\n{complete}\n</{tag}>\n"
            with self.subTest(raw_html_hidden=tag), self.assertRaises(ValueError):
                workflow.render_release_notes(hidden, "v0.8.0")

        compact = "\n".join(line for line in complete.splitlines() if line.strip())
        bounded_raw_blocks = (
            ("container", '<DIV data-release="hidden">', "</DIV>"),
            ("custom-tag", '<release-note data-hidden="true">', "</release-note>"),
            ("processing-instruction", "<?release-notes", "?>"),
            ("declaration", "<!HIDDEN", ">"),
            ("cdata", "<![CDATA[", "]]>")
        )
        for name, opener, closer in bounded_raw_blocks:
            raw_block = (
                f"{opener}\n### Added\nHidden raw block text.\n{closer}\n\n"
            )
            raw_duplicate = complete.replace(
                "### Added\n\nDocument layout accessors.",
                raw_block + "### Added\n\nDocument layout accessors.",
                1,
            )
            expected = stable.replace(
                "### Added\n\nDocument layout accessors.",
                raw_block + "### Added\n\nDocument layout accessors.",
                1,
            )
            with self.subTest(bounded_raw_duplicate=name):
                self.assertEqual(
                    workflow.render_release_notes(raw_duplicate, "v0.8.0"),
                    expected,
                )

            hidden = f"# Changelog\n\n{opener}\n{compact}\n{closer}\n"
            with self.subTest(bounded_raw_hidden=name), self.assertRaises(ValueError):
                workflow.render_release_notes(hidden, "v0.8.0")

        with self.assertRaises(ValueError):
            workflow.render_release_notes(f"<!--\n{complete}\n-->\n", "v0.8.0")
        for marker in ("````", "```````", "~~~~", "~~~~~~~"):
            inner = marker[0] * 3
            fenced = f"# Changelog\n\n{marker}markdown\n{inner}\n{complete}\n{marker}\n"
            with self.subTest(fence=marker), self.assertRaises(ValueError):
                workflow.render_release_notes(fenced, "v0.8.0")

        mutations = {
            "missing-section": complete.replace("### Fixed", "### Repairs", 1),
            "duplicate-section": complete.replace(
                "### Compatibility",
                "### Added\n\nDuplicate entry.\n\n### Compatibility",
                1,
            ),
            "empty-section": complete.replace(
                "### Contributors\n\nPedro Assumpcao and the rdocx maintainers.",
                "### Contributors\n",
                1,
            ),
            "placeholder": complete.replace(
                "Stable layout access is now available.", "TBD", 1
            ),
            "duplicate-tag": complete.replace(
                "## rpptx-v0.4.0", "## v0.8.0\n\nDuplicate.\n\n## rpptx-v0.4.0", 1
            ),
            "duplicate-tag-with-trailing-space": complete.replace(
                "## rpptx-v0.4.0",
                "## v0.8.0   \n\nDuplicate.\n\n## rpptx-v0.4.0",
                1,
            ),
        }
        for name, mutated in mutations.items():
            with self.subTest(name=name), self.assertRaises(ValueError):
                workflow.render_release_notes(mutated, "v0.8.0")

        for invalid in (
            "0.8.0",
            "release-v0.8.0",
            "v0.8",
            "rpptx-v0.4.0-rc1",
            "v01.2.3",
            "v1.02.3",
            "rpptx-v1.2.03",
        ):
            with self.subTest(tag=invalid), self.assertRaises(ValueError):
                workflow.render_release_notes(complete, invalid)

    def test_release_notes_cli_check_and_render_do_not_mutate_the_source(self) -> None:
        sections = "".join(
            f"### {heading}\n\nReviewed {heading.lower()} evidence.\n\n"
            for heading in workflow.RELEASE_NOTE_HEADINGS
        )
        changelog = f"# Changelog\n\n## v1.2.3\n\n{sections}"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "CHANGELOG.md"
            path.write_text(changelog, encoding="utf-8")
            before = path.read_bytes()
            with patch.object(workflow, "CHANGELOG", path):
                with contextlib.redirect_stdout(io.StringIO()) as output:
                    result = workflow.cmd_release_notes(
                        argparse.Namespace(tag="v1.2.3", check=True, render=False)
                    )
                self.assertEqual(result, 0)
                self.assertEqual(output.getvalue(), "release-notes v1.2.3: ok\n")

                with contextlib.redirect_stdout(io.StringIO()) as output:
                    result = workflow.cmd_release_notes(
                        argparse.Namespace(tag="v1.2.3", check=False, render=True)
                    )
                self.assertEqual(result, 0)
                self.assertEqual(output.getvalue(), sections.rstrip() + "\n")
            self.assertEqual(path.read_bytes(), before)

    def test_publish_workflow_uses_only_rendered_reviewed_release_notes(self) -> None:
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        self.assert_release_notes_publish_contract(publish)

        create_step = self.yaml_step(
            self.yaml_block(publish, "  release:"),
            "Create GitHub Release from reviewed notes",
        )
        publish_job = self.yaml_block(publish, "  publish:")
        prepublish_step = self.yaml_step(
            publish_job,
            "Verify reviewed release notes",
        )
        stable_publish_step = self.yaml_step(
            publish_job,
            "Publish stable allowlist",
        )
        incubating_publish_step = self.yaml_step(
            publish_job,
            "Publish incubating allowlist",
        )
        prepublish_after_stable = publish.replace(prepublish_step, "", 1).replace(
            stable_publish_step,
            stable_publish_step + prepublish_step,
            1,
        )
        prepublish_after_incubating = publish.replace(
            prepublish_step,
            "",
            1,
        ).replace(
            incubating_publish_step,
            incubating_publish_step + prepublish_step,
            1,
        )
        render_line = (
            "          python3 scripts/sprint_workflow.py release-notes "
            '"${{ github.ref_name }}" --render > '
            '"$RUNNER_TEMP/release-notes.md"\n'
        )
        compare_line = (
            "          python3 scripts/sprint_workflow.py release-notes "
            '"${{ github.ref_name }}" --render | cmp - '
            '"$RUNNER_TEMP/release-notes.md"\n'
        )
        release_line = (
            '          gh release create "${{ github.ref_name }}" --notes-file '
            '"$RUNNER_TEMP/release-notes.md"\n'
        )
        mutations = {
            "missing-prepublish-validation": publish.replace(
                prepublish_step,
                "",
                1,
            ),
            "prepublish-validation-after-stable-publish": prepublish_after_stable,
            "prepublish-validation-after-incubating-publish": (
                prepublish_after_incubating
            ),
            "wrong-prepublish-validator-executable": publish.replace(
                prepublish_step,
                prepublish_step.replace(
                    "python3 scripts/sprint_workflow.py",
                    "echo",
                    1,
                ),
                1,
            ),
            "conditional-prepublish-validation": publish.replace(
                prepublish_step,
                prepublish_step.replace(
                    "        run:",
                    "        if: startsWith(github.ref_name, 'v')\n        run:",
                    1,
                ),
                1,
            ),
            "ignored-prepublish-validation-failure": publish.replace(
                prepublish_step,
                prepublish_step.replace(
                    "        run:",
                    "        continue-on-error: true\n        run:",
                    1,
                ),
                1,
            ),
            "generated-notes": publish.replace(
                '--notes-file "$RUNNER_TEMP/release-notes.md"',
                "--generate-notes",
                1,
            ),
            "missing-extraction": publish.replace(render_line, "", 1),
            "different-notes-file": publish.replace(
                '--notes-file "$RUNNER_TEMP/release-notes.md"',
                '--notes-file "$RUNNER_TEMP/other.md"',
                1,
            ),
            "release-before-byte-comparison": publish.replace(
                compare_line + release_line,
                release_line + compare_line,
                1,
            ),
            "overwrite-rendered-artifact": publish.replace(
                release_line,
                '          printf \'tampered\' > "$RUNNER_TEMP/release-notes.md"\n'
                + release_line,
                1,
            ),
            "wrong-validator-executable": publish.replace(
                "          python3 scripts/sprint_workflow.py release-notes",
                "          echo release-notes",
                1,
            ),
            "unreviewed-step-before-validator": publish.replace(
                create_step,
                "      - name: Replace validator\n"
                "        run: cp other.py scripts/sprint_workflow.py\n"
                + create_step,
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, publish, name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_release_notes_publish_contract(mutated)

    def assert_verify_runs_the_release_regressions(self, verify: str) -> None:
        """The local gate must run the module holding the release preflights.

        Extracted so the mutation test can assert this fails when the step is
        removed. A gate that is only wired by prose nobody checks is the defect
        F-X025 exists to close.
        """
        self.assertIn("python3 -m unittest scripts.test_sprint_workflow", verify)

    def assert_ci_runs_golden_png_gate(self, ci: str) -> None:
        job = self.yaml_block(ci, "  test:")
        poppler = self.yaml_step(job, "Install pinned Poppler 26.01.0")
        workspace = self.yaml_step(job, "Run full workspace suite")
        golden = self.yaml_step(job, "Run golden-PNG gate")
        command = "python3 scripts/golden_png_harness.py --check"
        self.assertEqual(ci.count(command), 1)
        self.assertEqual(
            self.yaml_direct_lines(golden, 8),
            (f"run: {command}",),
        )
        self.assertLess(job.index(poppler), job.index(golden))
        self.assertLess(job.index(workspace), job.index(golden))
        self.assert_no_success_short_circuit(self.operative_lines(golden))

    def test_ci_runs_the_golden_png_gate_in_the_pinned_poppler_environment(
        self,
    ) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        self.assert_ci_runs_golden_png_gate(ci)
        job = self.yaml_block(ci, "  test:")
        golden = self.yaml_step(job, "Run golden-PNG gate")
        poppler = self.yaml_step(job, "Install pinned Poppler 26.01.0")
        mutations = {
            "missing-command": ci.replace(
                "      - name: Run golden-PNG gate\n"
                "        run: python3 scripts/golden_png_harness.py --check\n",
                "",
                1,
            ),
            "before-poppler": ci.replace(golden, "", 1).replace(
                poppler, golden + poppler, 1
            ),
            "missing-check": ci.replace(
                "python3 scripts/golden_png_harness.py --check",
                "python3 scripts/golden_png_harness.py",
                1,
            ),
            "successful-fallback": ci.replace(
                "python3 scripts/golden_png_harness.py --check",
                "python3 scripts/golden_png_harness.py --check || true",
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, ci, name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_ci_runs_golden_png_gate(mutated)

    def assert_ci_runs_release_regressions(self, ci: str) -> None:
        job = self.yaml_block(ci, "  release-regressions:")
        direct = self.yaml_direct_lines(job, 4)
        self.assertEqual(
            direct,
            (
                "name: Release regressions",
                "runs-on: ubuntu-latest",
                "steps:",
            ),
        )
        steps = self.yaml_steps(job)
        self.assertEqual(len(steps), 3)
        self.assertEqual(self.yaml_step_actions(steps[0]), ("actions/checkout@v5",))
        self.assertEqual(
            self.yaml_step_identity(steps[1], 2), "Install cargo-release 1.1.3"
        )
        self.assertEqual(
            self.yaml_direct_lines(steps[1], 8),
            ("run: cargo install cargo-release --version 1.1.3 --locked",),
        )
        self.assertEqual(self.yaml_step_identity(steps[2], 2), "Run release regressions")
        self.assertEqual(
            self.yaml_direct_lines(steps[2], 8),
            ("run: python3 -m unittest scripts.test_sprint_workflow",),
        )
        self.assertNotIn("continue-on-error", job)
        self.assert_no_success_short_circuit(self.operative_lines(steps[1]))
        self.assert_no_success_short_circuit(self.operative_lines(steps[2]))

    def test_ci_runs_release_regressions_in_a_named_job(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )

        self.assert_ci_runs_release_regressions(ci)

    def test_ci_release_regression_job_rejects_wiring_mutations(self) -> None:
        ci = (workflow.REPO / ".github/workflows/ci.yml").read_text(
            encoding="utf-8"
        )
        install_step = (
            "      - name: Install cargo-release 1.1.3\n"
            "        run: cargo install cargo-release --version 1.1.3 --locked\n"
        )
        run_step = (
            "      - name: Run release regressions\n"
            "        run: python3 -m unittest scripts.test_sprint_workflow\n"
        )
        mutations = {
            "removed-cargo-release-install": ci.replace(install_step, "", 1),
            "cargo-release-installed-after-regressions": ci.replace(
                install_step + run_step, run_step + install_step, 1
            ),
            "removed-command": ci.replace(
                "python3 -m unittest scripts.test_sprint_workflow", "", 1
            ),
            "narrowed-command": ci.replace(
                "python3 -m unittest scripts.test_sprint_workflow",
                "python3 -m unittest "
                "scripts.test_sprint_workflow.SprintWorkflowTests."
                "test_stable_release_family_is_prepared_at_0_13_0",
                1,
            ),
            "job-condition": ci.replace(
                "  release-regressions:\n",
                "  release-regressions:\n    if: false\n",
                1,
            ),
            "continue-on-error": ci.replace(
                "        run: python3 -m unittest scripts.test_sprint_workflow\n",
                "        continue-on-error: true\n"
                "        run: python3 -m unittest scripts.test_sprint_workflow\n",
                1,
            ),
            "successful-fallback": ci.replace(
                "python3 -m unittest scripts.test_sprint_workflow",
                "python3 -m unittest scripts.test_sprint_workflow || true",
                1,
            ),
        }
        for name, mutated in mutations.items():
            self.assertNotEqual(mutated, ci, name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_ci_runs_release_regressions(mutated)

    def test_verify_runs_the_release_regressions(self) -> None:
        # Without this, the preflights that publish.yml invokes by name run for
        # the first time on a tag, after the sprint is closed. S42 is the
        # demonstration: F-X022 moved every version carrier under crates/,
        # passed the whole local gate, and left the incubating preflight and the
        # ci.yml WASM literal asserting the old version.
        verify = (workflow.REPO / ".claude/commands/verify.md").read_text(
            encoding="utf-8"
        )

        self.assert_verify_runs_the_release_regressions(verify)

        removed = verify.replace(
            "python3 -m unittest scripts.test_sprint_workflow", "", 1
        )
        self.assertNotEqual(removed, verify)
        with self.assertRaises(AssertionError):
            self.assert_verify_runs_the_release_regressions(removed)

    def assert_verify_checks_both_wasm_targets(self, verify: str) -> None:
        normalized = " ".join(verify.split())
        self.assertIn(
            "cargo check --target wasm32-unknown-unknown -p rdocx-wasm "
            "-p rpptx-wasm",
            normalized,
        )

    def test_verify_checks_both_wasm_targets(self) -> None:
        verify = (workflow.REPO / ".claude/commands/verify.md").read_text(
            encoding="utf-8"
        )
        self.assert_verify_checks_both_wasm_targets(verify)

        mutated = verify.replace(" -p rpptx-wasm", "", 1)
        self.assertNotEqual(mutated, verify)
        with self.assertRaises(AssertionError):
            self.assert_verify_checks_both_wasm_targets(mutated)

    def test_every_test_publish_yml_names_resolves_to_a_real_test(self) -> None:
        # publish.yml invokes preflights by their full dotted path. A rename
        # breaks publication on a tag, which is the worst moment to find out.
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        module = "scripts.test_sprint_workflow."
        named = {
            token.removeprefix(module)
            for token in publish.split()
            if token.startswith(module)
        }

        self.assertTrue(named, "publish.yml names no test in this module")
        for path in sorted(named):
            class_name, _, method_name = path.partition(".")
            with self.subTest(path):
                # A dotted path publish.yml can invoke has to resolve at this
                # module's top level, which is the only place `python3 -m
                # unittest scripts.test_sprint_workflow.<Class>.<method>` looks.
                cls = globals().get(class_name)
                self.assertIsNotNone(
                    cls,
                    f"publish.yml names {path}, and this module defines no "
                    f"top-level {class_name}",
                )
                self.assertTrue(
                    callable(getattr(cls, method_name, None)),
                    f"publish.yml names {path}, and {class_name} has no "
                    f"{method_name}",
                )

    def test_completed_shared_and_powerpoint_crates_are_publication_candidates(
        self,
    ) -> None:
        incubating_packages = (
            "oxml-core",
            "oxml-opc",
            "oxml-media",
            "oxml-layout",
            "oxml-drawing",
            "oxml-pdf",
            "oxml-sml",
            "oxml-cli-support",
            "oxml-chart",
            "rpptx-oxml",
            "rpptx-chart",
            "rpptx-layout",
            "rpptx-render",
            "rpptx",
        )

        for name in incubating_packages:
            manifest = tomllib.loads(
                (workflow.REPO / f"crates/{name}/Cargo.toml").read_text(
                    encoding="utf-8"
                )
            )
            self.assertIs(manifest["package"].get("publish"), True, name)

    def test_chart_ownership_uses_shared_crate_and_exact_legacy_shim(self) -> None:
        root_text = (workflow.REPO / "Cargo.toml").read_text(encoding="utf-8")
        root = tomllib.loads(root_text)["workspace"]
        self.assertIn("crates/oxml-chart", root["members"])
        self.assertEqual(
            root["dependencies"]["oxml-chart"],
            {"path": "crates/oxml-chart", "version": "0.11.0"},
        )

        shim = tomllib.loads(
            (workflow.REPO / "crates/rpptx-chart/Cargo.toml").read_text(
                encoding="utf-8"
            )
        )
        self.assertEqual(shim["package"]["description"], "deprecated: moved to oxml-chart")
        self.assertEqual(shim["dependencies"], {"oxml-chart": {"workspace": True}})

        for consumer in ("rpptx", "rpptx-layout"):
            manifest = tomllib.loads(
                (workflow.REPO / f"crates/{consumer}/Cargo.toml").read_text(
                    encoding="utf-8"
                )
            )
            self.assertIn("oxml-chart", manifest["dependencies"], consumer)
            self.assertNotIn("rpptx-chart", manifest["dependencies"], consumer)

        for path in (
            "crates/rpptx/src/lib.rs",
            "crates/rpptx-layout/src/context.rs",
            "crates/rpptx-layout/src/lib.rs",
        ):
            source = (workflow.REPO / path).read_text(encoding="utf-8")
            self.assertIn("oxml_chart", source, path)
            self.assertNotIn("rpptx_chart", source, path)

    def test_publish_workflow_routes_exact_dependency_ordered_allowlists(self) -> None:
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        self.assert_publish_workflow_contract(publish)

    def test_publish_workflow_rejects_swapped_namespace_predicates(self) -> None:
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        mutated = publish.replace(
            "if: startsWith(github.ref_name, 'v')",
            "if: TEMPORARY_PREDICATE",
            1,
        )
        mutated = mutated.replace(
            "if: startsWith(github.ref_name, 'rpptx-v')",
            "if: startsWith(github.ref_name, 'v')",
            1,
        ).replace(
            "if: TEMPORARY_PREDICATE",
            "if: startsWith(github.ref_name, 'rpptx-v')",
            1,
        )

        with self.assertRaises(AssertionError):
            self.assert_publish_workflow_contract(mutated)

    def test_publish_workflow_rejects_an_extra_package(self) -> None:
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        mutated = publish.replace(
            "\n  release:\n",
            "\n      - name: Publish an extra package\n"
            "        if: startsWith(github.ref_name, 'v')\n"
            "        run: |\n"
            "          cargo publish -p rdocx-wasm\n"
            "\n  release:\n",
            1,
        )

        with self.assertRaises(AssertionError):
            self.assert_publish_workflow_contract(mutated)

    def test_publish_workflow_rejects_continue_on_error(self) -> None:
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        mutated = publish.replace(
            "      - name: Publish stable allowlist\n",
            "      - name: Publish stable allowlist\n"
            "        continue-on-error: true\n",
            1,
        )

        with self.assertRaises(AssertionError):
            self.assert_publish_workflow_contract(mutated)

    def test_publish_workflow_rejects_successful_fallback_commands(self) -> None:
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        mutated = publish.replace(
            "          cargo publish -p rdocx-opc\n",
            "          cargo publish -p rdocx-opc || true\n",
            1,
        )

        with self.assertRaises(AssertionError):
            self.assert_publish_workflow_contract(mutated)

    def test_publish_workflow_preflights_and_propagates_failures(self) -> None:
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        stable_check = (
            "scripts.test_sprint_workflow.SprintWorkflowTests."
            "test_stable_release_family_is_prepared_at_0_13_0"
        )
        incubating_check = (
            "scripts.test_sprint_workflow.SprintWorkflowTests."
            "test_incubating_release_family_is_prepared_at_0_11_0"
        )
        packaged_registry_check = (
            "scripts.test_sprint_workflow.SprintWorkflowTests."
            "test_prepared_rdocx_layout_0_13_0_requires_published_oxml_layout_0_10_0"
        )
        historical_registry_check = (
            "scripts.test_sprint_workflow.SprintWorkflowTests."
            "test_immutable_rdocx_layout_0_10_1_registry_graph_remains_at_oxml_layout_0_6_0"
        )
        metadata_command = (
            "python3 -m unittest "
            f"{stable_check} {historical_registry_check} {incubating_check}"
        )
        stable_registry_command = f"python3 -m unittest {packaged_registry_check}"

        def assert_stable_registry_step(candidate: str) -> None:
            stable_step = self.yaml_step(
                self.yaml_block(candidate, "  publish:"),
                "Verify published shared family for stable release",
            )
            self.assertEqual(candidate.count(stable_registry_command), 1)
            self.assertIn("if: startsWith(github.ref_name, 'v')", stable_step)
            self.assertIn('RDOCX_VERIFY_PUBLISHED_SHARED: "1"', stable_step)
            self.assertLess(
                candidate.index(metadata_command),
                candidate.index(stable_registry_command),
            )
            self.assertLess(
                candidate.index(stable_registry_command),
                candidate.index("cargo publish --workspace --dry-run"),
            )

        self.assert_publish_preflight_contract(publish)
        self.assertEqual(publish.count(metadata_command), 1)
        assert_stable_registry_step(publish)
        self.assertLess(
            publish.index(stable_check), publish.index(historical_registry_check)
        )
        self.assertLess(
            publish.index(historical_registry_check), publish.index(incubating_check)
        )
        self.assertLess(
            publish.index("python3 scripts/hash_harness.py --check"),
            publish.index(metadata_command),
        )
        self.assertLess(
            publish.index("cargo publish --workspace --dry-run"),
            publish.index("cargo publish -p rdocx-opc"),
        )
        self.assertNotIn("--no-verify", publish)
        self.assertNotIn("continue-on-error", publish)

        mutations = (
            (
                "missing-stable-only-condition",
                publish.replace(
                    "        if: startsWith(github.ref_name, 'v')\n",
                    "",
                    1,
                ),
            ),
            (
                "missing-published-shared-authority",
                publish.replace(
                    '          RDOCX_VERIFY_PUBLISHED_SHARED: "1"\n',
                    "",
                    1,
                ),
            ),
            (
                "wrong-published-shared-authority",
                publish.replace(
                    '          RDOCX_VERIFY_PUBLISHED_SHARED: "1"',
                    '          RDOCX_VERIFY_PUBLISHED_SHARED: "0"',
                    1,
                ),
            ),
        )
        for name, mutated in mutations:
            self.assertNotEqual(mutated, publish, name)
            with self.subTest(name=name):
                with self.assertRaises(AssertionError):
                    assert_stable_registry_step(mutated)

    def test_publish_workflow_rejects_a_missing_local_patch(self) -> None:
        publish = (workflow.REPO / ".github/workflows/publish.yml").read_text(
            encoding="utf-8"
        )
        mutated = publish.replace(
            "            --config 'patch.crates-io.oxml-core.path=\"crates/oxml-core\"' \\\n",
            "",
            1,
        )

        with self.assertRaises(AssertionError):
            self.assert_publish_preflight_contract(mutated)

    def test_review_and_verification_evidence_is_bound_to_head(self) -> None:
        data = {
            "reviews": [{"pass": 4, "blocking": 0, "head": "current"}],
            "verifications": [
                {"scope": "full", "passed": True, "head": "current"}
            ],
        }
        self.assertEqual(workflow.closure_evidence_problems(data, "current"), [])

        data["reviews"][-1]["head"] = "reviewed-old"
        data["verifications"][-1]["head"] = "verified-old"
        self.assertEqual(
            workflow.closure_evidence_problems(data, "current"),
            [
                "latest sprint review covered reviewed-old, current HEAD is current",
                "no passing `/verify --full` recorded for current HEAD current",
            ],
        )

    def test_recorded_evidence_captures_head(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            scratch = Path(directory)
            state = {
                "schema_version": workflow.SCHEMA_VERSION,
                "sprint": "S01",
                "phase": "review",
                "max_review_passes": 3,
                "features": {},
                "reviews": [],
                "verifications": [],
            }
            (scratch / "S01-run.json").write_text(json.dumps(state), encoding="utf-8")
            review_args = argparse.Namespace(
                sprint="S01",
                passno=4,
                blocking=0,
                should_fix=0,
                nice_to_have=0,
                extend=True,
            )
            verify_args = argparse.Namespace(
                sprint="S01",
                scope="full",
                passed=True,
                harness="unchanged",
            )

            with (
                patch.object(workflow, "SCRATCH", scratch),
                patch.object(workflow, "git_head", return_value="abc123"),
            ):
                workflow.cmd_record_review(review_args)
                workflow.cmd_record_verification(verify_args)

            saved = json.loads((scratch / "S01-run.json").read_text(encoding="utf-8"))
            self.assertEqual(saved["reviews"][-1]["head"], "abc123")
            self.assertEqual(saved["verifications"][-1]["head"], "abc123")

    def test_run_sprint_phase_sequence_is_accepted(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            scratch = Path(directory)
            state = {
                "schema_version": workflow.SCHEMA_VERSION,
                "sprint": "S01",
                "phase": "design",
                "features": {},
                "reviews": [],
                "verifications": [],
            }
            (scratch / "S01-run.json").write_text(json.dumps(state), encoding="utf-8")

            with patch.object(workflow, "SCRATCH", scratch):
                for phase in (
                    "questions",
                    "implementation",
                    "integration",
                    "verification",
                    "review",
                    "implementation",
                    "integration",
                    "verification",
                    "review",
                    "implementation",
                    "integration",
                    "verification",
                    "review",
                    "ready_to_close",
                ):
                    workflow.cmd_set_phase(argparse.Namespace(sprint="S01", phase=phase))
                    saved = json.loads((scratch / "S01-run.json").read_text(encoding="utf-8"))
                    self.assertEqual(saved["phase"], phase)

    def test_run_sprint_ordinary_dependency_chain_is_accepted(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            scratch = Path(directory)
            state = {
                "schema_version": workflow.SCHEMA_VERSION,
                "sprint": "S01",
                "phase": "implementation",
                "features": {
                    fid: {
                        "state": "approved",
                        "size": "S",
                        "title": title,
                        "owner": "codex",
                    }
                    for fid, title in (
                        ("F-001", "A"),
                        ("F-002", "B depends on A"),
                        ("F-003", "C depends on B"),
                    )
                },
                "reviews": [],
                "verifications": [],
            }
            (scratch / "S01-run.json").write_text(json.dumps(state), encoding="utf-8")

            def mark(fid: str, feature_state: str) -> None:
                workflow.cmd_mark_feature(
                    argparse.Namespace(
                        sprint="S01",
                        fid=fid,
                        state=feature_state,
                        owner=None,
                        clear_owner=feature_state == "completed",
                    )
                )

            with patch.object(workflow, "SCRATCH", scratch):
                for fid in ("F-001", "F-002"):
                    mark(fid, "running")
                    mark(fid, "reviewed")
                    for phase in ("integration", "verification", "review"):
                        workflow.cmd_set_phase(
                            argparse.Namespace(sprint="S01", phase=phase)
                        )
                    mark(fid, "completed")
                    workflow.cmd_set_phase(
                        argparse.Namespace(sprint="S01", phase="implementation")
                    )
                mark("F-003", "running")

            saved = json.loads((scratch / "S01-run.json").read_text(encoding="utf-8"))
            self.assertEqual(saved["features"]["F-001"]["state"], "completed")
            self.assertEqual(saved["features"]["F-002"]["state"], "completed")
            self.assertEqual(saved["features"]["F-003"]["state"], "running")
            self.assertEqual(saved["phase"], "implementation")

    def test_completed_feature_requires_every_delivery_record(self) -> None:
        with tempfile.TemporaryDirectory(dir=workflow.REPO) as directory:
            root = Path(directory)
            current = root / "CURRENT_SPRINT.md"
            backlog = root / "BACKLOG.md"
            tracker = root / "SPRINT_TRACKER.md"
            as_built = root / "AS_BUILT.md"
            plans = root / "plans"
            plans.mkdir()
            current.write_text(
                "# Current Sprint, S01\n\n"
                "| F-ID | Title | Size | Status | Owner |\n"
                "|---|---|---|---|---|\n"
                "| F-001 | Example | S | done | - |\n",
                encoding="utf-8",
            )
            backlog.write_text(
                "| F-ID | Title | Sprint | Size | Status |\n"
                "|---|---|---|---|---|\n"
                "| F-001 | Example | S01 | S | done |\n",
                encoding="utf-8",
            )
            tracker.write_text("| F-001 | S01 | S | 1 | 1 | date | note |\n", encoding="utf-8")
            as_built.write_text("### F-001, Example\n", encoding="utf-8")
            (plans / "F-001-design.md").write_text(
                "**Status**: completed\n", encoding="utf-8"
            )

            with patch.multiple(
                workflow,
                CURRENT_SPRINT=current,
                BACKLOG=backlog,
                SPRINT_TRACKER=tracker,
                AS_BUILT=as_built,
                PLANS=plans,
            ):
                self.assertEqual(workflow.completed_record_problems("S01", "F-001"), [])
                current.write_text(
                    "# Current Sprint, S01\n\n"
                    "| F-ID | Title | Size | Status | Owner |\n"
                    "|---|---|---|---|---|\n"
                    "| F-001 | Example | S | done | |\n",
                    encoding="utf-8",
                )
                self.assertEqual(workflow.completed_record_problems("S01", "F-001"), [])
                current.write_text(
                    "# Current Sprint, S01\n\n"
                    "| F-ID | Title | Size | Status | Owner |\n"
                    "|---|---|---|---|---|\n"
                    "| F-001 | Example | S | done | codex |\n",
                    encoding="utf-8",
                )
                self.assertEqual(
                    workflow.completed_record_problems("S01", "F-001"),
                    ["F-001 is completed but CURRENT_SPRINT.md owner is 'codex'"],
                )
                tracker.write_text("", encoding="utf-8")
                current.write_text(
                    "# Current Sprint, S01\n\n"
                    "| F-ID | Title | Size | Status | Owner |\n"
                    "|---|---|---|---|---|\n"
                    "| F-001 | Example | S | done | |\n",
                    encoding="utf-8",
                )
                self.assertEqual(
                    workflow.completed_record_problems("S01", "F-001"),
                    ["F-001 has no S01 row in SPRINT_TRACKER.md"],
                )

    def test_completed_run_state_requires_a_cleared_owner(self) -> None:
        data = {
            "features": {
                "F-001": {"state": "completed", "owner": None},
                "F-002": {"state": "completed", "owner": "codex"},
                "F-003": {"state": "carried", "owner": "claude"},
            }
        }

        self.assertEqual(
            workflow.completed_owner_problems(data),
            ["F-002 is completed but run-state owner is 'codex'"],
        )

    def test_close_preflight_rejects_a_completed_run_state_owner(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            reviews = Path(directory)
            (reviews / "S01-sprint-review-pass-1.md").write_text(
                "**Verdict**: 0 blocking, 0 should-fix, 0 nice-to-have\n",
                encoding="utf-8",
            )
            data = {
                "phase": "review",
                "features": {
                    "F-001": {"state": "completed", "owner": "codex"},
                    "F-002": {"state": "carried", "owner": "claude"},
                },
                "reviews": [{"pass": 1, "blocking": 0, "head": "current"}],
                "verifications": [
                    {
                        "scope": "full",
                        "passed": True,
                        "harness": "unchanged",
                        "head": "current",
                    }
                ],
            }

            with (
                patch.object(workflow, "load", return_value=data),
                patch.object(workflow, "git_head", return_value="current"),
                patch.object(workflow, "HANDOFFS", reviews / "handoffs"),
                patch.object(workflow, "REVIEWS", reviews),
                patch.object(workflow, "backlog_statuses", return_value={"F-001": "done"}),
                patch.object(workflow, "completed_record_problems", return_value=[]),
                contextlib.redirect_stderr(io.StringIO()),
            ):
                self.assertEqual(
                    workflow.cmd_close_preflight(argparse.Namespace(sprint="S01")),
                    1,
                )
                data["features"]["F-001"]["owner"] = None
                self.assertEqual(
                    workflow.cmd_close_preflight(argparse.Namespace(sprint="S01")),
                    0,
                )

    def repository_path_claims(self, source: str) -> set[str]:
        rooted = re.findall(
            r"(?<![A-Za-z0-9_-])"
            r"((?:\.agents|\.claude|\.github|crates|docs|samples|scripts|target|tools)"
            r"(?:/[A-Za-z0-9_.<>{}*?$-]*)+)",
            source,
        )
        standalone = re.findall(
            r"(?<![/A-Za-z0-9_.-])"
            r"([A-Za-z0-9*?][A-Za-z0-9_.*?-]*\."
            r"(?:crate|json|lock|md|pptx|py|rs|sh|toml|tsv|ttf|yaml|yml)"
            r"(?::[0-9]+(?:-[0-9]+)?)?)(?![A-Za-z0-9_.-])",
            source,
        )
        return {claim.rstrip(".,:;)") for claim in rooted + standalone}

    def assert_repository_path_claims_resolve(
        self,
        source: str,
        *,
        generated_claims: set[str],
    ) -> set[str]:
        claims = self.repository_path_claims(source)
        self.assertTrue(claims)
        tracked = subprocess.run(
            ["git", "ls-files"],
            cwd=workflow.REPO,
            check=True,
            capture_output=True,
            text=True,
        ).stdout.splitlines()
        tracked_names = {Path(path).name for path in tracked}
        found_generated: set[str] = set()
        for claim in claims:
            path = re.sub(r":[0-9]+(?:-[0-9]+)?$", "", claim)
            if path in generated_claims:
                found_generated.add(path)
                continue
            if path.startswith(("samples/", "target/")):
                self.fail(f"unrecognised generated path claim: {claim}")
            if any(marker in path for marker in ("*", "?")):
                self.assertTrue(tuple(workflow.REPO.glob(path)), claim)
                continue
            if any(marker in path for marker in ("<", "{", "$")):
                static_prefix = re.split(r"[<{$]", path, maxsplit=1)[0].rstrip("/")
                self.assertTrue(static_prefix, claim)
                self.assertTrue((workflow.REPO / static_prefix).exists(), claim)
                continue
            if "/" in path:
                self.assertTrue((workflow.REPO / path.rstrip("/")).exists(), claim)
                continue
            self.assertIn(path, tracked_names, claim)
        self.assertEqual(found_generated, generated_claims)
        return claims

    def assert_agent_facing_repository_claims(
        self,
        *,
        claude: str | None = None,
        verify: str | None = None,
    ) -> None:
        if claude is None:
            claude = (workflow.REPO / "CLAUDE.md").read_text(encoding="utf-8")
        if verify is None:
            verify = (workflow.REPO / ".claude/commands/verify.md").read_text(
                encoding="utf-8"
            )
        root = tomllib.loads(
            (workflow.REPO / "Cargo.toml").read_text(encoding="utf-8")
        )
        workspace = root["workspace"]
        workspace_version = workspace["package"]["version"]
        packages: dict[str, tuple[Path, dict[str, object]]] = {}
        for member in workspace["members"]:
            member_path = workflow.REPO / member
            manifest = tomllib.loads(
                (member_path / "Cargo.toml").read_text(encoding="utf-8")
            )
            packages[manifest["package"]["name"]] = (member_path, manifest)

        claude_paths = self.assert_repository_path_claims_resolve(
            claude,
            generated_claims={"samples/"},
        )
        verify_paths = self.assert_repository_path_claims_resolve(
            verify,
            generated_claims={"*.crate", "target/package"},
        )
        self.assertIn("docs/hld/00-vision.md", claude_paths)
        self.assertIn(".github/workflows/publish.yml", verify_paths)
        self.assertIn("scripts/hash_harness.py", verify_paths)

        stated_versions = re.findall(
            r"prepared\s+at\s+([0-9]+\.[0-9]+\.[0-9]+) "
            r"across the exact seven-package stable family",
            claude,
        )
        self.assertEqual(stated_versions, [workspace_version])
        self.assertIn(
            "attempt published only `rdocx-opc` and `rdocx-oxml`",
            claude,
        )
        for name, (_, manifest) in packages.items():
            if not name.startswith("rdocx") or name == "rdocx-py":
                continue
            version = manifest["package"]["version"]
            effective_version = (
                workspace_version
                if isinstance(version, dict) and version.get("workspace")
                else version
            )
            self.assertEqual(effective_version, workspace_version, name)

        font_claims = re.findall(r"`(crates/[^`]+/fonts/)`", claude)
        self.assertEqual(font_claims, ["crates/oxml-layout/fonts/"])
        font_package = Path(font_claims[0]).parts[1]
        font_path, font_manifest = packages[font_package]
        features = font_manifest.get("features", {})
        self.assertIn("system-fonts", features)
        self.assertNotIn("bundled-fonts", features)
        claimed_font_count = re.findall(r"([0-9]+) bundled TTFs", claude)
        self.assertEqual(claimed_font_count, ["24"])
        fonts = font_path / "fonts"
        self.assertEqual(len(tuple(fonts.glob("*.ttf"))), int(claimed_font_count[0]))
        for legal_file in (
            "LICENSE-Caladea",
            "NOTICE-Caladea",
            "LICENSE-Carlito",
            "LICENSE-Liberation",
            "LICENSE-Noto",
            "NOTICE-Noto",
            "SUBSET-NotoSansSC.md",
        ):
            self.assertTrue((fonts / legal_file).is_file(), legal_file)

        feature_claims = set(re.findall(r"`([a-z0-9-]+)` feature", claude))
        self.assertIn("system-fonts", feature_claims)
        available_features = {
            feature
            for _, manifest in packages.values()
            for feature in manifest.get("features", {})
        }
        self.assertLessEqual(feature_claims, available_features)

        verify_packages = re.findall(
            r"(?:-p|--package) ([a-z0-9][a-z0-9-]+)", verify
        )
        self.assertTrue(verify_packages)
        for package in verify_packages:
            self.assertIn(package, packages)
        no_default_packages = re.findall(
            r"cargo test -p ([a-z0-9-]+) --no-default-features", verify
        )
        self.assertEqual(no_default_packages, ["oxml-layout"])

    def test_agent_facing_repository_claims_resolve_against_the_workspace(
        self,
    ) -> None:
        self.assert_agent_facing_repository_claims()

    def test_agent_facing_claim_contract_rejects_stale_mutations(self) -> None:
        claude = (workflow.REPO / "CLAUDE.md").read_text(encoding="utf-8")
        verify = (workflow.REPO / ".claude/commands/verify.md").read_text(
            encoding="utf-8"
        )
        mutations = {
            "path": (
                claude.replace(
                    "crates/oxml-layout/fonts/",
                    "crates/rdocx-layout/fonts/",
                    1,
                ),
                verify,
            ),
            "non-crates-path": (
                claude.replace(
                    "docs/hld/00-vision.md",
                    "docs/hld/00-missing.md",
                    1,
                ),
                verify,
            ),
            "verify-path": (
                claude,
                verify.replace(
                    ".github/workflows/publish.yml",
                    ".github/workflows/missing.yml",
                    1,
                ),
            ),
            "version": (
                claude.replace(
                    "prepared at\n  0.13.0",
                    "prepared at\n  0.2.0",
                    1,
                ),
                verify,
            ),
            "stale-prepared-version": (
                claude.replace(
                    "attempt published only `rdocx-opc` and `rdocx-oxml`",
                    "attempt published only `rdocx-opc`",
                    1,
                ),
                verify,
            ),
            "feature": (
                claude.replace("`system-fonts` feature", "`bundled-fonts` feature", 1),
                verify,
            ),
            "package": (
                claude,
                verify.replace(
                    "cargo test -p oxml-layout --no-default-features",
                    "cargo test -p legacy-layout --no-default-features",
                    1,
                ),
            ),
        }
        for name, (mutated_claude, mutated_verify) in mutations.items():
            self.assertNotEqual((mutated_claude, mutated_verify), (claude, verify), name)
            with self.subTest(name=name), self.assertRaises(AssertionError):
                self.assert_agent_facing_repository_claims(
                    claude=mutated_claude,
                    verify=mutated_verify,
                )


if __name__ == "__main__":
    unittest.main()
