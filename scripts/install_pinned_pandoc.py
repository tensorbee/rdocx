#!/usr/bin/env python3
"""Install the exact Pandoc texmath oracle used by CI."""

from __future__ import annotations

import argparse
import hashlib
import os
from pathlib import Path
import platform
import shutil
import stat
import subprocess
import tarfile
import tempfile
import urllib.request


PANDOC_VERSION = "3.10"
PANDOC_SHA256 = "e0f8af62d0f267d22baa5bcefe6d5dda3a097ccc60de794b759fe03159923244"
PANDOC_ARCHIVE = f"pandoc-{PANDOC_VERSION}-linux-amd64.tar.gz"
PANDOC_URL = (
    f"https://github.com/jgm/pandoc/releases/download/{PANDOC_VERSION}/{PANDOC_ARCHIVE}"
)
ARCHIVE_ROOT = f"pandoc-{PANDOC_VERSION}"
MAX_DOWNLOAD_BYTES = 40 * 1024 * 1024
MAX_ARCHIVE_MEMBERS = 256
MEBIBYTE = 1024 * 1024
MAX_EXTRACTED_BYTES = 128 * MEBIBYTE


def download_archive(destination: Path) -> None:
    digest = hashlib.sha256()
    written = 0
    request = urllib.request.Request(
        PANDOC_URL,
        headers={"User-Agent": "rdocx-ci-pandoc-installer/1"},
    )
    with urllib.request.urlopen(request, timeout=60) as response, destination.open(
        "wb"
    ) as output:
        while chunk := response.read(MEBIBYTE):
            written += len(chunk)
            if written > MAX_DOWNLOAD_BYTES:
                raise RuntimeError("Pandoc archive exceeds the download bound")
            digest.update(chunk)
            output.write(chunk)
    if digest.hexdigest() != PANDOC_SHA256:
        raise RuntimeError("Pandoc archive SHA-256 does not match the reviewed source")


def safe_extract(archive_path: Path, destination: Path) -> Path:
    destination = destination.resolve()
    destination.mkdir(parents=True)
    member_count = 0
    extracted_bytes = 0
    with tarfile.open(archive_path, mode="r|gz") as archive:
        for member in archive:
            member_count += 1
            if member_count > MAX_ARCHIVE_MEMBERS:
                raise RuntimeError("Pandoc archive exceeds the member-count bound")
            if member.size < 0:
                raise RuntimeError("Pandoc archive contains a negative member size")
            extracted_bytes += member.size
            if extracted_bytes > MAX_EXTRACTED_BYTES:
                raise RuntimeError("Pandoc archive exceeds the extracted-size bound")
            member_path = Path(member.name)
            if not member_path.parts or member_path.parts[0] != ARCHIVE_ROOT:
                raise RuntimeError("Pandoc archive has an unexpected root layout")
            target = (destination / member_path).resolve()
            try:
                target.relative_to(destination)
            except ValueError as error:
                raise RuntimeError("Pandoc archive contains an unsafe path") from error
            if member.isdir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            if not member.isfile():
                raise RuntimeError("Pandoc archive contains a non-file entry")
            source = archive.extractfile(member)
            if source is None:
                raise RuntimeError("Pandoc archive member could not be read")
            target.parent.mkdir(parents=True, exist_ok=True)
            with source, target.open("wb") as output:
                shutil.copyfileobj(source, output)
            target.chmod(stat.S_IMODE(member.mode))
    executable = destination / ARCHIVE_ROOT / "bin" / "pandoc"
    if not executable.is_file():
        raise RuntimeError("Pandoc archive does not contain the expected executable")
    return executable


def verify_pandoc(executable: Path) -> None:
    result = subprocess.run(
        [str(executable), "--version"],
        check=True,
        capture_output=True,
        text=True,
    )
    identity = result.stdout.splitlines()
    if not identity or identity[0] != f"pandoc {PANDOC_VERSION}":
        raise RuntimeError(f"unexpected Pandoc identity: {identity[:1]}")


def expose_pandoc(executable: Path) -> None:
    github_path = os.environ.get("GITHUB_PATH")
    if github_path:
        with Path(github_path).open("a", encoding="utf-8") as path_file:
            path_file.write(f"{executable.parent}\n")
    github_env = os.environ.get("GITHUB_ENV")
    if github_env:
        with Path(github_env).open("a", encoding="utf-8") as env_file:
            env_file.write(f"RDOCX_PANDOC={executable}\n")


def install(prefix: Path) -> Path:
    if platform.system() != "Linux" or platform.machine() not in ("x86_64", "amd64"):
        raise RuntimeError("the pinned Pandoc archive supports Linux x86-64 only")
    prefix = prefix.resolve()
    executable = prefix / "bin" / "pandoc"
    if prefix.exists() and (not prefix.is_dir() or any(prefix.iterdir())):
        raise RuntimeError("Pandoc prefix must be absent or empty")
    runner_temp = Path(os.environ.get("RUNNER_TEMP", tempfile.gettempdir()))
    with tempfile.TemporaryDirectory(prefix="rdocx-pandoc-", dir=runner_temp) as work:
        work_root = Path(work)
        archive_path = work_root / PANDOC_ARCHIVE
        download_archive(archive_path)
        extracted = safe_extract(archive_path, work_root / "source")
        executable.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(extracted, executable)
    verify_pandoc(executable)
    expose_pandoc(executable)
    print(f"Installed Pandoc {PANDOC_VERSION} at {executable}")
    return executable


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    default_prefix = Path(
        os.environ.get("RUNNER_TEMP", tempfile.gettempdir())
    ) / f"pandoc-{PANDOC_VERSION}"
    parser.add_argument("--prefix", type=Path, default=default_prefix)
    arguments = parser.parse_args()
    install(arguments.prefix)


if __name__ == "__main__":
    main()
