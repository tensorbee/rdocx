#!/usr/bin/env python3
"""Install the exact LibreOffice Linux viewer oracle used by CI."""

from __future__ import annotations

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


LIBREOFFICE_VERSION = "26.2.5.2"
LIBREOFFICE_SHA256 = "2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e"
LIBREOFFICE_ARCHIVE = "LibreOffice_26.2.5.2_Linux_x86-64_deb.tar.gz"
LIBREOFFICE_URL = (
    "https://downloadarchive.documentfoundation.org/libreoffice/old/26.2.5.2/"
    f"deb/x86_64/{LIBREOFFICE_ARCHIVE}"
)
ARCHIVE_ROOT = f"LibreOffice_{LIBREOFFICE_VERSION}_Linux_x86-64_deb"
INSTALL_ROOT = Path("/opt/libreoffice26.2")
SOFFICE = INSTALL_ROOT / "program/soffice"
EXPECTED_IDENTITY = (
    "LibreOffice 26.2.5.2 cd7284b4cbbfeb507e630c1aac019f4157393acb"
)
MAX_DOWNLOAD_BYTES = 224 * 1024 * 1024
MAX_ARCHIVE_MEMBERS = 256
MAX_EXTRACTED_BYTES = 256 * 1024 * 1024
SYSTEM_RUNTIME_PACKAGES = (
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
)


def download_archive(destination: Path) -> None:
    digest = hashlib.sha256()
    written = 0
    request = urllib.request.Request(
        LIBREOFFICE_URL,
        headers={"User-Agent": "rdocx-ci-libreoffice-installer/1"},
    )
    with urllib.request.urlopen(request, timeout=60) as response, destination.open(
        "wb"
    ) as output:
        while chunk := response.read(1024 * 1024):
            written += len(chunk)
            if written > MAX_DOWNLOAD_BYTES:
                raise RuntimeError("LibreOffice archive exceeds the download bound")
            digest.update(chunk)
            output.write(chunk)
    if digest.hexdigest() != LIBREOFFICE_SHA256:
        raise RuntimeError(
            "LibreOffice archive SHA-256 does not match the reviewed source"
        )


def safe_extract(archive_path: Path, destination: Path) -> Path:
    destination = destination.resolve()
    destination.mkdir(parents=True)
    member_count = 0
    extracted_bytes = 0
    with tarfile.open(archive_path, mode="r|gz") as archive:
        for member in archive:
            member_count += 1
            if member_count > MAX_ARCHIVE_MEMBERS:
                raise RuntimeError("LibreOffice archive exceeds the member-count bound")
            if member.size < 0:
                raise RuntimeError("LibreOffice archive contains a negative member size")
            extracted_bytes += member.size
            if extracted_bytes > MAX_EXTRACTED_BYTES:
                raise RuntimeError("LibreOffice archive exceeds the extracted-size bound")
            member_path = Path(member.name)
            if not member_path.parts or member_path.parts[0] != ARCHIVE_ROOT:
                raise RuntimeError("LibreOffice archive has an unexpected root layout")
            target = (destination / member_path).resolve()
            try:
                target.relative_to(destination)
            except ValueError as error:
                raise RuntimeError(
                    "LibreOffice archive contains an unsafe path"
                ) from error
            if member.isdir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            if not member.isfile():
                raise RuntimeError("LibreOffice archive contains a non-file entry")
            source = archive.extractfile(member)
            if source is None:
                raise RuntimeError("LibreOffice archive member could not be read")
            target.parent.mkdir(parents=True, exist_ok=True)
            with source, target.open("wb") as output:
                shutil.copyfileobj(source, output)
            target.chmod(stat.S_IMODE(member.mode))

    source_root = destination / ARCHIVE_ROOT
    deb_root = source_root / "DEBS"
    packages = tuple(sorted(deb_root.glob("*.deb")))
    if not packages:
        raise RuntimeError("LibreOffice archive contains no Debian packages")
    required = (
        f"libobasis26.2-core_{LIBREOFFICE_VERSION}-2_amd64.deb",
        f"libreoffice26.2-impress_{LIBREOFFICE_VERSION}-2_amd64.deb",
    )
    for package in required:
        if not (deb_root / package).is_file():
            raise RuntimeError(f"LibreOffice archive is missing {package}")
    return deb_root


def verify_soffice(executable: Path = SOFFICE) -> None:
    if not executable.is_file():
        raise RuntimeError(f"missing LibreOffice executable: {executable}")
    result = subprocess.run(
        [str(executable), "--version"],
        check=True,
        capture_output=True,
        text=True,
    )
    actual = result.stdout.strip() or result.stderr.strip()
    if actual.splitlines()[:1] != [EXPECTED_IDENTITY]:
        raise RuntimeError(f"unexpected LibreOffice identity: {actual!r}")


def expose_soffice(executable: Path = SOFFICE) -> None:
    github_path = os.environ.get("GITHUB_PATH")
    if github_path:
        with Path(github_path).open("a", encoding="utf-8") as path_file:
            path_file.write(f"{executable.parent}\n")


def install() -> None:
    if platform.system() != "Linux" or platform.machine() != "x86_64":
        raise RuntimeError("the pinned LibreOffice installer requires Linux x86_64")
    if INSTALL_ROOT.exists():
        raise RuntimeError(
            f"LibreOffice prefix must be absent before installation: {INSTALL_ROOT}"
        )

    runner_temp = Path(os.environ.get("RUNNER_TEMP", tempfile.gettempdir()))
    with tempfile.TemporaryDirectory(
        prefix="rdocx-libreoffice-", dir=runner_temp
    ) as work:
        work_root = Path(work)
        archive_path = work_root / LIBREOFFICE_ARCHIVE
        download_archive(archive_path)
        deb_root = safe_extract(archive_path, work_root / "source")
        packages = tuple(str(package) for package in sorted(deb_root.glob("*.deb")))
        subprocess.run(
            [
                "sudo",
                "apt-get",
                "install",
                "--yes",
                "--no-install-recommends",
                *SYSTEM_RUNTIME_PACKAGES,
                *packages,
            ],
            check=True,
        )

    verify_soffice()
    expose_soffice()
    print(f"Installed LibreOffice {LIBREOFFICE_VERSION} from reviewed packages")


if __name__ == "__main__":
    install()
