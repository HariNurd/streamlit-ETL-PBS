#!/usr/bin/env python3
"""Fail when repository content looks unsafe for a public Git repository."""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path, PurePosixPath


ROOT = Path(__file__).resolve().parents[1]

ALLOWED_SENSITIVE_SUFFIX_PATHS = {
    "Streamlit Project Structure.txt",
    "requirements.txt",
}

FORBIDDEN_SUFFIXES = (
    ".7z",
    ".accdb",
    ".bak",
    ".bmp",
    ".csv",
    ".db",
    ".doc",
    ".docx",
    ".duckdb",
    ".gz",
    ".heic",
    ".joblib",
    ".jpeg",
    ".jpg",
    ".jsonl",
    ".log",
    ".mdb",
    ".mt940",
    ".ndjson",
    ".ods",
    ".odt",
    ".ofx",
    ".orig",
    ".p12",
    ".parquet",
    ".pdf",
    ".pem",
    ".pfx",
    ".pickle",
    ".pkl",
    ".png",
    ".pyc",
    ".qif",
    ".rar",
    ".rtf",
    ".sqlite",
    ".sqlite3",
    ".tar",
    ".tgz",
    ".tif",
    ".tiff",
    ".tmp",
    ".tsv",
    ".txt",
    ".xls",
    ".xlsb",
    ".xlsm",
    ".xlsx",
    ".zip",
)

SENSITIVE_DIRECTORIES = {
    "cache",
    "input_train",
    "outputs",
    "temp",
    "train",
    "uploads",
}

SECRET_PATTERNS = (
    (
        "private key",
        re.compile(
            r"-----BEGIN (?:RSA |EC |OPENSSH |DSA |PGP )?PRIVATE KEY(?: BLOCK)?-----"
        ),
    ),
    ("AWS access key", re.compile(r"\bAKIA[0-9A-Z]{16}\b")),
    ("GitHub token", re.compile(r"\b(?:github_pat_|gh[pousr]_)[A-Za-z0-9_]{20,}\b")),
    ("OpenAI-style API key", re.compile(r"\bsk-[A-Za-z0-9_-]{20,}\b")),
    ("Slack token", re.compile(r"\bxox[baprs]-[A-Za-z0-9-]{10,}\b")),
    ("Google API key", re.compile(r"\bAIza[0-9A-Za-z_-]{30,}\b")),
    (
        "credential embedded in URL",
        re.compile(r"\b[a-z][a-z0-9+.-]*://[^/\s:@]+:[^/\s@]+@", re.IGNORECASE),
    ),
    (
        "hardcoded credential assignment",
        re.compile(
            r"""(?ix)
            \b(?:api[_-]?key|client[_-]?secret|access[_-]?key|password|passwd|pwd)
            \s*[:=]\s*
            ["'](?!example|placeholder|dummy|test|changeme|your[_-])[^"']{8,}["']
            """
        ),
    ),
)

PRIVACY_PATTERNS = (
    (
        "email address",
        re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.IGNORECASE),
    ),
    (
        "absolute user-home path",
        re.compile(
            r"""(?ix)
            (?:[A-Z]:[\\/](?:Users|Documents[ ]and[ ]Settings)[\\/][^\\/\s]+)
            |(?:/(?:home|Users)/[^/\s]+)
            """
        ),
    ),
    (
        "long numeric identifier",
        re.compile(r"(?<!\d)\d{10,19}(?!\d)"),
    ),
    (
        "Indonesian mobile number",
        re.compile(r"(?<!\d)(?:\+?62|0)8[1-9](?:[\s.-]?\d){7,11}(?!\d)"),
    ),
    (
        "formatted NPWP-like identifier",
        re.compile(
            r"(?<!\d)\d{2}[.\s-]?\d{3}[.\s-]?\d{3}[.\s-]?\d"
            r"[-\s]?\d{3}[.\s-]?\d{3}(?!\d)"
        ),
    ),
)

LABELED_PII_PATTERN = re.compile(
    r"""(?imx)
    ^\s*
    (?P<label>
        nama[ ](?:nasabah|debitur)
        |alamat
        |nik
        |npwp
        |no\.\s*rekening
    )
    \s*[:=]\s*(?P<value>\S.*)$
    """
)

SAFE_LABELED_VALUE_PATTERN = re.compile(
    r"""(?ix)
    ^(?:
        <[^>]+>
        |\{[^}]+\}
        |\d{1,5}
        |.*\b(?:synthetic|example|contoh|dummy|placeholder)\b.*
    )$
    """
)

SENSITIVE_COMMIT_SUBJECT_PATTERN = re.compile(
    r"\b(?:case|client|customer|debitur|identitas|input_train|nasabah)\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class Finding:
    location: str
    category: str
    line: int | None = None

    def render(self) -> str:
        suffix = f":{self.line}" if self.line is not None else ""
        return f"{self.location}{suffix}: {self.category}"


def run_git(*args: str, text: bool = False) -> bytes | str:
    return subprocess.check_output(
        ["git", *args],
        cwd=ROOT,
        stderr=subprocess.DEVNULL,
        text=text,
    )


def repository_paths() -> list[str]:
    raw = run_git("ls-files", "--cached", "--others", "--exclude-standard", "-z")
    assert isinstance(raw, bytes)
    return [
        item.decode("utf-8", errors="surrogateescape")
        for item in raw.split(b"\0")
        if item
    ]


def path_findings(path_text: str, location: str) -> list[Finding]:
    path = PurePosixPath(path_text.replace("\\", "/"))
    lowered = path.as_posix().lower()
    findings: list[Finding] = []

    if lowered not in {item.lower() for item in ALLOWED_SENSITIVE_SUFFIX_PATHS}:
        if lowered.endswith(FORBIDDEN_SUFFIXES):
            findings.append(Finding(location, "tracked data/binary file type"))

    if path.name != ".gitkeep" and any(
        part.lower() in SENSITIVE_DIRECTORIES for part in path.parts
    ):
        findings.append(Finding(location, "file inside a local data/output directory"))

    if re.search(r"(?<!\d)\d{8,19}(?!\d)", path.name):
        findings.append(Finding(location, "identifier-like number in filename"))

    basename = path.name.lower()
    if (
        basename == ".env"
        or (basename.startswith(".env.") and basename != ".env.example")
        or basename in {"secrets.toml", "credentials.json"}
        or basename.startswith(("credentials-", "client_secret", "service-account"))
    ):
        findings.append(Finding(location, "credential/secrets filename"))

    return findings


def line_number(text: str, index: int) -> int:
    return text.count("\n", 0, index) + 1


def content_findings(data: bytes, location: str) -> list[Finding]:
    if b"\0" in data:
        return [Finding(location, "tracked binary content")]

    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError:
        try:
            text = data.decode("cp1252")
        except UnicodeDecodeError:
            return [Finding(location, "non-text content")]

    findings: list[Finding] = []
    for category, pattern in (*SECRET_PATTERNS, *PRIVACY_PATTERNS):
        for match in pattern.finditer(text):
            findings.append(
                Finding(location, category, line_number(text, match.start()))
            )

    if not location.lower().endswith(".py"):
        for match in LABELED_PII_PATTERN.finditer(text):
            value = match.group("value").strip()
            if not SAFE_LABELED_VALUE_PATTERN.fullmatch(value):
                findings.append(
                    Finding(
                        location,
                        f"literal value after {match.group('label').strip()} label",
                        line_number(text, match.start()),
                    )
                )

    return findings


def scan_worktree() -> list[Finding]:
    findings: list[Finding] = []
    for path_text in repository_paths():
        location = path_text.replace("\\", "/")
        findings.extend(path_findings(location, location))
        path = ROOT / Path(path_text)
        if path.is_file():
            findings.extend(content_findings(path.read_bytes(), location))
    return findings


def history_blobs() -> list[tuple[str, str]]:
    raw = run_git("rev-list", "--objects", "--branches", "--tags", text=True)
    assert isinstance(raw, str)
    blobs: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for line in raw.splitlines():
        object_id, separator, path_text = line.partition(" ")
        if not separator or not path_text:
            continue
        try:
            object_type = run_git("cat-file", "-t", object_id, text=True)
        except subprocess.CalledProcessError:
            continue
        assert isinstance(object_type, str)
        item = (object_id, path_text)
        if object_type.strip() == "blob" and item not in seen:
            seen.add(item)
            blobs.append(item)
    return blobs


def scan_history() -> list[Finding]:
    findings: list[Finding] = []
    for object_id, path_text in history_blobs():
        location = f"history {object_id[:12]}:{path_text}"
        findings.extend(path_findings(path_text, location))
        data = run_git("cat-file", "blob", object_id)
        assert isinstance(data, bytes)
        findings.extend(content_findings(data, location))
    findings.extend(history_metadata_findings())
    return findings


def history_metadata_findings() -> list[Finding]:
    raw = run_git(
        "log",
        "--branches",
        "--tags",
        "--format=%H%x1f%ae%x1f%ce%x1f%s%x1e",
    )
    assert isinstance(raw, bytes)
    findings: list[Finding] = []
    for record in raw.split(b"\x1e"):
        record = record.strip(b"\r\n")
        if not record:
            continue
        parts = record.decode("utf-8", errors="replace").split("\x1f", 3)
        if len(parts) != 4:
            continue
        commit_id, author_email, committer_email, subject = parts
        location = f"history commit {commit_id[:12]}"
        for role, email in (
            ("author", author_email),
            ("committer", committer_email),
        ):
            normalized = email.strip().lower()
            if normalized and not (
                normalized.endswith("@" + "users.noreply.github.com")
                or normalized
                in {
                    "noreply" + "@" + "github.com",
                    "noreply" + "@" + "localhost",
                }
            ):
                findings.append(
                    Finding(location, f"personally identifying {role} email metadata")
                )
        if SENSITIVE_COMMIT_SUBJECT_PATTERN.search(subject):
            findings.append(Finding(location, "potentially case-specific commit subject"))
    return findings


def unique_findings(findings: list[Finding]) -> list[Finding]:
    return sorted(set(findings), key=lambda item: (item.location, item.line or 0, item.category))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check tracked/unignored repository content for publication risks."
    )
    parser.add_argument(
        "--history",
        action="store_true",
        help="also scan blobs reachable from local branches and tags",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        findings = scan_worktree()
        if args.history:
            findings.extend(scan_history())
    except (OSError, subprocess.CalledProcessError) as exc:
        print(f"Audit could not run: {exc}", file=sys.stderr)
        return 2

    findings = unique_findings(findings)
    if findings:
        print("Publication-safety audit failed:")
        for finding in findings:
            print(f"- {finding.render()}")
        return 1

    scope = "working tree and branch/tag history" if args.history else "working tree"
    print(f"Publication-safety audit passed for the {scope}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
