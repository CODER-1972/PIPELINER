#!/usr/bin/env python3
"""Valida higiene de encoding dos módulos VBA.

Checks:
1) blobs versionados (.bas/.cls/.frm) precisam ser UTF-8 válidos no Git object store;
2) working tree precisa ser decodificável em Windows-1252 (compatível com VBE);
3) deteção heurística de mojibake comum em português (ex.: MÃ³dulo, nÃ£o, validaÃ§Ã£o).
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
VBA_GLOBS = ("*.bas", "*.cls", "*.frm")

MOJIBAKE_PATTERNS = [
    re.compile(r"Ã[\x80-\xBF]"),
    re.compile(r"Â[\x80-\xBF]"),
    re.compile(r"â[\x80-\xBF]{1,2}"),
    re.compile(r"ï»¿"),
]

# Exceções legítimas para textos com caracteres reais (não mojibake)
SAFE_SUBSTRINGS = {
    '"Á", "À", "Ã", "Â", "Ä"',
    '"á", "à", "ã", "â", "ä"',
}


def iter_vba_files() -> list[Path]:
    files: list[Path] = []
    src = ROOT / "src" / "vba"
    for pattern in VBA_GLOBS:
        files.extend(sorted(src.glob(pattern)))
    return files


def check_git_blob_utf8(path: Path, errors: list[str]) -> None:
    rel = path.relative_to(ROOT).as_posix()
    try:
        blob = subprocess.check_output(["git", "show", f"HEAD:{rel}"], cwd=ROOT)
    except subprocess.CalledProcessError as exc:
        errors.append(f"[blob] não foi possível ler HEAD:{rel} ({exc})")
        return

    try:
        blob.decode("utf-8")
    except UnicodeDecodeError as exc:
        errors.append(f"[blob] {rel}: conteúdo versionado não é UTF-8 válido (offset {exc.start})")


def looks_like_mojibake(line: str) -> bool:
    if any(s in line for s in SAFE_SUBSTRINGS):
        return False
    return any(p.search(line) for p in MOJIBAKE_PATTERNS)


def check_worktree_cp1252(path: Path, errors: list[str]) -> None:
    rel = path.relative_to(ROOT).as_posix()
    data = path.read_bytes()

    try:
        text = data.decode("cp1252")
    except UnicodeDecodeError as exc:
        errors.append(f"[worktree] {rel}: não decodifica em cp1252 (offset {exc.start})")
        return

    # newline policy conforme .gitattributes
    if b"\n" in data.replace(b"\r\n", b""):
        errors.append(f"[worktree] {rel}: contém LF sem CRLF (esperado eol=crlf)")

    for idx, line in enumerate(text.splitlines(), start=1):
        if looks_like_mojibake(line):
            errors.append(f"[mojibake] {rel}:{idx}: {line.strip()[:180]}")


def main() -> int:
    files = iter_vba_files()
    if not files:
        print("Nenhum módulo VBA encontrado em src/vba.")
        return 0

    errors: list[str] = []
    for path in files:
        check_git_blob_utf8(path, errors)
        check_worktree_cp1252(path, errors)

    if errors:
        print("Falha na validação de encoding VBA:")
        for err in errors:
            print(f" - {err}")
        return 1

    print(f"OK: {len(files)} módulos VBA validados (blob UTF-8 + worktree cp1252 + heurística mojibake).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
