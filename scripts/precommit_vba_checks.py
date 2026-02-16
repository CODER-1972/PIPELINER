#!/usr/bin/env python3
"""
Checks de pré-commit para módulos VBA (.bas/.cls/.frm).

Validações (fail-fast):
1) Encoding cp1252 (compatível com VBE/ANSI no working tree).
2) Uso proibido de escape C-style em strings VBA: \"texto\"

Uso:
- scripts/precommit_vba_checks.py            # usa ficheiros staged por default
- scripts/precommit_vba_checks.py --all      # valida todos os módulos VBA no repo
"""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
from pathlib import Path

ALLOWED_EXTS = {".bas", ".cls", ".frm"}
SUSPICIOUS_C_ESCAPE_RE = re.compile(r'\\"[A-Za-z0-9_<({$]')


class CheckError(Exception):
    pass


def run_git(args: list[str]) -> str:
    proc = subprocess.run(["git", *args], capture_output=True, text=True)
    if proc.returncode != 0:
        raise CheckError(proc.stderr.strip() or proc.stdout.strip() or f"git {' '.join(args)} falhou")
    return proc.stdout


def get_staged_vba_files() -> list[Path]:
    out = run_git(["diff", "--cached", "--name-only", "--diff-filter=ACMR"])
    files = []
    for line in out.splitlines():
        p = Path(line.strip())
        if p.suffix.lower() in ALLOWED_EXTS:
            files.append(p)
    return files


def get_all_vba_files() -> list[Path]:
    out = run_git(["ls-files", "*.bas", "*.cls", "*.frm"])
    return [Path(line.strip()) for line in out.splitlines() if line.strip()]


def line_number_from_offset(text: str, offset: int) -> int:
    return text.count("\n", 0, offset) + 1


def check_file(path: Path) -> list[str]:
    issues: list[str] = []
    if not path.exists():
        return [f"{path}: ficheiro staged não existe no working tree."]

    raw = path.read_bytes()
    try:
        text = raw.decode("cp1252")
    except UnicodeDecodeError as exc:
        return [
            f"{path}: encoding inválido para cp1252 (ANSI/VBE). "
            f"Erro: {exc}."
        ]

    # Deteta escapes C-style (ex.: Environ(\"OPENAI_API_KEY\")) e evita
    # falsos positivos comuns de literais com caminho ("\\").
    for m in SUSPICIOUS_C_ESCAPE_RE.finditer(text):
        line = line_number_from_offset(text, m.start())
        issues.append(
            f"{path}:{line}: escape C-style suspeito (\\\"). "
            'Use aspas VBA duplicadas ("") ou Chr$(34).'
        )

    return issues


def main() -> int:
    parser = argparse.ArgumentParser(description="Checks de pré-commit para VBA")
    parser.add_argument("--all", action="store_true", help="Valida todos os módulos VBA versionados")
    args = parser.parse_args()

    try:
        files = get_all_vba_files() if args.all else get_staged_vba_files()
    except CheckError as exc:
        print(f"❌ Erro ao listar ficheiros via git: {exc}")
        return 2

    if not files:
        print("✅ Sem módulos VBA para validar.")
        return 0

    all_issues: list[str] = []
    for f in files:
        all_issues.extend(check_file(f))

    if all_issues:
        print("❌ Falha nos checks de pré-commit VBA:")
        for issue in all_issues:
            print(f"  - {issue}")
        print("\nDica: confirme .gitattributes e corrija as strings VBA antes de commitar.")
        return 1

    print(f"✅ Checks VBA OK ({len(files)} ficheiro(s) validado(s)).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
