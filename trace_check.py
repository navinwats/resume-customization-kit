#!/usr/bin/env python3
"""
trace_check.py — Fabrication check: every number on the resume must come from the bank.

A made-up or rounded claim almost always introduces a figure that the bullet bank
never states. This script collects every numeral from the content JSON's
headline, summary, bullets, skills and education, and checks that each one appears somewhere
in master_resume_bullets.md (anywhere: bullets, notes, or the fixed header).

Exempt:
  - Bare years (19xx / 20xx), wherever they appear
  - Everything in "dates" fields

Usage:
    python3 trace_check.py resume_content_[company].json
    python3 trace_check.py resume_content_[company].json path/to/master_resume_bullets.md

Exit 0 = every figure is traceable. Exit 1 = at least one is not (each is listed
with its JSON path). Standard library only.
"""

import sys
import os
import re
import json
import argparse

# ── Constants ────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_BANK = os.path.join(SCRIPT_DIR, "master_resume_bullets.md")

NUM_RE = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?")
YEAR_RE = re.compile(r"(19|20)\d\d")


# ── Tokenizing ────────────────────────────────────────────────────────────────

def numerals(text):
    """
    Numeric tokens in text, normalized so formatting differences still match:
    thousands commas dropped ("1,200" -> "1200"), leading zeros dropped,
    trailing decimal zeros dropped ("1.50" -> "1.5", "3.0" -> "3"), and
    symbols ignored ("$2.5M", "25%", "10+" -> "2.5", "25", "10").
    Bare years are skipped; a comma-formatted figure ("2,000") is never a year.
    """
    out = set()
    for m in NUM_RE.finditer(text):
        tok = m.group(0)
        if "," not in tok and YEAR_RE.fullmatch(tok):
            continue
        tok = tok.replace(",", "")
        if "." in tok:
            tok = tok.rstrip("0").rstrip(".")
        out.add(tok.lstrip("0") or "0")
    return out


def content_fields(content):
    """Yield (json_path, text) for every field whose numbers must be traceable."""
    if isinstance(content.get("headline"), str):
        yield "$.headline", content["headline"]
    if isinstance(content.get("summary"), str):
        yield "$.summary", content["summary"]
    if isinstance(content.get("education"), str):
        yield "$.education", content["education"]
    for i, role in enumerate(content.get("experience") or []):
        if not isinstance(role, dict):
            continue
        for j, bullet in enumerate(role.get("bullets") or []):
            if isinstance(bullet, str):
                yield f"$.experience[{i}].bullets[{j}]", bullet
    for i, skill in enumerate(content.get("skills") or []):
        if not isinstance(skill, dict):
            continue
        for key in ("header", "items"):
            if isinstance(skill.get(key), str):
                yield f"$.skills[{i}].{key}", skill[key]


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Check that every number in a resume content JSON appears in the bullet bank.")
    parser.add_argument("content", help="resume content JSON file")
    parser.add_argument("bank", nargs="?", default=DEFAULT_BANK,
                        help="bullet bank (default: master_resume_bullets.md next to this script)")
    args = parser.parse_args()

    if not os.path.isfile(args.bank):
        print(f"ERROR: bullet bank not found: {args.bank}")
        sys.exit(1)

    with open(args.bank, "r", encoding="utf-8") as f:
        bank_nums = numerals(f.read())
    with open(args.content, "r", encoding="utf-8") as f:
        content = json.load(f)

    untraceable = []
    checked = 0
    for path, text in content_fields(content):
        for num in sorted(numerals(text), key=lambda n: float(n)):
            checked += 1
            if num not in bank_nums:
                untraceable.append((path, num, text))

    if untraceable:
        print(f"⚠ {len(untraceable)} figure(s) not found in {os.path.basename(args.bank)}:")
        for path, num, text in untraceable:
            print(f"  - {num} at {path}: {text}")
        print("  → Use a figure the bank states, or add the verified figure to the bank first.")
        sys.exit(1)

    print(f"✓ All {checked} figure(s) trace to {os.path.basename(args.bank)}")


if __name__ == "__main__":
    main()
