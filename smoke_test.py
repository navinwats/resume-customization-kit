#!/usr/bin/env python3
"""
smoke_test.py — Check that the kit works end to end on a fresh clone.

Copies the kit's scripts into a temporary directory and runs each one there with
a built-in sample resume, so your real master resume, bullet bank and output are
never touched. Prints ✓/✗ per check. Exit 0 if every check passes, 1 otherwise.

Checks:
  1. build_master.py writes master_resume.docx
  2. generate_resume.py builds a .docx and .md from sample content
  3. Content validation rejects a missing "experience" and an em dash
  4. Master lookup fails clearly when there is no master
  5. trace_check.py passes a traceable figure and fails an invented one
  6. score_keywords.py reports on requirements_example.json

Usage:
    python3 smoke_test.py
    python3 smoke_test.py --pdf

By default PDF export is skipped (--no-pdf), so Word and LibreOffice never open.
--pdf also runs generate_resume.py with PDF export and --pages 1, and expects a
PDF and a page count. That may open Word or LibreOffice. Its temp folder lives
under this kit's output/ folder (not /tmp) because Word on macOS is sandboxed
and cannot save to /tmp; it is deleted afterwards.

Needs python-docx (and pypdf for --pdf), see requirements.txt. Standard library
otherwise.
"""

import os
import sys
import json
import shutil
import argparse
import tempfile
import subprocess

# ── Constants ────────────────────────────────────────────────────────────────

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
KIT_FILES = ["generate_resume.py", "build_master.py", "trace_check.py",
             "score_keywords.py", "requirements_example.json"]

SAMPLE = {
    "headline": "Operations Analyst | Process Improvement | Reporting",
    "summary": "**Operations analyst** with 6+ years improving reporting and "
               "process quality for mid-size teams.",
    "experience": [
        {
            "company": "EXAMPLE CORP",
            "title": "Senior Analyst",
            "location": "Springfield, ST",
            "dates": "Jan 2021-Present",
            "bullets": [
                "**Reporting:** Built weekly dashboards used by 12 managers, "
                "cutting review time **40%**.",
                "**Data Quality:** Fixed SQL checks that caught 300 bad records per month.",
            ],
        },
        {
            "company": "SAMPLE LLC",
            "title": "Analyst",
            "location": "Shelbyville, ST",
            "dates": "Jun 2017-Dec 2020",
            "bullets": [
                "**Process:** Documented 25 workflows and removed 3 manual handoffs.",
            ],
        },
    ],
    "skills": [
        {"header": "Analysis", "items": "SQL, Tableau, Excel"},
        {"header": "Practice", "items": "Stakeholder Management, Experimentation"},
    ],
    "education": "BS Business | Example University, Springfield, ST",
}

# Supports every figure in SAMPLE (bare years and dates are exempt)
BANK = """# Bullet bank (sample)
- Built weekly dashboards used by 12 managers, cutting review time 40%.
- Fixed SQL checks that caught 300 bad records per month.
- Documented 25 workflows and removed 3 manual handoffs.
- 6+ years of experience.
"""


# ── Helpers ───────────────────────────────────────────────────────────────────

results = []


def check(name, ok, detail=""):
    """Record and print one check result."""
    results.append(ok)
    print(f"{'✓' if ok else '✗'} {name}")
    if not ok and detail:
        for line in detail.strip().splitlines()[-8:]:
            print(f"    {line}")


def run(cwd, script, *args, env=None):
    """Run a kit script in cwd with this interpreter; return the CompletedProcess."""
    return subprocess.run(
        [sys.executable, os.path.join(cwd, script)] + list(args),
        cwd=cwd, env=env, stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        universal_newlines=True)


def write_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def make_kit(root, name):
    """Copy the kit's scripts into a fresh subfolder of root."""
    kit = os.path.join(root, name)
    os.makedirs(kit)
    for f in KIT_FILES:
        shutil.copy(os.path.join(SCRIPT_DIR, f), kit)
    return kit


def clean_env():
    """Child environment with MASTER_RESUME removed."""
    env = dict(os.environ)
    env.pop("MASTER_RESUME", None)
    return env


# ── Checks ────────────────────────────────────────────────────────────────────

def run_checks(tmp):
    env = clean_env()
    kit = make_kit(tmp, "kit")
    docx_name = "Sample Person Resume.docx"
    out = os.path.join(tmp, "out")
    content = os.path.join(kit, "sample.json")
    write_json(content, SAMPLE)

    # 1. build_master
    r = run(kit, "build_master.py", env=env)
    check("build_master.py creates master_resume.docx",
          r.returncode == 0 and os.path.isfile(os.path.join(kit, "master_resume.docx")),
          r.stdout)

    # 2. generate
    r = run(kit, "generate_resume.py", content, docx_name,
            "--outdir", out, "--md", "--no-pdf", env=env)
    docx = os.path.join(out, docx_name)
    md = os.path.splitext(docx)[0] + ".md"
    check("generate_resume.py writes .docx and .md from sample content",
          r.returncode == 0 and os.path.isfile(docx) and os.path.isfile(md),
          r.stdout)

    # 3. validation
    bad = dict(SAMPLE)
    del bad["experience"]
    bad_path = os.path.join(kit, "missing.json")
    write_json(bad_path, bad)
    out3 = os.path.join(tmp, "out3")
    r = run(kit, "generate_resume.py", bad_path, "Bad.docx",
            "--outdir", out3, "--no-pdf", env=env)
    wrote = os.path.isdir(out3) and any(f.endswith(".docx") for f in os.listdir(out3))
    check("Validation: missing \"experience\" exits 1 and writes no .docx",
          r.returncode == 1 and not wrote and "experience" in r.stdout, r.stdout)

    dash = json.loads(json.dumps(SAMPLE))
    dash["summary"] += " Led work — end to end."
    dash_path = os.path.join(kit, "dash.json")
    write_json(dash_path, dash)
    out3b = os.path.join(tmp, "out3b")
    r = run(kit, "generate_resume.py", dash_path, "Dash.docx",
            "--outdir", out3b, "--no-pdf", env=env)
    wrote = os.path.isdir(out3b) and any(f.endswith(".docx") for f in os.listdir(out3b))
    check("Validation: em dash exits 1 and writes no .docx",
          r.returncode == 1 and not wrote and "em dash" in r.stdout, r.stdout)

    # 4. master preflight
    empty = make_kit(tmp, "nomaster")
    r = run(empty, "generate_resume.py", content, "NoMaster.docx",
            "--outdir", os.path.join(tmp, "out4"), "--no-pdf", env=env)
    check("Master lookup: no master exits 1 with a lookup error",
          r.returncode == 1 and "No master resume" in r.stdout, r.stdout)

    # 5. trace_check
    bank = os.path.join(kit, "bank.md")
    with open(bank, "w", encoding="utf-8") as f:
        f.write(BANK)
    r = run(kit, "trace_check.py", content, bank, env=env)
    check("trace_check.py passes content the bank supports", r.returncode == 0, r.stdout)

    fake = json.loads(json.dumps(SAMPLE))
    fake["experience"][0]["bullets"][0] += " Saved $987,654 a year."
    fake_path = os.path.join(kit, "fake.json")
    write_json(fake_path, fake)
    r = run(kit, "trace_check.py", fake_path, bank, env=env)
    check("trace_check.py fails a figure absent from the bank",
          r.returncode == 1 and "987,654" in r.stdout, r.stdout)

    # 6. score_keywords
    r = run(kit, "score_keywords.py", os.path.join(kit, "requirements_example.json"),
            content, env=env)
    check("score_keywords.py reports on requirements_example.json",
          r.returncode == 0, r.stdout)


def run_pdf_check():
    """Check 2 again with PDF export and --pages 1, under the kit's output/ folder."""
    out_root = os.path.join(SCRIPT_DIR, "output")
    made_root = not os.path.isdir(out_root)
    os.makedirs(out_root, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="smoke_", dir=out_root)
    try:
        env = clean_env()
        kit = make_kit(tmp, "kit")
        content = os.path.join(kit, "sample.json")
        write_json(content, SAMPLE)
        r = run(kit, "build_master.py", env=env)
        if r.returncode != 0:
            check("PDF export: build_master.py", False, r.stdout)
            return
        name = "Sample Person Resume.docx"
        out = os.path.join(tmp, "out")
        r = run(kit, "generate_resume.py", content, name,
                "--outdir", out, "--pages", "1", env=env)
        pdf = os.path.splitext(os.path.join(out, name))[0] + ".pdf"
        check("PDF export: .pdf written and page count is 1",
              r.returncode == 0 and os.path.isfile(pdf) and "Page count: 1" in r.stdout,
              r.stdout)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        if made_root:
            try:
                os.rmdir(out_root)
            except OSError:
                pass


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Check that the kit works end to end, in a temporary directory.")
    parser.add_argument("--pdf", action="store_true",
                        help="also test PDF export (may open Word or LibreOffice)")
    args = parser.parse_args()

    tmp = tempfile.mkdtemp(prefix="kit_smoke_")
    try:
        run_checks(tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    if args.pdf:
        run_pdf_check()

    passed = sum(results)
    print()
    if all(results):
        print(f"✓ All {len(results)} checks passed.")
        sys.exit(0)
    print(f"✗ {len(results) - passed} of {len(results)} checks failed.")
    sys.exit(1)


if __name__ == "__main__":
    main()
