#!/usr/bin/env python3
"""
score_keywords.py — Report how well a content JSON covers a job's weighted terms.

Report only: it never blocks a build. Use it to see which job-description terms
the resume already shows, where, and which it misses, then ask the user about
the misses. Never add a term the bullet bank does not support.

requirements.json:
{
  "title": "Senior Analyst",
  "terms": [
    {"term": "data quality", "weight": 3, "alternates": ["data integrity"]},
    {"term": "SQL", "weight": 2}
  ]
}
weight is 1-3 (3 = must-have); alternates is optional.

Matching is on word boundaries. Between words, whitespace or a hyphen matches
either way ("data quality" finds "data-quality"). The last word matches its
singular or plural both ways ("dashboards" finds "dashboard" and vice versa);
only a trailing s / es is handled, so irregular plurals ("category" /
"categories") are not. Matching is case-insensitive, except that a term of 3
characters or fewer containing an uppercase letter (Go, R, AI, SQL, AWS) is
case-sensitive and takes no plural, so "Go" does not hit "go-to-market".
A term with no weight defaults to 1 and prints a warning.

Reports, per term: hit or miss, which alternate matched, and where it was found
(headline, summary, experience[i].company / .title / .bullets, skills, education). Then a weighted coverage %,
the misses sorted by weight, and the share of bullets carrying a real figure
against the 60-75% quantification target in resume_instructions_template.md.

Usage:
    python3 score_keywords.py requirements_example.json resume_content_[company].json

Exit 0 always, except exit 2 on bad input. Standard library only.
"""

import sys
import re
import json
import argparse

# ── Constants ────────────────────────────────────────────────────────────────

NUM_RE = re.compile(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?")
YEAR_RE = re.compile(r"(19|20)\d\d")

# resume_instructions_template.md: "60–75% of bullets should include a metric,
# benchmark, or scale marker."
TARGET_LOW, TARGET_HIGH = 60, 75


class BadInput(Exception):
    pass


# ── Matching ──────────────────────────────────────────────────────────────────

def plural_forms(word):
    """
    The word plus its simple singular or plural: "dashboards" <-> "dashboard",
    "process" <-> "processes". Only a trailing s / es is handled; irregular
    plurals ("analysis"/"analyses", "category"/"categories") are not.
    """
    low = word.lower()
    forms = {word}
    if re.search(r"(s|x|z|ch|sh)es$", low) and len(low) > 4:
        forms.add(word[:-2])
    if low.endswith("s") and not low.endswith("ss"):
        if len(low) > 3:
            forms.add(word[:-1])
    else:
        forms.add(word + "s")
        if re.search(r"(s|x|z|ch|sh)$", low):
            forms.add(word + "es")
    return sorted(forms, key=len, reverse=True)


def term_regex(term):
    """
    Word-boundary match; whitespace or hyphen between words. The last word also
    matches its simple singular or plural (see plural_forms). A term of 3
    characters or fewer that has an uppercase letter (Go, R, AI, SQL, AWS) is
    case-sensitive and takes no plural; everything else is case-insensitive.
    """
    term = term.strip()
    words = [w for w in re.split(r"[\s-]+", term) if w]
    if not words:
        raise BadInput(f"empty term: {term!r}")
    exact = len(term) <= 3 and any(c.isupper() for c in term)
    escaped = [re.escape(w) for w in words]
    if not exact and words[-1][-1].isalpha():
        escaped[-1] = "(?:" + "|".join(re.escape(f) for f in plural_forms(words[-1])) + ")"
    return re.compile(r"(?<![A-Za-z0-9])" + r"[\s-]+".join(escaped) + r"(?![A-Za-z0-9])",
                      0 if exact else re.IGNORECASE)


def carries_figure(bullet):
    """True if the text holds a numeral that is not a bare year."""
    for m in NUM_RE.finditer(bullet.replace("**", "")):
        tok = m.group(0)
        if "," in tok or not YEAR_RE.fullmatch(tok):
            return True
    return False


# ── Loading ───────────────────────────────────────────────────────────────────

def load_json(path, what):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError) as e:
        raise BadInput(f"cannot read {what} {path}: {e}")


def load_terms(req):
    if not isinstance(req, dict) or not isinstance(req.get("terms"), list) or not req["terms"]:
        raise BadInput('requirements needs a non-empty "terms" list')
    terms = []
    seen = set()
    for i, t in enumerate(req["terms"]):
        if not isinstance(t, dict) or not isinstance(t.get("term"), str) or not t["term"].strip():
            raise BadInput(f'terms[{i}] needs a "term" string')
        if "weight" not in t:
            print(f'⚠ Term "{t["term"].strip()}" (terms[{i}]) has no weight; using 1.')
        weight = t.get("weight", 1)
        if isinstance(weight, bool) or not isinstance(weight, int) or not 1 <= weight <= 3:
            raise BadInput(f"terms[{i}] weight must be an integer 1-3")
        alts = t.get("alternates", [])
        if not isinstance(alts, list) or not all(isinstance(a, str) and a.strip() for a in alts):
            raise BadInput(f"terms[{i}] alternates must be a list of strings")
        key = t["term"].strip().lower()
        if key in seen:
            print(f'⚠ Duplicate term "{t["term"].strip()}" (terms[{i}]) ignored; counted once.')
            continue
        seen.add(key)
        terms.append({"term": t["term"].strip(), "weight": weight, "alternates": alts})
    return terms


def sections(content):
    """List of (label, text) for every searchable part of the content JSON."""
    if not isinstance(content, dict):
        raise BadInput("content must be a JSON object")
    out = []
    if isinstance(content.get("headline"), str):
        out.append(("headline", content["headline"]))
    if isinstance(content.get("summary"), str):
        out.append(("summary", content["summary"]))
    experience = content.get("experience") or []
    if not isinstance(experience, list):
        raise BadInput('content "experience" must be a list')
    for i, role in enumerate(experience):
        if isinstance(role, dict):
            for key in ("company", "title"):
                if isinstance(role.get(key), str):
                    out.append((f"experience[{i}].{key}", role[key]))
            bullets = [b for b in role.get("bullets") or [] if isinstance(b, str)]
            out.append((f"experience[{i}].bullets", "\n".join(bullets)))
    skills = []
    for s in content.get("skills") or []:
        if isinstance(s, dict):
            skills += [s[k] for k in ("header", "items") if isinstance(s.get(k), str)]
    if skills:
        out.append(("skills", "\n".join(skills)))
    if isinstance(content.get("education"), str):
        out.append(("education", content["education"]))
    return [(label, text.replace("**", "")) for label, text in out]


def all_bullets(content):
    return [b for role in content.get("experience") or [] if isinstance(role, dict)
            for b in role.get("bullets") or [] if isinstance(b, str)]


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Report weighted keyword coverage of a resume content JSON (report only).")
    parser.add_argument("requirements", help="requirements JSON: title + weighted terms")
    parser.add_argument("content", help="resume content JSON file")
    args = parser.parse_args()

    try:
        req = load_json(args.requirements, "requirements")
        terms = load_terms(req)
        content = load_json(args.content, "content")
        parts = sections(content)
        results = []
        for t in terms:
            matched, where = None, []
            for name in [t["term"]] + t["alternates"]:
                rx = term_regex(name)
                hit = [label for label, text in parts if rx.search(text)]
                if hit:
                    matched = name
                    where = hit
                    break
            results.append((t, matched, where))
    except BadInput as e:
        print(f"ERROR: {e}")
        sys.exit(2)

    title = req.get("title") if isinstance(req.get("title"), str) else ""
    print(f"Keyword coverage{': ' + title if title else ''}")
    print()
    for t, matched, where in results:
        if matched is None:
            print(f"  ⚠ MISS  [w{t['weight']}] {t['term']}")
        else:
            via = "" if matched == t["term"] else f' (via alternate "{matched}")'
            print(f"  ✓ HIT   [w{t['weight']}] {t['term']}{via}: {', '.join(where)}")

    total = sum(t["weight"] for t, _, _ in results)
    got = sum(t["weight"] for t, m, _ in results if m is not None)
    n_hit = sum(1 for _, m, _ in results if m is not None)
    print()
    print(f"Weighted coverage: {got}/{total} = {100 * got / total:.0f}% "
          f"({n_hit}/{len(results)} terms)")

    misses = sorted((t for t, m, _ in results if m is None), key=lambda t: -t["weight"])
    if misses:
        print("Misses, highest weight first (check the bank first; ask the user only if it has no support; never add unsupported terms):")
        for t in misses:
            print(f"  - [w{t['weight']}] {t['term']}")
    else:
        print("No misses.")

    bullets = all_bullets(content)
    print()
    if bullets:
        n = sum(1 for b in bullets if carries_figure(b))
        pct = 100 * n / len(bullets)
        if pct < TARGET_LOW:
            note = f"below the {TARGET_LOW}-{TARGET_HIGH}% target"
        elif pct > TARGET_HIGH:
            note = f"above the {TARGET_LOW}-{TARGET_HIGH}% target"
        else:
            note = f"within the {TARGET_LOW}-{TARGET_HIGH}% target"
        print(f"Bullets with a figure: {n}/{len(bullets)} = {pct:.0f}% ({note})")
    else:
        print("Bullets with a figure: no bullets found")


if __name__ == "__main__":
    main()
