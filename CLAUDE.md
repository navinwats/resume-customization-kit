# Resume Customization Workflow

This folder generates tailored resumes from a master bullet bank. When a job description is provided, customize and generate a resume unless asked to evaluate the role first.

---

## File Structure

| File | Purpose |
|---|---|
| `master_resume_bullets.md` | Single source of truth for all bullets, organized by role, with verified metrics |
| `resume_instructions.md` | Full customization playbook: role families, headline/summary formulas, bullet selection logic, role inclusion rules |
| `tiering_instructions.md` | Optional. Job evaluation framework (use when asked to tier a role before generating). Created from `tiering_instructions_template.md` |
| `generate_resume.py` | Converts a JSON content file into a formatted `.docx` using the master resume as a style template |
| `build_master.py` | Writes a starter `master_resume.docx` with the structure the generator needs, for users without their own master |
| `trace_check.py` | Fabrication check: every number in a content JSON must appear in `master_resume_bullets.md` |
| `score_keywords.py` | Optional report: weighted job-description term coverage of a content JSON, plus the share of bullets with a figure |
| `requirements_[company].json` | Optional, per role: weighted terms extracted from the job description (see `requirements_example.json`) |
| `master_resume.docx` (or one `*MASTER*.docx`) | Style template — fonts, margins, and formatting are inherited from this file |
| `resume_content_[company].json` | Generated per role; contains headline, summary, experience, skills, education |

---

## Default Behavior

When a job description is dropped in:
- **Generate immediately** unless the user asks for a tiering evaluation first
- Always produce both a `.docx` and `.pdf`
- Output filename format: `[Full Name] Resume MM.DD.YYYY [Company] [Role Short].docx`

---

## Generation Steps

1. Read `resume_instructions.md` and `master_resume_bullets.md` fully before writing any content
2. Classify the role family per Step 0 of `resume_instructions.md` — this drives everything
3. Write `resume_content_[company].json` following the JSON schema at the bottom of `resume_instructions.md`
4. Check that every figure traces to the bank:
   ```
   python3 trace_check.py resume_content_[company].json
   ```
   Every number in the headline, summary, bullets and skills must appear in `master_resume_bullets.md` (bare years and `dates` fields are exempt). On exit 1, replace each listed figure with one the bank states. Never add a number to the bank just to pass.
5. Run the generator from this directory:
   ```
   python3 generate_resume.py resume_content_[company].json "[Full Name] Resume MM.DD.YYYY Company Role.docx"
   ```
   It validates the JSON first (required keys, every role's company/title/dates/bullets, no em dashes) and exits 1 listing every problem by JSON path before writing anything. Fix them all, then rerun.
6. Read the generator's exit status and last lines — **do not skip this step**. After saving the `.docx` it exports the PDF (Word first, LibreOffice as fallback), prints which backend produced it, then counts the PDF's pages with pypdf:
   - `✓ Page count: 2` and exit 0 — done.
   - `⚠ Page count: N (expected 2)` and exit 1 — the page count is wrong.
   - `⚠ PDF export failed` and exit 1 — no PDF was produced, so the page count is unverified. Never report success in this case; see PDF Troubleshooting.
7. If page count is not 2, make one informed cut and regenerate once. See page count rules below.

## Analyze / score (optional)

When asked to score a resume, or before generating if the user wants a keyword check:
1. Extract the job description's weighted terms (weight 1-3, with `alternates` for synonyms) into `requirements_[company].json`, in the format of `requirements_example.json`.
2. Run `python3 score_keywords.py requirements_[company].json resume_content_[company].json`. It is a report only and always exits 0 (exit 2 on bad input). It lists each term as hit or miss with where it matched, the weighted coverage %, the misses by weight, and the share of bullets with a figure against the 60-75% target.
3. A miss is measured against the page, not the bank. Search `master_resume_bullets.md` for each miss first (and its synonyms). If the bank supports it, it is a selection choice: use that bullet or skill. Only if the bank does not support it, ask the user ("The posting asks for X; I don't see it in your bank. Is there something that fits?"). Never tell the user the bank lacks something without checking, and never insert a term the bank does not support just to raise the score.

---

## Hard Rules

### Page count
- Must be exactly 2 pages. The generator checks this on every run and exits 1 otherwise — do not declare done unless it printed `✓ Page count: 2`.
- Target ~18 bullets total across all roles. This leaves rendering headroom for Word's PDF conversion.
- Resumes must fill close to 2 full pages — do not leave significant white space.
- If 3 pages: make one targeted cut (remove 1–2 bullets from least-relevant roles), regenerate, verify. Never loop through generation more than 2–3 times total.

### Dashes
- Never use em dashes (—) anywhere in resume content.
- Prefer commas, semicolons, or rephrased constructions.
- Use parentheses for appositives (e.g., `SoundSink (a B2B music licensing CRM)`).
- Use en dash (–) only when a dash is truly unavoidable.
- `generate_resume.py` rejects any `—` in the JSON, listing each location, and exits 1 before writing anything.

### Titles
- Always spell out "Senior" — never abbreviate as "Sr" or "Sr."
- This applies to all title variants (e.g., "Senior Product Marketing Manager", "Senior GTM Strategy Manager").

### Content integrity
- Never fabricate data, metrics, or outcomes. Pull only from `master_resume_bullets.md`.
- Summary must be exactly two sentences and must not open with "I" or the candidate's name.
- Bold only 1–2 anchor phrases in the summary.
- **Near-miss skills become questions.** If the job description names a skill the bank lacks but something adjacent exists, do not stop to ask first: generate from supported, verified content, leave the near-miss skill out, and do not insert it or upgrade the adjacent one. Right after delivering, ask the user whether they have real experience with it. Search the whole bank before calling anything a gap: a requirement not on the page is a selection question until the full bank has been searched. Requirements the bank cannot support go to the user in a short gap list, also right after delivering.
- **Reword vocabulary, never the claim.** Mirroring the job description's vocabulary is fine. The claim, metric, scope, verb strength, employer names and dates stay exactly as in the bank. "Led" never becomes "built", and numbers never round up.
- **Mirror the concept, not the sentences.** Do not lift a distinctive phrase from the posting whole, and do not borrow a negation such as "not just X" (it defends against a charge nobody made). State the work positively, in the bank's own words.
- **Bank status tags.** A bullet may end with an optional tag. Untagged bullets are treated as `[verified]`. `[verified]` bullets are freely usable. `[unverified]` bullets are left out of the generated resume; list them right after delivering and ask the user to confirm, and once confirmed, change the tag to `[verified]` in the bank (they can then be used on the next run). `[flagged]` bullets never ship, whatever the job description asks. Never copy a tag into the content JSON.

---

## PDF Troubleshooting

The generator tries Word first (osascript on macOS, docx2pdf on Windows) because Word renders justified text correctly, then falls back to LibreOffice on every platform. The `⚠ PDF export failed` message lists why each backend failed.

If the PDF fails to generate or Word hangs:
1. On macOS, close all open Word documents: `osascript -e 'tell application "Microsoft Word" to close every document saving no'`
2. Wait 2 seconds, then rerun the generator once.
3. Do **not** run iterative generation loops — each attempt opens and closes Microsoft Word, which is disruptive. Make a content decision first, then generate once.

If Word is not installed, install LibreOffice (`brew install --cask libreoffice` on macOS, or libreoffice.org) so the fallback can run. To write only the `.docx`, pass `--no-pdf`; the page count is then unverified.

---

## Customizing for Your Own Resume

To use this workflow with your own resume:

1. **Build your `master_resume_bullets.md`** — list every role with verified, metrics-backed bullets. This is the most important file. Do not include anything you cannot verify.
2. **Write your `resume_instructions.md`** — define your role families, headline formulas, bullet selection priorities, and role inclusion rules. See the existing file as a reference.
3. **Add your master `.docx`** — no script edits needed. If you don't have one, run `python3 build_master.py --name "..." --contact "..." --links "..."` to write a starter `master_resume.docx` with the required structure (add `--force` to overwrite an existing one). `generate_resume.py` uses the first match of: the `MASTER_RESUME` environment variable (full path), then `master_resume.docx` in this folder, then exactly one `*MASTER*.docx` in this folder (Word `~$` lock files are ignored). It stops with an error if none is found or several `*MASTER*.docx` files match. The master must define the paragraph styles `Body A`, `Section Title` and `List Paragraph` (bullets use the list attached to `List Paragraph`, else numId 2), and start with a header table whose first cell holds the name, then headline + line break + contact, then links. The script checks this before building and names anything missing.
4. **Optional: write your `tiering_instructions.md`** — copy `tiering_instructions_template.md` and fill in your hard filters, soft signals, and tier thresholds. Only used when you ask to tier a role.
5. **Install dependencies:**
   ```
   pip install -r requirements.txt
   ```
6. Open Claude Code in this folder and paste any job description to generate a resume.

---

## Quick Reference: JSON Schema

```json
{
  "headline": "One-line identity statement",
  "summary": "Two sentences. **Bold** 1-2 anchors. No em dashes.",
  "experience": [
    {
      "company": "COMPANY NAME",
      "title": "Full Title (never abbreviate Senior)",
      "location": "City, ST",
      "dates": "Mon YYYY-Mon YYYY",
      "bullets": [
        "**Bullet Header:** Bullet body with **metrics** inline."
      ]
    }
  ],
  "skills": [
    {
      "header": "Skill Group Title",
      "items": "Skill 1, Skill 2, Skill 3"
    }
  ],
  "education": "Degree | University, City, ST"
}
```
