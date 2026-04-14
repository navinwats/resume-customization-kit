# Resume Customization Workflow

This folder generates tailored resumes from a master bullet bank. When a job description is provided, customize and generate a resume unless asked to evaluate the role first.

---

## File Structure

| File | Purpose |
|---|---|
| `master_resume_bullets.md` | Single source of truth for all bullets, organized by role, with verified metrics |
| `resume_instructions.md` | Full customization playbook: role families, headline/summary formulas, bullet selection logic, role inclusion rules |
| `tiering_instructions.md` | Job evaluation framework (use when asked to tier a role before generating) |
| `generate_resume.py` | Converts a JSON content file into a formatted `.docx` using the master resume as a style template |
| `[Name] MASTER Resume [date].docx` | Style template — fonts, margins, and formatting are inherited from this file |
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
4. Run the generator from this directory:
   ```
   python3 generate_resume.py resume_content_[company].json "[Full Name] Resume MM.DD.YYYY Company Role.docx"
   ```
5. Verify page count — **do not skip this step**:
   ```
   sleep 8 && mdls -name kMDItemNumberOfPages "/full/path/to/output.pdf"
   ```
6. If page count is not 2, make one informed cut and regenerate once. See page count rules below.

---

## Hard Rules

### Page count
- Must be exactly 2 pages. Verify with `mdls` every time — do not declare done without confirming.
- Target ~18 bullets total across all roles. This leaves rendering headroom for Word's PDF conversion.
- Resumes must fill close to 2 full pages — do not leave significant white space.
- If 3 pages: make one targeted cut (remove 1–2 bullets from least-relevant roles), regenerate, verify. Never loop through generation more than 2–3 times total.

### Dashes
- Never use em dashes (—) anywhere in resume content.
- Prefer commas, semicolons, or rephrased constructions.
- Use parentheses for appositives (e.g., `SoundSink (a B2B music licensing CRM)`).
- Use en dash (–) only when a dash is truly unavoidable.
- Scan the JSON for `—` before running the generator.

### Titles
- Always spell out "Senior" — never abbreviate as "Sr" or "Sr."
- This applies to all title variants (e.g., "Senior Product Marketing Manager", "Senior GTM Strategy Manager").

### Content integrity
- Never fabricate data, metrics, or outcomes. Pull only from `master_resume_bullets.md`.
- Summary must be exactly two sentences and must not open with "I" or the candidate's name.
- Bold only 1–2 anchor phrases in the summary.

---

## PDF Troubleshooting

If the PDF fails to generate or Word hangs:
1. Close all open Word documents: `osascript -e 'tell application "Microsoft Word" to close every document saving no'`
2. Wait 2 seconds, then retry the conversion directly:
   ```python
   from docx2pdf import convert
   convert("input.docx", "output.pdf")
   ```
3. Do **not** run iterative generation loops — each attempt opens and closes Microsoft Word, which is disruptive. Make a content decision first, then generate once.

LibreOffice is a more reliable PDF backend if available: `brew install --cask libreoffice`

---

## Customizing for Your Own Resume

To use this workflow with your own resume:

1. **Build your `master_resume_bullets.md`** — list every role with verified, metrics-backed bullets. This is the most important file. Do not include anything you cannot verify.
2. **Write your `resume_instructions.md`** — define your role families, headline formulas, bullet selection priorities, and role inclusion rules. See the existing file as a reference.
3. **Replace the MASTER `.docx`** — update `generate_resume.py` to reference your own master resume file as the style template (search for `MASTER` in the script).
4. **Install dependencies:**
   ```
   pip install python-docx docx2pdf
   ```
5. Open Claude Code in this folder and paste any job description to generate a resume.

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
