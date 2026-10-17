# Resume Customization Kit

A Claude Code workflow that generates tailored, two-page resumes from a master bullet bank. Paste in a job description, get back a formatted `.docx` and `.pdf` — in under a minute.

Built and used by a working product marketer. Every resume is grounded in real, verified bullet points pulled from your actual career. No hallucinations, no filler.

---

## How it works

1. You build a **master bullet bank** (`master_resume_bullets.md`) — every role, every metric, every proof point from your career.
2. You write **customization rules** (`resume_instructions.md`) — how to classify a role, which bullets lead, what to include or cut.
3. You open Claude Code in this folder and **paste a job description**.
4. Claude classifies the role, selects the right bullets, writes a `resume_content.json`, runs the generator, and confirms the output is exactly 2 pages.

---

## Setup (do this once)

### 1. Install dependencies

You need Python 3.9+ and the packages in `requirements.txt`:

```bash
pip install -r requirements.txt
```

That installs `python-docx` (builds the `.docx`) and `pypdf` (counts PDF pages), plus `docx2pdf` on Windows only, where it drives Word.

For PDF export you need **Microsoft Word** (Mac or Windows) or [LibreOffice](https://www.libreoffice.org/download/) (any platform). The script tries Word first because it renders justified text correctly, and falls back to LibreOffice:

```bash
brew install --cask libreoffice   # macOS; on Windows/Linux use the installer from libreoffice.org
```

### 2. Add your master resume `.docx`

**No master `.docx`?** Build a starter one that already has the required structure:

```bash
python3 build_master.py --name "Jane Smith" --contact "555-555-5555 • jane@example.com • City, ST" --links "linkedin.com/in/janesmith"
```

This writes `master_resume.docx` (Letter, 0.8" margins, Garamond 11pt). Every argument is optional; placeholders fill anything you leave out, and you can edit the text or styles in Word afterwards. It won't overwrite an existing `master_resume.docx` unless you pass `--force`.

**Have your own?** Place your existing resume `.docx` in this folder. The generator uses the first match of:
1. The `MASTER_RESUME` environment variable, set to the full path of your `.docx`
2. `master_resume.docx` in this folder (simplest)
3. Exactly one file with `MASTER` in its name (e.g., `Jane Smith MASTER Resume.docx`). If more than one matches, the script stops and lists them; Word lock files (`~$...`) are ignored.

The generator uses this file purely for its formatting — fonts, margins, styles. Your content comes from the JSON. It does need a specific structure, and checks for it before building:
- Paragraph styles named `Body A` (body text), `Section Title` (section headings) and `List Paragraph` (bullets)
- A header table as the first table in the document. Its first cell holds the name, then a paragraph with your headline, a line break (Shift+Enter in Word) and your contact line, then your links. Only the headline is replaced on each resume.

The dates tab stop is placed at the right margin of your master's page setup, so any paper size or margins work.

### 3. Fill in the template files

Copy and rename the two template files:

```bash
cp resume_instructions_template.md resume_instructions.md
cp master_resume_bullets_template.md master_resume_bullets.md
```

Then edit both files to reflect your own career, rules, and preferences. See the templates for guidance on what each section should contain.

Optionally, add a job evaluation framework so Claude can tier a role (apply, skip, or verify first) before generating anything:

```bash
cp tiering_instructions_template.md tiering_instructions.md
```

Skip this if you want every job description to go straight to a resume.

### 4. Install Claude Code

If you haven't already: [claude.ai/code](https://claude.ai/code)

Open Claude Code in this folder:

```bash
cd /path/to/this/folder
claude
```

---

## Usage

Open Claude Code in this folder and paste any job description. That's it.

Claude will:
- Read your master bullets and instructions
- Classify the role and select the right content
- Write a `resume_content_[company].json`
- Run `generate_resume.py` to produce the `.docx` and `.pdf`
- Verify the output is exactly 2 pages before finishing

Output files are named: `Your Name Resume MM.DD.YYYY Company Role.docx`

---

## Files in this repo

| File | What it is |
|---|---|
| `CLAUDE.md` | Instructions Claude reads automatically when you open this folder. Contains all formatting rules, workflow steps, and hard constraints. |
| `generate_resume.py` | The script that converts a JSON content file into a formatted `.docx`. Auto-detects your master resume. |
| `resume_instructions_template.md` | Template for your customization rules: role families, headline formulas, bullet priorities, role inclusion logic. |
| `master_resume_bullets_template.md` | Template for your bullet bank: every role, organized by company, with verified metrics. |
| `build_master.py` | Writes a starter `master_resume.docx` with the header table, styles and bullet list the generator needs. Use it if you don't have a master `.docx`. |
| `trace_check.py` | Fabrication check: every number in a content JSON (headline, summary, bullets, skills) must appear in `master_resume_bullets.md`. `python3 trace_check.py resume_content_[company].json`; exit 1 lists each untraceable figure. |
| `score_keywords.py` | Optional keyword report: `python3 score_keywords.py requirements_[company].json resume_content_[company].json` lists each weighted job-description term as hit or miss (and where it matched), the weighted coverage %, the misses by weight, and the share of bullets with a figure. Report only; always exits 0 (2 on bad input). |
| `requirements_example.json` | Example requirements file: a title plus weighted terms (`weight` 1-3, optional `alternates`). Copy it to `requirements_[company].json`. |
| `requirements.txt` | Python dependencies: `pip install -r requirements.txt`. |
| `tiering_instructions_template.md` | Optional template for a job evaluation framework: hard filters, scored soft signals, tier definitions, and a verdict-first output format. Copy to `tiering_instructions.md`. |
| `README.md` | This file. |

---

## Tips

- **The bullet bank is the most important file.** Invest time in writing complete, metrics-backed bullets for every role. This is what separates a good resume from a generic one.
- **Be specific in your instructions.** The more precisely you define role families and bullet priorities, the better Claude's selections will be.
- **Estimate line count before generating.** EB Garamond at 11pt wraps at roughly 95 characters per line. Before running the script, estimate total lines using this formula:

  ```python
  import math, re

  def estimated_lines(text, wrap=95):
      chars = len(re.sub(r'\*\*', '', text))  # strip bold markers
      return max(1, math.ceil(chars / wrap))

  total = sum(estimated_lines(b) for role in content["experience"] for b in role["bullets"])
  print(f"Estimated bullet lines: {total}")
  ```

  **Target range: 47–54 lines.** Below 47 is too thin. Above 54 will overflow. Adjust bullet count or length before generating rather than after.

- **Let the page check do the counting.** After exporting the PDF, `generate_resume.py` counts its pages and exits 1 unless there are exactly 2 (`⚠ Page count: 3 (expected 2)`). It also exits 1 if no PDF could be produced, so an unverified run never looks like a pass. Use `--pages 1` for a one-page resume, or `--no-pdf` to write only the `.docx` and skip the check.

- **Write output to its own folder.** `--outdir DIR` (for example `--outdir output/`) puts the `.docx`, PDF and page check in `DIR`, creating it if missing (a relative path is resolved against the current directory); the default is the script's folder. Recommended: a generated filename containing "MASTER" (such as "Mastercard") next to the script makes master auto-detection ambiguous. `--md` also writes `<output>.md` beside the `.docx`, a plain Markdown mirror (name, contact, links, then the resume, with `**bold**` kept) for pasting into ATS text boxes.

- **Verify every output visually.** Claude checks page count automatically, but render the PDF and scan for fill quality before sending. A resume that fits 2 pages but only fills 75% of the second page looks unfinished.
- **Keep your real files out of git.** The `.gitignore` in this repo excludes `*.docx`, `*.pdf`, `resume_content_*.json`, and your personal instruction files. Your actual resume files stay private.

---

## Troubleshooting

**PDF isn't generating:**
The script tries Microsoft Word first (via AppleScript on macOS, via `docx2pdf` on Windows), then falls back to LibreOffice on every platform. It prints which backend produced the PDF, or why each one failed. If neither works:
- Install LibreOffice: `brew install --cask libreoffice` on macOS, or the installer from libreoffice.org
- On Windows, make sure `docx2pdf` is installed (`pip install -r requirements.txt`)
- If Word is stuck on macOS, reset it: `osascript -e 'tell application "Microsoft Word" to close every document saving no'`
- If the output says `⚠ Word export failed ... (-1708)` on macOS, Word couldn't write to the output folder. Word is sandboxed and can only save where it has access (a folder under your home directory works; `/tmp` does not). Write the output somewhere in your home folder, or keep the LibreOffice PDF.

**Resume is 3 pages:**
Run the line estimation formula above before generating — it catches overflow before the script runs. If you're already over, the generator reports `⚠ Page count: 3 (expected 2)` and exits 1; ask Claude to tighten the longest bullets (shorten, not cut) rather than removing them entirely. You preserve the proof point while buying back lines.

**Font looks wrong in the PDF:**
Install EB Garamond from [fonts.google.com/specimen/EB+Garamond](https://fonts.google.com/specimen/EB+Garamond). The script swaps Word's built-in Garamond for EB Garamond so both the DOCX and LibreOffice PDF render identically.

**Master resume not found:**
Name your `.docx` `master_resume.docx` and put it in the same folder as `generate_resume.py`, or keep exactly one `.docx` with `MASTER` in its filename there, or set `MASTER_RESUME` to its full path. If the script says several files match, rename or move all but one.

**Master resume is missing structure:**
The script lists what it could not find (a style name or the header table). Add the named style in Word (Format → Style → New), or rebuild the header as described in Setup step 2.

---

## Credits

Workflow design and prompt engineering by [Navin Watumull](https://linkedin.com/in/navinwatumull). Built with [Claude Code](https://claude.ai/code).
