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

You need Python 3 and two packages:

```bash
pip install python-docx docx2pdf
```

You also need **Microsoft Word** installed (Mac or Windows) for PDF export. On Mac, [LibreOffice](https://www.libreoffice.org/download/) works too and is more reliable:

```bash
brew install --cask libreoffice
```

### 2. Add your master resume `.docx`

Place your existing resume `.docx` in this folder. Name it one of these ways:
- `master_resume.docx` (simplest)
- Anything with `MASTER` in the filename (e.g., `Jane Smith MASTER Resume.docx`)

The generator uses this file purely for its formatting — fonts, margins, styles. Your content comes from the JSON.

### 3. Fill in the template files

Copy and rename the two template files:

```bash
cp resume_instructions_template.md resume_instructions.md
cp master_resume_bullets_template.md master_resume_bullets.md
```

Then edit both files to reflect your own career, rules, and preferences. See the templates for guidance on what each section should contain.

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
| `README.md` | This file. |

---

## Tips

- **The bullet bank is the most important file.** Invest time in writing complete, metrics-backed bullets for every role. This is what separates a good resume from a generic one.
- **Be specific in your instructions.** The more precisely you define role families and bullet priorities, the better Claude's selections will be.
- **Verify every output.** Claude checks page count automatically, but always open the `.docx` to review the content before sending.
- **Keep your real files out of git.** The `.gitignore` in this repo excludes `*.docx`, `*.pdf`, `resume_content_*.json`, and your personal instruction files. Your actual resume files stay private.

---

## Troubleshooting

**PDF isn't generating:**
The script tries LibreOffice first, then falls back to Word. If neither works:
- Make sure LibreOffice is installed: `brew install --cask libreoffice`
- If using Word and it's stuck, reset it: `osascript -e 'tell application "Microsoft Word" to close every document saving no'`

**Resume is 3 pages:**
Claude runs a generate/check/adjust loop automatically. If it still overflows, ask Claude to tighten the longest bullets rather than cutting them — you usually lose less proof that way.

**Font looks wrong in the PDF:**
Install EB Garamond from [fonts.google.com/specimen/EB+Garamond](https://fonts.google.com/specimen/EB+Garamond). The script swaps Word's built-in Garamond for EB Garamond so both the DOCX and LibreOffice PDF render identically.

**Master resume not found:**
Make sure your `.docx` is in the same folder as `generate_resume.py` and has `MASTER` in the filename, or is named `master_resume.docx`.

---

## Credits

Workflow design and prompt engineering by [Navin Watumull](https://linkedin.com/in/navinwatumull). Built with [Claude Code](https://claude.ai/code).
