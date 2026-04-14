# Resume Customization Instructions

This file tells Claude how to customize your resume for any job description. Fill in every section with your own rules and preferences. The more specific you are, the better the output.

---

## Step 0: Role Family Classification

Before doing anything else, classify the role into one of your defined families. Everything — headline, summary, lead bullets — flows from this classification.

Define your role families here. Examples:

- **Product Marketing (Inbound):** Roles focused on positioning, messaging, audience insights, competitive intelligence, analyst relations
- **GTM / Launch:** Roles focused on go-to-market strategy, launch execution, cross-functional coordination
- **Consumer / Growth:** Roles focused on lifecycle, acquisition, retention, activation, freemium
- **B2B / SaaS / Enterprise:** Roles focused on sales enablement, enterprise pipeline, customer success partnership
- **Integrated Marketing / Campaign:** Roles focused on integrated campaigns, creative briefs, brand storytelling, events
- **[Add your own families here]**

---

## Step 1: JD Analysis

Weight job description requirements in this order:

1. Direct role match — skills and experience explicitly named
2. Functional match — adjacent experience that transfers
3. Proof via outcomes — specific results that validate the claim
4. Scope and complexity — scale, team size, budget, geography
5. Supporting evidence — anything else that builds the picture

Mine your full bullet bank. The strongest proof point is often in an unexpected role.

---

## Step 2: Headline

One line. Identity-based. Not a copy of the job title.

Define your headline formula per role family:

- **Product Marketing:** `[Your Identity] | GTM Strategy | [Your Key Categories]`
- **Integrated Marketing:** `[Your Identity] | [Campaign Type], [Strength] & [Strength]`
- **Consumer / Growth:** `[Your Identity] | Lifecycle, Growth & [Your Angle]`
- **B2B / SaaS:** `[Your Identity] | GTM Strategy | B2B SaaS & Enterprise`
- **[Add your own formulas here]**

Example: `Product Marketing Leader | GTM Strategy | Consumer Tech, Creative AI & SaaS`

---

## Step 3: Summary

Exactly **two sentences**. Never open with "I" or your name. Bold 1–2 anchor phrases only.

- **Sentence 1:** Category, scale, breadth — who you are and where you've done it
- **Sentence 2:** The specific value most relevant to THIS role

Use the language of the role family, not generic PMM language.

Example structure:
> **[Identity phrase]** with [X]+ years driving [core strength] across [notable companies/scale]. [Specific value statement most relevant to this role].

---

## Step 4: Your Title at Current Company

Define how to display your current role based on what you're applying for:

- **For most external roles:** Use `[Cleaner/clearer title]`
- **For [specific type] roles:** Use `[Your actual internal title]`
- **When in doubt:** Use the version that is clearest to an outside hiring manager

Always spell out "Senior" — never abbreviate as "Sr" or "Sr."

Date range: `[Start month/year]-Present`

---

## Step 5: Role Inclusion Rules

Define which companies/roles to include based on the type of role you're applying for. Be specific.

**[Company 1]:** Always include. Always listed first.

**[Company 2]:** Include for [role types]. Key stat to reference: [your key stat].

**[Company 3]:** Include for [role types]. Prioritize when JD mentions [specific keywords].

**[Company 4]:** Include only for [very specific niche]. Do not include for general roles.

**[Older/smaller companies]:** Include selectively when [specific conditions]. Almost never otherwise.

---

## Step 6: Bullet Selection

### Lead bullets (non-negotiable)

Set the first 3–4 bullets for your primary role based on role family BEFORE selecting anything else:

- **Product Marketing / Inbound:** Lead with [your strongest PMM bullets]
- **GTM / Launch:** Lead with [your strongest GTM bullets]
- **Consumer / Growth:** Lead with [your strongest growth/lifecycle bullets]
- **B2B / SaaS:** Lead with [your strongest B2B bullets]
- **Integrated Marketing:** Lead with [your strongest campaign bullets]

### Remaining bullet selection

Rank remaining bullets by:
1. Direct match to JD requirements
2. Strongest verified impact (metrics, benchmarks, scale)
3. Cross-functional relevance
4. Executive readability

Quantification target: 60–75% of bullets should include a metric, benchmark, or scale marker.

---

## Step 7: Length

- **Target:** Fill both pages. ~18 bullets total. Never exceed 2 pages.
- **Primary role:** 6–8 bullets
- **Secondary role:** 5–6 bullets
- **Supporting roles:** 2–3 bullets each
- **Hard limit:** 2 pages. Always verify with `mdls -name kMDItemNumberOfPages` after generating.

---

## Step 8: Skills

- Use 4–6 skill groups with Title Case headers
- Mirror JD terminology where truthful
- Comma-separated items within each group
- Tools listed last

---

## Step 9: Final Checklist

Before generating, verify:

- [ ] Role family classification is correct and drives the entire resume
- [ ] No fabricated data, titles, metrics, or outcomes
- [ ] Summary is exactly two sentences
- [ ] Current company is listed first with correct title
- [ ] No duplicate bullets across roles
- [ ] No em dashes (—) anywhere — use commas, semicolons, or parentheses instead
- [ ] En dash (–) used only when a dash is truly necessary
- [ ] "Senior" spelled out fully, never "Sr"
- [ ] Skills headers in Title Case
- [ ] Bullet count is ~18 (to stay within 2 pages after PDF rendering)

---

## JSON Output Format

```json
{
  "headline": "Your headline here",
  "summary": "**Bolded anchor phrase** with sentence one content. Sentence two content.",
  "experience": [
    {
      "company": "COMPANY NAME",
      "title": "Full Title Here",
      "location": "City, ST",
      "dates": "Mon YYYY-Mon YYYY",
      "bullets": [
        "**Bullet Header:** Bullet body with **bolded metric** inline."
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
