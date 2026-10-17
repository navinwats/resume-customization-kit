# Tiering Instructions

This file tells Claude how to evaluate a job description BEFORE any tailoring. It is optional. Use it when you are applying at volume and want to stay selective: every role gets a verdict, and only roles worth the effort get a tailored resume. Fill in every section with your own filters and preferences. The more specific you are, the more useful the verdicts.

Trigger: run this when you ask to "tier" a role. If you do not ask, Claude generates the resume immediately (see `CLAUDE.md`).

---

## Step 0: Verdict Rules

Every evaluation ends in one of four verdicts:

- **Tier 1 (Apply now, tailor carefully):** Passes every hard filter and scores strongly on soft signals.
- **Tier 2 (Apply):** Passes every hard filter with a middling soft-signal score.
- **Tier 3 (Skip):** Fails any hard filter, or passes but scores below your cutoff.
- **Verify first:** A hard filter cannot be confirmed from the JD. Claude lists exactly what to check and does not assign a tier until you answer.

You can override any verdict. When you do, Claude states the override in one line: `Overrode [verdict] because [reason].`

---

## Step 1: Hard Filters (any failure = Tier 3)

Dealbreakers. A single failure ends the evaluation. Define each one in your own terms. Delete rows you do not care about and add your own.

1. **Role family fit:** The role is in one of [your target role families, matching Step 0 of `resume_instructions.md`]. Fail if it is [role types you will not do].
2. **Seniority:** Title and scope are at [your minimum level] or above. Fail if [titles or scope that are a step down].
3. **Location / remote:** Passes if [remote / hybrid in [City] / onsite in [list of cities]]. Fail if [on-site requirement outside your range, relocation required, etc.].
4. **Compensation floor:** Posted or stated base is at least [$X]. If no range is posted, see Step 4 (do not guess).
5. **Industry exclusions:** Fail if the company operates in [industries you will not work in].
6. **Work authorization:** Fail if the JD requires [citizenship, clearance, or sponsorship status you cannot meet].
7. **Reporting line:** Passes if the role reports to [acceptable functions or titles]. Fail if it reports to [unacceptable ones].
8. **Company stage / size:** Passes if [stage, headcount, or funding range]. Fail if [outside that range].
9. **[Add your own dealbreakers here]**

---

## Step 2: Soft Signals (scored)

Only evaluated for roles that pass every hard filter. Score each signal from 0 to 2 (0 = absent or negative, 1 = partial or unclear, 2 = clear match), multiply by the weight, and add up.

| Signal | Weight | What a 2 looks like |
|---|---|---|
| Skills overlap with your bullet bank | [3] | [Most named requirements map to bullets you can verify] |
| Scope and ownership | [2] | [You would own [strategy / a program / a budget]] |
| Company quality and momentum | [2] | [Funding, growth, product reputation you care about] |
| Team and leadership | [1] | [Named manager or team you respect] |
| Compensation above floor | [1] | [At least [X]% above your floor] |
| Growth path | [1] | [Clear route to [next level]] |
| [Add your own signals] | [weight] | [description] |

**Thresholds** (set your own):

- **Tier 1:** score of [X] or higher out of [max]
- **Tier 2:** score of [Y] to [X minus 1]
- **Tier 3:** score below [Y]

---

## Step 3: What Each Tier Triggers

- **Tier 1:** Run the full resume workflow from `CLAUDE.md`. [Optional: also draft a cover note, flag a referral to chase, apply within [N] days.]
- **Tier 2:** Run the resume workflow with [a lighter touch / the standard bullet selection, no extra customization]. [Optional: batch these weekly.]
- **Tier 3:** Do not generate a resume. State the failing filter or the score that ruled it out.
- **Verify first:** Do not generate a resume. Ask the questions from the output format below and wait.

---

## Step 4: Unverifiable Facts (never guess)

A JD often omits things your filters depend on (comp, reporting line, remote policy, sponsorship, team size).

- If a fact cannot be confirmed from the JD text, Claude flags it as **UNVERIFIED**. It never assumes, infers from the company name, or fills in from memory.
- An unverified fact on a hard filter means the verdict is **Verify first**, not a pass and not a fail.
- An unverified fact on a soft signal scores 1 and is noted.
- Claude states where each confirmed fact came from (quote or section of the JD).

---

## Step 5: Required Output Format

Claude's response must follow this order. Verdict first, then the reasoning.

```
VERDICT: [Tier 1 / Tier 2 / Tier 3 / Verify first]
One sentence on why.

HARD FILTERS
1. [Filter name]: PASS / FAIL / UNVERIFIED. [Evidence from the JD, quoted or cited]
2. [Filter name]: PASS / FAIL / UNVERIFIED. [Evidence]
...

SOFT SIGNALS (only if all hard filters passed)
- [Signal]: [0-2] x [weight] = [points]. [One-line reason]
...
TOTAL: [score] / [max]

QUESTIONS FOR YOU (only if anything is UNVERIFIED)
- [Specific question, e.g. "The JD gives no salary range. What is the posted or expected base?"]

NEXT STEP
[What happens next per Step 3]
```

Keep it short. No resume content in a tiering response.

---

## Final Checklist

Before sending a verdict, verify:

- [ ] The verdict line is first
- [ ] Every hard filter has a result and evidence
- [ ] Nothing was assumed: every unknown is marked UNVERIFIED and turned into a question
- [ ] The tier matches the thresholds in Step 2
- [ ] The next step matches Step 3
- [ ] No resume was generated for a Tier 3 or Verify first result
