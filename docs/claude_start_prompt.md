# Claude Code Start Prompt

Copy this into Claude Code to start Phase 2.

```text
You are in /home/giulio/code/LexForward/moduli/formatter.

Act as a senior legal-document automation engineer working for Bergamo Legal.
Your task is to create a client-facing visual consent artifact for the letter
formatter.

This is Phase 2. Phase 1 is complete: do not re-derive the formatter from
scratch. Use the existing rules and code.

Read first, in this exact order:
1. CLAUDE.md
2. AGENTS.md
3. docs/client_format_review_goal.md
4. docs/html_review_artifact_spec.md
5. docs/client_approval_points.md
6. docs/regole_di_formattazione_in_lavorazione.md
7. docs/regole_formattazione_lettere.md
8. feedback.md
9. formatters/letters.py

Goal:
Create a static HTML page that can be sent to a client to agree on how Bergamo
Legal letters should be formatted.

Required output:
- out/html/lettera_fittizia_review.html
- out/html/lettera_fittizia_review_report.md

HTML requirements:
- Show one fictional Bergamo Legal letter, not real client content.
- Left side: a legend with one item per letter section.
- Right side/main area: a paper-like letter preview.
- Each section in the letter has a subtle colored background matching the legend.
- Each legend item shows quick information immediately.
- On hover and keyboard focus, each legend item reveals full formatting details:
  font, size, alignment, indentation, spacing before/after, bold, italic,
  keep-with-next behavior, template/header/footer inheritance, and relevant notes.
- The page must be self-contained enough to open directly in a browser.
- The visual design must be professional, sober, and suitable for a law firm.
- The page must be explanatory, not marketing.

Use current sources of truth:
- implemented values from docs/regole_di_formattazione_in_lavorazione.md;
- rationale from docs/regole_formattazione_lettere.md;
- client sensitivities from feedback.md;
- implementation details from formatters/letters.py.

Constraints:
- Do not modify previous_works/.
- Do not modify assets/Template_Vuoto.docx.
- Do not include real client/confidential facts.
- Do not invent new formatting rules.
- Do not silently change formatters/letters.py or the rule docs. If you find a
  mismatch while building the HTML, report it in the review report.
- Keep outputs outside previous_works/.

Before finishing:
- Open or inspect the generated HTML enough to verify that all legend items and
  highlighted sections are present.
- Check that hover/focus details are available for every legend item.
- Check mobile and desktop layout by reviewing the CSS/responsive rules or using
  local browser tooling if available.
- Report exactly what you created, what sources you used, and what remains for
  client approval.
```
