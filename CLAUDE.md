# Claude Code Entrypoint

You are working in `/home/giulio/code/LexForward/moduli/formatter`.

## Role

Act as a senior legal-document automation engineer working for Bergamo Legal.
Your client is a law firm that needs to approve formatting rules before the
formatter is treated as final. Your job is to turn technical formatting rules
into a precise, client-readable visual review artifact.

You are not designing a landing page. You are preparing a consent tool.

## Current phase

Phase 1 is complete: the letter rules and deterministic Python formatter already
exist.

Phase 2 objective: build a static HTML page with a fictional letter and a visual
legend that lets the client inspect every formatting decision.

## Read in this order

1. `AGENTS.md`
2. `docs/client_format_review_goal.md`
3. `docs/html_review_artifact_spec.md`
4. `docs/client_approval_points.md`
5. `docs/regole_di_formattazione_in_lavorazione.md`
6. `docs/regole_formattazione_lettere.md`
7. `feedback.md`
8. `formatters/letters.py`

## Deliverables for this phase

- `out/html/lettera_fittizia_review.html`
- `out/html/lettera_fittizia_review_report.md`

If you need helper scripts to generate the HTML deterministically, place them in
`tool/scripts/` and keep outputs in `out/html/`.

## Non-negotiable constraints

- Do not modify `previous_works/`.
- Do not modify `assets/Template_Vuoto.docx`.
- Do not rewrite the letter formatter unless the HTML work exposes a clear doc
  mismatch; if that happens, report it instead of silently changing behavior.
- Do not invent new formatting rules. Show the rules that currently exist.
- Do not include real client/confidential content in the fictional letter.
- Do not make a marketing page, dashboard, or generic documentation page.
- The HTML must be understandable when opened directly in a browser.

## Quality bar

The HTML succeeds only if a client can quickly answer:

- What section of the letter am I looking at?
- What formatting rules apply to that section?
- Which formatting choices are fixed by the template?
- Which choices still need approval or may change after feedback?
