# Claude Code Entrypoint

You are working in `/home/giulio/code/LexForward/moduli/formatter`.

This is a clean restart. Do not rely on previous failed implementation choices.
The useful artifacts are:

- `previous_works/manifest.json`
- `previous_works/input_*.docx`
- `previous_works/output_*.docx`
- `previous_works/report_*.md`
- `previous_works/verification_*.md`
- `assets/Template_Vuoto.docx`

## Mission

Derive a robust, reusable Python formatter for Bergamo Legal letters.

The final result must include:

1. A written Markdown document explaining the formatting rules for letters.
2. A deterministic Python script that reproduces the letter structure by filling
   predefined semantic sections and paragraphs.

## Working order

1. Read `AGENTS.md`.
2. Read `docs/letter_formatter_goal.md`.
3. Read `docs/letters_formatter_deliverables.md`.
4. Read `previous_works/manifest.json`.
5. Select only manifest records with `document_type == "letters"`.
6. Inspect the matching input/output/report/verification files.
7. Infer the shared letter skeleton and the variation points.
8. Write the rules document.
9. Implement the Python formatter.
10. Validate it on representative letters from the corpus.

## Non-negotiable constraints

- Do not modify files inside `previous_works/`.
- Do not modify `assets/Template_Vuoto.docx`.
- Do not implement `acts` or `other_pending_name`.
- Do not build an AI-heavy formatter.
- Do not hardcode one historical letter as the solution.
- Do not use output reports as proof if the generated DOCX has not been opened
  or inspected programmatically.

## Expected implementation direction

Use `python-docx` unless a better local dependency already exists. Open
`assets/Template_Vuoto.docx`, preserve its header, footer, margins, and
letterhead assets, clear only the empty body content, then append formatted
paragraphs according to a semantic letter schema.

