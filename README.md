# Letter Formatter Workspace

This workspace is a clean restart for deriving a reusable Bergamo Legal letter
formatter from approved historical examples.

## Current objective

Read the historical `letters` examples in `previous_works/`, infer the stable
formatting structure, and produce:

- `docs/regole_formattazione_lettere.md`: human-readable formatting rules.
- `formatters/letters.py`: deterministic Python code that can compose new
  letters by replacing semantic section content.

The formatter must not be a one-off converter for the existing files. It must be
a reusable letter skeleton with explicit semantic sections and hardcoded
formatting rules.

## Source material

- `previous_works/manifest.json` is the source of truth for document type,
  subject, and file paths.
- Use only records where `document_type` is `letters` for the first pass.
- Reference outputs are `previous_works/output_*.docx`.
- Raw inputs are `previous_works/input_*.docx`.
- Reports and verification files are supporting evidence only.
- `assets/Template_Vuoto.docx` is the canonical blank letterhead template.

## Read first

1. `CLAUDE.md`
2. `AGENTS.md`
3. `docs/letter_formatter_goal.md`
4. `docs/letters_formatter_deliverables.md`
5. `docs/claude_start_prompt.md`

