# Agent Rules

## Project boundary

This repository is only for the letter formatter restart. Treat the historical
corpus as read-only evidence and the new formatter as production code.

## Source of truth

- Corpus classification: `previous_works/manifest.json`.
- Template: `assets/Template_Vuoto.docx`.
- Target family for this task: `letters` only.
- Reference layout: historical `previous_works/output_*.docx` for letters.

## Engineering principles

- Prefer deterministic Python over AI-generated layout.
- Separate semantic extraction from DOCX composition.
- Store formatting decisions in code, not prompts.
- Keep all generated outputs outside `previous_works/`.
- Use reports as clues, not as a substitute for inspecting DOCX structure.
- Add concise tests or validation scripts where they materially reduce risk.

## Required deliverables

- `docs/regole_formattazione_lettere.md`
- `formatters/letters.py`
- A small example or validation command that shows the formatter can create a
  DOCX from replacement section content.
- A short run report describing which corpus letters were inspected and what
  remains uncertain.

## Quality bar

The formatter is successful only if a future developer can create a new letter
by changing semantic content fields without rewriting layout logic.

