# Formatter Workspace

This workspace is now in **Phase 2**.

Phase 1 produced the reusable Bergamo Legal letters formatter:

- `docs/regole_formattazione_lettere.md`
- `docs/regole_di_formattazione_in_lavorazione.md`
- `formatters/letters.py`
- `validate_letters.py`
- `feedback.md`

## Current objective

Create a **client-facing HTML review artifact** that explains the letter
formatting rules visually.

The artifact must show one fictional Bergamo Legal letter with colored section
backgrounds and a left-side legend. Each legend item gives a short summary and,
on mouseover/focus, reveals full formatting details: font, size, alignment,
spacing, indentation, bold/italic, keep-with-next behavior, template inheritance,
and any other detail a client may contest during formatting approval.

The goal is not marketing. The goal is consent: the client must be able to see
and approve how letters will be formatted before we update the final docs and
Python formatter.

## Read first

1. `CLAUDE.md`
2. `AGENTS.md`
3. `docs/client_format_review_goal.md`
4. `docs/html_review_artifact_spec.md`
5. `docs/client_approval_points.md`
6. `docs/claude_start_prompt.md`

## Source material

- Use `formatters/letters.py` and
  `docs/regole_di_formattazione_in_lavorazione.md` as the current formatting
  truth.
- Use `docs/regole_formattazione_lettere.md` for rationale and corpus context.
- Use `feedback.md` to understand known client sensitivities.
- Keep `previous_works/` read-only.
