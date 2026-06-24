# Agent Rules

## Project boundary

This workspace contains a legal document formatter and the artifacts used to
reach client agreement on formatting rules.

Current work is **Phase 2: client-facing HTML format review**.

## Source of truth

- Current implemented style values:
  `docs/regole_di_formattazione_in_lavorazione.md`
- Rationale and corpus evidence:
  `docs/regole_formattazione_lettere.md`
- Client feedback and sensitivities:
  `feedback.md`
- Formatter implementation:
  `formatters/letters.py`
- Historical corpus:
  `previous_works/` read-only

## Engineering principles

- Build deterministic artifacts. Prefer a repeatable script over hand-editing a
  large HTML file if generation reduces drift.
- Keep style facts synchronized with the Python formatter and rules docs.
- Make visual explanations concrete: show actual colored sections on a letter,
  not abstract prose.
- Do not expose real client data. Use fictional legal-letter content.
- Put generated outputs under `out/`.
- Keep historical material untouched.

## Required deliverables

- `out/html/lettera_fittizia_review.html`
- `out/html/lettera_fittizia_review_report.md`

The report must say which docs/code were used as sources, what the HTML shows,
and what the client still needs to approve.

## Quality bar

The artifact is successful only if it can be sent to the client as a precise
format-consensus tool: every visible section has a matching legend item, and
every legend item exposes the formatting details needed to approve or contest
that section.

