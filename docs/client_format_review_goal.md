# Client Format Review Goal

## Purpose

Build a client-facing HTML artifact that makes the letter formatter visible and
negotiable before it becomes final.

The client should be able to review a fictional letter, inspect the formatting
rules section by section, and give precise feedback. After client agreement, the
approved decisions will be folded back into:

- `docs/regole_formattazione_lettere.md`
- `docs/regole_di_formattazione_in_lavorazione.md`
- `formatters/letters.py`

## Audience

The page is for legal professionals and operational stakeholders, not engineers.
It must be precise enough for formatting approval but readable without knowing
Python or DOCX internals.

## Core user experience

- A fictional letter is shown as a paper preview.
- The left legend lists all semantic sections.
- Each legend item has a matching color.
- The matching section in the letter uses the same color as a subtle background
  highlight.
- Quick facts are visible in the legend.
- Full technical details appear on hover and keyboard focus.

## What the page must communicate

For each section:

- semantic role;
- font family;
- font size;
- alignment;
- indentation;
- spacing before/after;
- bold and italic behavior;
- keep-with-next / anti-orphan behavior where relevant;
- whether it comes from the template or body composer;
- client-facing notes and unresolved approval points.

## What the page must not do

- It must not use real client facts.
- It must not change the formatter.
- It must not present tentative rules as already approved.
- It must not hide known client-sensitive choices such as OGGETTO layout,
  recipient appellative handling, title hierarchy, or signature spacing.

