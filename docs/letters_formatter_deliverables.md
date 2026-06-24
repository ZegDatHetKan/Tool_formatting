# Letter Formatter Deliverables

## 1. Rules document

Create `docs/regole_formattazione_lettere.md`.

It must explain:

- the canonical letter skeleton;
- each semantic block and when it appears;
- style rules for every block;
- template usage and what must be inherited from `Template_Vuoto.docx`;
- normalization rules;
- cases that should set `needs_review`;
- known uncertainties supported by corpus evidence.

The document must cite historical example IDs from the manifest, for example
`001`, `003`, `025`, `027`, `028`, `032`.

## 2. Python formatter

Create `formatters/letters.py`.

The formatter must expose a reusable API similar to:

```python
from formatters.letters import LetterDocument, render_letter

letter = LetterDocument(
    date_place="Bergamo, 24 giugno 2026",
    delivery_method="A mezzo PEC",
    recipient_block=["Spett.le ..."],
    subject="Oggetto: ...",
    opening="Gentili Signori,",
    body_blocks=[...],
    closing="Distinti saluti,",
    signature_block=["Avv. ..."],
    attachments=[],
)

render_letter(letter, "assets/Template_Vuoto.docx", "out/example.docx")
```

The exact class names can change if the code stays equally clear, but the
formatter must be content-driven and reusable.

## 3. Validation

Add a lightweight validation path. It may be a script, CLI flag, or test file.
It must demonstrate:

- the template opens;
- output DOCX is created;
- historical corpus remains untouched;
- at least three representative letter structures can be expressed through the
  semantic schema.

## 4. Report

Create a short report in `out/` or `docs/` explaining:

- which letter IDs were inspected;
- which formatting rules are confirmed;
- which rules remain uncertain;
- what must be reviewed before relying on the formatter in production.

