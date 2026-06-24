# Claude Code Start Prompt

Copy this into Claude Code when starting the implementation.

```text
You are in /home/giulio/code/LexForward/moduli/formatter.

This is a clean restart. Build only the Bergamo Legal letters formatter.

Read first:
1. CLAUDE.md
2. AGENTS.md
3. docs/letter_formatter_goal.md
4. docs/letters_formatter_deliverables.md
5. previous_works/manifest.json

Use previous_works/manifest.json as the only corpus classification source.
Select only records where document_type is letters.

Task:
- Inspect the corresponding input/output/report/verification files.
- Infer the reusable formatting structure of letters.
- Write docs/regole_formattazione_lettere.md.
- Implement formatters/letters.py as a deterministic python-docx composer based
  on semantic sections, not a one-off converter.
- Use assets/Template_Vuoto.docx as the base template and preserve its
  letterhead/header/footer.
- Add a small validation path showing that new content can be rendered through
  the same schema.

Constraints:
- Do not modify previous_works/.
- Do not modify assets/Template_Vuoto.docx.
- Do not implement acts or other document families.
- Do not create an AI-heavy formatter.
- Keep generated outputs outside previous_works/.

Before finishing, report exactly what was inspected, what was implemented, and
what remains uncertain.
```

