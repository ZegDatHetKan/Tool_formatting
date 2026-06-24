# Letter Formatter Goal

## Goal

Infer the stable structure of Bergamo Legal letters from historical examples and
turn it into a reusable deterministic formatter.

The formatter must compose letters from semantic content such as date, delivery
method, recipient, subject, opening, body sections, closing, signature, and
attachments. A user should be able to replace those fields and receive a
properly formatted DOCX using the same layout.

## Corpus scope

Use `previous_works/manifest.json` to identify the letter records. At the time
of this setup, the manifest contains the letter family among 32 historical jobs.
The first implementation must ignore non-letter records.

Important: the manifest is the classification source of truth. Do not use a
separate copied planning list as a second classification source.

## What to infer

For every letter example, inspect:

- page setup inherited from `assets/Template_Vuoto.docx`;
- whether body text duplicates or omits letterhead material;
- paragraph order and block boundaries;
- fonts, sizes, bold/italic usage, alignment, indentation, spacing;
- handling of `Oggetto`;
- recipient block placement;
- numbered lists and lettered lists;
- closings, signatures, and attachments;
- recurring placeholders and review warnings.

## What not to infer

Do not infer rules for acts, mediation documents, ricorsi, memorie, querele, or
the third pending family. They can be documented as out of scope if encountered.

Do not infer legal content edits. This task is about structure and formatting.

