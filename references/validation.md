# DOCX Validation

Validation must cover both visual appearance and Word-native structure.

## 1. Required structural checks

### Styles

Confirm:

- expected semantic styles exist;
- heading paragraphs actually use heading/outline styles;
- repeated body/caption/reference elements use reusable styles rather than only direct formatting.

### Numbering

Confirm:

- `word/numbering.xml` exists when numbering is required;
- multilevel definitions exist for numbered headings;
- heading levels are linked to the appropriate numbering levels;
- body lists use appropriate numbering definitions;
- heading numbers are not merely typed into text.

### Pagination

Confirm required chapter-level headings use `pageBreakBefore` when that is the template's behavior.

Do not accept blank paragraphs as an equivalent implementation.

### Captions

Confirm:

- figure/table caption paragraphs use caption styles or template-equivalent styles;
- automatic captions contain `SEQ` fields or equivalent native fields;
- figure and table counters are independent;
- chapter-aware numbering is implemented when required;
- when reusable Word/WPS caption labels are required, `word/settings.xml` registers the expected labels (for the generic template: `图`, `表`, and `式`).

### TOC

Confirm a real `TOC` field exists when a TOC is required.

Confirm source headings are structurally valid.

### Cross-references

When automatic cross-references are required, confirm `REF` / `PAGEREF` or equivalent fields exist.

Also confirm every `REF` target bookmark exists and scan ordinary body paragraphs for recognizable typed citations (`[n]`, `图x-y`, `表x-y`, `式x-y`). Do not count TOC `PAGEREF` fields as proof that body cross-references are native.

### Page numbers

Confirm footer/header page numbering uses Word fields rather than typed digits.

### Section behavior

Confirm section breaks are used only where section-level formatting requires them.

Check page-number restarts, orientation, margins, and header/footer linking where relevant.

## 2. Visual checks

For speed, inspect all rendered pages first on a labeled contact sheet. Full-size inspection is required only for pages whose text, tables, figures, breaks, or margins cannot be judged reliably from the sheet.

Render or inspect the document for:

- margins and page size;
- heading typography and spacing;
- body typography, indentation, and line spacing;
- orphaned headings;
- figure/table placement;
- caption spacing;
- table borders and repeated header rows;
- TOC layout;
- headers, footers, and page-number position;
- section-start behavior.

## 3. Failure conditions

Treat the document as structurally non-compliant if any required native feature was replaced by a visual imitation, including:

- typed heading numbers instead of multilevel numbering;
- typed figure/table numbers instead of fields;
- manually typed TOC;
- typed page numbers;
- blank paragraphs used to force chapter pagination;
- manual list prefixes instead of Word lists when editable lists are required.

## 4. Validation result

A useful validator should report:

- PASS: native feature exists and appears correctly configured;
- WARN: structure exists but cannot be fully verified automatically;
- FAIL: required native structure is missing or replaced by static text.

Do not claim full compliance when only visual checks were performed.
