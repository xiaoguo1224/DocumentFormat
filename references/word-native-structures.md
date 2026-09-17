# Word Native Structures

## Principle

If Microsoft Word has a native semantic mechanism for a feature, use that mechanism rather than plain text that only looks equivalent.

## 1. Headings

A heading must use a real heading/outline style.

For numbered headings, use one multilevel numbering definition bound to heading levels.

Typical hierarchy:

- level 1 -> `1`
- level 2 -> `1.1`
- level 3 -> `1.1.1`
- level 4 -> `1.1.1.1`

Do not type the number into the paragraph text.

Subordinate levels must restart according to the multilevel numbering definition when a parent level changes.

Prefer the template's existing `numbering.xml` definitions.

## 2. Heading pagination

When a chapter or other heading must start on a new page, prefer paragraph property `pageBreakBefore`.

Also preserve or set `keepWithNext` and `keepTogether` where appropriate.

Do not create chapter starts with blank paragraphs.

Use a section break only when section-level behavior changes, such as:

- page-number restart;
- header/footer change;
- orientation change;
- margin change;
- columns or other section properties.

## 3. Body lists

Ordinary numbered/bulleted lists are separate from heading numbering.

They should use Word numbering/list definitions rather than manually typed prefixes when editable lists are required.

Do not reuse heading numbering definitions for body lists unless the template intentionally does so.

## 4. Figure captions

A figure caption should use:

- a dedicated caption paragraph style;
- a native `SEQ` field or equivalent Word field for the sequence;
- chapter-aware numbering when the template requires it;
- editable caption text after the generated number.

Example rendered result:

`图1-1 系统总体架构`

`图1-1` must not be a fixed typed string when automatic numbering is required.

Figure and table counters must be independent.

## 5. Table captions

Use a dedicated table-caption style and an independent native sequence field.

Example:

`表2-3 模型性能对比`

Preserve continuation-table behavior from the template when present.

## 6. Caption labels

Use the labels required by the active template/profile, for example:

- `图`
- `表`
- `式`
- `Figure`
- `Table`

Do not hard-code Chinese labels globally; the active template decides them.

When the template expects the labels to appear in Word/WPS **Insert Caption** behavior, register them as native caption definitions in `word/settings.xml` (`w:captions` / `w:caption`) in addition to using the corresponding `SEQ` fields. A visual prefix plus `SEQ` alone is not enough when reusable caption-label metadata is required.

## 7. Formulas

When formulas are numbered, use native Word structures where practical and preserve the template's numbering convention.

Prefer Word equation/OMML objects for editable formulas. If formula numbers use fields, retain those fields.

Do not rasterize equations merely to preserve appearance.

## 8. Cross-references

When automatic references are required, use `REF`, `PAGEREF`, or equivalent native Word cross-reference fields.

Examples:

- `如图2-3所示`
- `见表4-1`
- `见第3.2节`

The referenced object should have a stable bookmark/reference target when needed.

For existing typed citations, use stable ASCII bookmark names and clickable `REF ... \\h` fields. Apply the same mechanism to bibliography entries, figures, tables, and numbered equations. Preserve OMML equation content; only repair its number/bookmark/reference structure.

## 9. Table of contents

Use a real Word `TOC` field.

The TOC must derive from real heading/outline styles or the template's configured styles.

Do not manually type page numbers or dot leaders.

## 10. Page numbers

Use native Word page-number fields (`PAGE`, and where appropriate `NUMPAGES`, etc.).

Preserve section-specific numbering formats and restarts.

## 11. Styles

Use reusable Word styles for repeated semantic elements.

Prefer style-based formatting to repeated direct formatting.

Typical roles:

- Heading 1-N
- body
- captions
- notes/sources
- TOC levels
- references
- appendix headings/body

Custom template style names are valid and should be preserved.

## 12. Fields and cached display values

A `.docx` may contain both field instructions and cached visible results.

If the runtime cannot force Microsoft Word to recalculate fields, preserve the valid field code and provide a reasonable cached result where possible.

Do not remove the field simply because updating it is inconvenient.

## 13. Native-structure-first rule

Implementation preference:

1. reuse template-native structure;
2. modify existing styles/numbering/fields;
3. create missing OOXML structure;
4. only use static text for features that are genuinely static.
