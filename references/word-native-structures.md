# Word Native Structures

## Principle

If Microsoft Word has a native semantic mechanism for a feature, use that mechanism rather than plain text that only looks equivalent. Template-native structure wins; the active profile supplies only semantic mappings/overrides that cannot be inferred reliably.

## Headings and pagination

A heading must use a real heading/outline style. Numbered headings use one multilevel numbering definition bound to heading levels; numbers are not typed into paragraph text. Lower levels restart through the numbering definition.

When a chapter must start on a new page, prefer `pageBreakBefore`, with `keepWithNext`/`keepTogether` where appropriate. Use section breaks only for real section-level changes such as page-number restart, header/footer change, orientation, margins, or columns.

## Body lists

Ordinary numbered/bulleted lists are separate from heading numbering. Use native list definitions; do not reuse heading numbering unless the template intentionally does so.

## Figure/table captions

Captions use dedicated semantic styles plus `SEQ` fields. Chapter-aware numbering uses the template/profile's heading level and separator. Figure and table counters are independent. When Insert Caption compatibility is required, register labels in `word/settings.xml` (`w:captions`).

## Bibliography and citations

For a numeric bibliography profile:

- bibliography entries use one native Word numbered-list definition;
- the level text is profile-driven (generic default: `[%1]`);
- the list suffix is a real TAB (`w:suff w:val="tab"`), with an actual numbering tab stop;
- bibliography entry text must not contain a typed `[1]`, `1.`, `1、`, etc. once native numbering is installed;
- each bibliography entry gets a stable ASCII bookmark;
- body citations use `REF <bookmark> \\n \\h` (paragraph-number + hyperlink switches) so they retrieve the numbered-list value, not the bibliography paragraph text;
- do not maintain a separate SEQ counter for bibliography numbers when native list numbering is the source of truth.

A plain `REF <bookmark>` to a whole bibliography paragraph is not sufficient because it can resolve to paragraph text rather than the paragraph number.

## Formulas

Prefer editable OMML equations. Preserve or create native numbering fields when the template requires numbered formulas. Do not rasterize equations merely to preserve appearance.

## Other cross-references

Use `REF`, `PAGEREF`, or equivalent native fields. Targets must have stable bookmarks. Existing complex paragraphs containing fields, drawings, comments, footnotes/endnotes, content controls, bookmarks, tabs/breaks, or mixed run formatting should not be destructively rebuilt just to create a cross-reference; report/skip ambiguous cases instead.

## TOC and page numbers

Use a real `TOC` field driven by real heading styles. Page numbers use `PAGE` (and related) fields and preserve section-specific numbering formats/restarts.

## Styles and fields

Use reusable styles for repeated semantics. A `.docx` can contain field instructions and cached visible values; keep the field even when the runtime cannot force Word to recalculate it. Mark fields dirty/update-on-open where appropriate.

## Native-structure-first order

1. reuse template-native structure;
2. modify existing styles/numbering/fields;
3. create missing OOXML structures;
4. use static text only for genuinely static content.
