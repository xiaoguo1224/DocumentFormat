# DOCX Validation

Validation covers both rendered appearance and Word-native structure. The active profile defines semantic style names, caption labels, heading depth, bibliography numbering grammar, and other template-specific expectations; the validator must not hard-code the generic template's names.

## Structural checks

Check semantic styles, heading outline levels, multilevel numbering, body-list numbering, `pageBreakBefore`, caption `SEQ` fields/registered labels, TOC, PAGE fields, section behavior, bookmarks, and `REF`/`PAGEREF` target integrity.

### Bibliography-specific checks

When numeric bibliography numbering is required:

- the profile's reference style exists;
- it is bound to a native numbered-list `numId`/level;
- the resolved level text matches the profile (generic default `[%1]`);
- the resolved level suffix is `tab`;
- bibliography paragraph text does not still contain manually typed numbering prefixes;
- bibliography entries have stable reference bookmarks;
- body bibliography citations are native `REF` fields;
- bibliography `REF` fields use the paragraph-number switch (`\\n`) and hyperlink switch when required;
- every bibliography REF target bookmark exists;
- recognizable typed body citations fail strict validation when native cross-references are required.

## Visual checks

Render the document and inspect all pages (a contact sheet is acceptable for first pass). Check page geometry, typography, indentation/spacing, orphaned headings, captions, tables, TOC, headers/footers/page numbers, and section starts.

## Failure conditions

Treat required native features as FAIL when replaced by visual imitation, including typed heading numbers, typed figure/table numbers, manually typed bibliography `[1]`, manually typed TOC/page numbers, blank paragraphs used for chapter starts, manual list prefixes, or copied citation numbers that will not update.

## Result levels

- PASS: native feature exists and matches the active profile/template.
- WARN: structure exists but cannot be fully verified automatically, or an optional feature is absent.
- FAIL: required native structure is missing, malformed, has unresolved targets, or was replaced by static text.

Do not claim full compliance when only visual checks were performed.
