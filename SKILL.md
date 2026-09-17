---
name: document-format-writer
description: Template-driven Word formatter for formal documents. Use when the user wants to create, rewrite, or reformat a .docx according to a Word template or explicit formatting requirements. The final document must preserve or create native Word structures such as heading styles, multilevel numbering, page-break-before, captions, bibliography numbering, fields, TOC, page numbers, and cross-references instead of visually imitating them with plain text.
---

# Document Format Writer

## Core principle

The Word template is the primary source of truth for formatting.

If the user does not supply a template, use `templates/generic-formal/template.docx` together with `templates/generic-formal/profile.yaml` as the default common formal-document baseline.

Do not maintain a separate institution-specific Markdown rule file when the same rule already exists in the template. Markdown references define how to interpret, preserve, implement, and validate Word structures.

A document is compliant only when it both looks correct and has correct Word-native structure. Visual imitation alone is not acceptable for numbering, captions, bibliography numbering, TOC, page numbers, page-break-before, lists, or cross-references.

## Inputs and priority

The user may provide a content `.docx`, a format `.docx`/`.dotx`, plain text/Markdown, a small semantic profile YAML, and explicit instructions.

Apply rules in this order:

1. explicit user instruction;
2. current user-supplied template and paired profile;
3. bundled `generic-formal` template/profile when no user template is supplied;
4. safe fallback behavior in this skill.

The profile is semantic configuration, not a duplicate visual style sheet. Font, size, spacing, margins, borders, and most indentation belong in the Word template unless an explicit override is truly necessary.

## Required references

Read:

- `references/template-contract.md`
- `references/word-native-structures.md`
- `references/validation.md`

## Fast default path

For ordinary “通用模板格式化” requests, use the deterministic formatter first:

```powershell
python scripts/fast_format_docx.py --source <input.docx> --output <output.docx> --profile templates/generic-formal/profile.yaml --ensure-toc --strict-cross-references
python scripts/validate_docx.py <output.docx> --profile templates/generic-formal/profile.yaml --require-toc --require-captions --require-multilevel --require-crossrefs --require-reference-numbering
```

The formatter imports template-native styles/numbering/settings, maps document semantics through the active profile, applies safe section geometry from the template, converts typed figure/table captions to native `SEQ` fields, normalizes bibliography entries to a native `[1]`/`[2]` numbered list with TAB suffix, builds bookmarks, converts body citations to native clickable `REF` fields, adds PAGE/TOC fields when required, and marks fields for refresh on open.

Use `scripts/repair_cross_references.py` when only bibliography/cross-reference repair is needed.

If Microsoft Word refreshes fields and rewrites template-native definitions, run `scripts/restore_template_parts.py <docx> <docx> --template <template.docx>` before final validation.

For visual QA, render once, build a contact sheet, inspect every page on the sheet, and open only suspicious pages at full size.

## Bibliography contract

Bibliography handling is structural, not cosmetic.

When the active profile requires numeric references:

- bibliography entries must use one native Word numbered-list definition;
- list text must be `[1]`, `[2]`, `[3]`... via numbering, never typed into entry text;
- the numbering level must use a TAB suffix so the entry body begins after a real tab stop;
- bibliography paragraphs must use the profile's reference style;
- manually typed prefixes such as `[1]`, `1.`, `1、` must be removed once native numbering is installed;
- every bibliography entry must have a stable bookmark target;
- body citations such as `[1]` must use native `REF` fields to the bibliography bookmark with the paragraph-number switch (`\\n`) and hyperlink switch (`\\h`), so renumbering/reordering references updates citations after field refresh;
- a `REF` field that returns bibliography paragraph text instead of the paragraph number is invalid;
- do not use a separate manually maintained `SEQ` counter for bibliography numbers when the bibliography itself is a native numbered list.

## Template-driven workflow

### No supplied template

1. Use `templates/generic-formal/template.docx` and its profile.
2. Preserve native multilevel headings, page-break-before, captions, bibliography numbering, TOC/PAGE fields, section/page-number configuration, and reusable semantic styles.
3. Replace sample content rather than recreating the formatting manually.

### Supplied template

1. Inspect the template first with `scripts/inspect_docx.py`.
2. Treat the template itself as the base formatting source whenever practical.
3. Preserve styles, numbering definitions, heading bindings, section properties, headers/footers, page numbering, caption definitions/fields, TOC fields/styles, table styles, theme/font bindings, and required settings/custom XML.
4. Load the supplied profile when present. Do not hard-code `Heading 1`, `Figure Caption`, `图`, `表`, or other generic-template names in the formatting engine.
5. Reuse existing native structures before creating new ones.
6. If the template uses custom style names, map semantic roles through the profile or structural inspection.

### Existing content document

Preserve content unless rewriting is requested. Map paragraphs to semantic roles such as title/front matter, Heading 1-N, body, body lists, figure/table captions, notes, references, acknowledgements, and appendices. Remove manual numbering only when replacing it with a genuine Word-native mechanism.

Section state must be explicit. Entering `参考文献` must not cause following `致谢` or appendix body paragraphs to remain reference paragraphs.

### Plain text / Markdown

Determine semantic structure first, then insert content using template styles and native numbering/field structures.

## New-format workflow

A reusable format is normally:

```text
templates/<template-name>/
├── template.docx
└── profile.yaml    # optional and small
```

The profile stores semantic mappings/ambiguities such as custom style names, section-title aliases, caption labels, TOC depth, bibliography numbering semantics, and section behavior that cannot be reliably inferred from the Word file.

Do not create another long Markdown file that restates visual formatting already encoded in the template.

## Native Word requirements

At minimum:

- headings use real outline styles;
- numbered headings use a real multilevel list;
- chapter starts use `pageBreakBefore` when required;
- ordinary lists use Word list definitions;
- figure/table captions use native `SEQ` fields;
- bibliography numbering uses native numbered-list structure when the profile requires numeric references;
- bibliography citations use `REF` fields to list-number bookmarks;
- TOC uses a Word TOC field;
- page numbers use PAGE fields;
- other cross-references use REF/PAGEREF or equivalent fields;
- section breaks are used only for true section-level changes.

## No silent degradation

Never silently replace a requested native feature with a visual imitation, including typed heading numbers, typed figure/table numbers, manually typed bibliography `[1]`, manually typed TOC/page numbers, blank paragraphs for chapter pagination, or copied citation numbers that do not update.

## Validation gate

Before delivery run structural validation. The validator must be profile-driven rather than assuming English Word style names or Chinese caption labels.

For numeric bibliography formats, validation must confirm:

- the reference style is bound to a native list;
- its level text matches the profile (generic default: `[%1]`);
- its suffix is a real TAB;
- reference entry text contains no manually typed numeric prefix;
- body bibliography REF fields use the paragraph-number switch;
- every REF target bookmark exists.

A final document is not compliant merely because it renders correctly.

## Output

Default final artifact: `.docx`. Markdown is not the final deliverable unless explicitly requested.
