---
name: document-format-writer
description: Template-driven Word formatter for formal documents. Use when the user wants to create, rewrite, or reformat a .docx according to a Word template or explicit formatting requirements. The final document must preserve or create native Word structures such as heading styles, multilevel numbering, page-break-before, captions, fields, TOC, page numbers, and cross-references instead of visually imitating them with plain text.
---

# Document Format Writer

## Core principle

The Word template is the primary source of truth for formatting.

If the user does not supply a template, use `templates/generic-formal/template.docx` together with `templates/generic-formal/profile.yaml` as the default common formal-document baseline. This default template implements the existing common requirements: formal long-document structure, true multilevel heading numbering, chapter page-break-before, native figure/table captions, TOC/page-number fields, and the established paragraph/caption/reference/appendix rules.

Do not maintain a separate institution-specific Markdown rule file when the same rule already exists in the template. Markdown references in this skill only define how to interpret, preserve, implement, and validate Word structures.

A document is compliant only when both of these are true:

1. it looks correct;
2. its Word-native structure is correct.

Visual imitation alone is not acceptable for numbering, captions, TOC, page numbers, page-break-before, lists, or cross-references.

## Inputs

The user may provide any combination of:

- a content `.docx` to be reformatted;
- a format template `.docx` or `.dotx`;
- plain text / Markdown content to be turned into a `.docx`;
- an optional template profile YAML containing only semantic mappings or overrides that cannot be inferred reliably from the template;
- explicit user instructions that override the template.

## Priority

Apply rules in this order:

1. explicit user instruction;
2. current user-supplied template and its paired profile;
3. `templates/generic-formal/template.docx` and its profile when no user template is supplied;
4. safe fallback behavior in this skill.

If a template and profile conflict, prefer explicit profile values only for fields the profile intentionally overrides. Do not use a profile to duplicate all visual formatting already present in the template.

## Required references

Read these files when performing document work:

- `references/template-contract.md`
- `references/word-native-structures.md`
- `references/validation.md`

## Fast default path

For ordinary requests that say “通用模板格式化” and do not provide an institution-specific template, start with the bundled deterministic formatter instead of writing a document-specific script:

```powershell
python scripts/fast_format_docx.py --source <input.docx> --output <output.docx> --ensure-toc --strict-cross-references
python scripts/validate_docx.py <output.docx> --require-toc --require-captions --require-multilevel --require-crossrefs
```

The formatter preserves the source file, imports the generic template's native styles/numbering/settings, maps common heading/list/caption/reference/equation semantics, converts typed figure/table/equation/reference numbers to native fields, creates stable bookmarks, replaces resolvable body citations with clickable `REF` fields, adds a PAGE field when missing, and marks fields for refresh on open. `scripts/repair_cross_references.py` is the reusable second stage when only cross-reference repair is needed.

If Microsoft Word is used to refresh fields, immediately run `scripts/restore_template_parts.py <docx> <docx>` before final validation. Word may rewrite heading-numbering bindings and caption-label registrations while saving; this restoration preserves refreshed cached field results while reinstating the template-native definitions.

Use its JSON summary as the first routing decision. Escalate to custom document-specific code only when validation reports a real FAIL, `fallback_styles` is nonzero, important semantics were not detected, the user supplied a special template, or visual QA shows an actual defect. Do not escalate solely for cached field text or validator warnings inside a valid TOC.

For visual QA, render once, build a contact sheet with `scripts/make_contact_sheet.py`, inspect every page on the sheet, and open only suspicious pages at full size. This satisfies full-page coverage without one tool call per page.

## Template-driven workflow

### A0. When no template is supplied

1. Use the fast default path above first. Start from `templates/generic-formal/template.docx` only when the fast formatter must be extended or bypassed.
2. Read `templates/generic-formal/profile.yaml` for semantic rules that are not fully represented by Word styles/fields.
3. Preserve its native multilevel heading numbering, page-break-before behavior, TOC/PAGE fields, caption fields, section/page-number configuration, and reusable paragraph styles.
4. Replace the placeholder/sample content with the user's content rather than recreating the formatting manually.

### A. When a template is supplied

1. Inspect the template before drafting or reformatting.
2. Treat the template itself as the base document whenever practical.
3. Preserve its existing:
   - styles;
   - numbering definitions;
   - heading-to-numbering bindings;
   - section properties;
   - headers and footers;
   - page numbering;
   - caption styles and fields;
   - TOC fields and TOC styles;
   - table styles;
   - theme/font bindings;
   - custom XML or settings required by Word.
4. Identify semantic styles by structure and sample usage rather than by visible appearance alone.
5. Replace or insert content without flattening the template into direct formatting.
6. Reuse existing native structures before creating new ones.

### B. When formatting an existing content document

1. Preserve user content unless rewriting is requested.
2. Map content paragraphs to template semantic styles:
   - title/front matter;
   - Heading 1-N;
   - body text;
   - lists;
   - figure captions;
   - table captions;
   - notes/sources;
   - references;
   - appendices.
3. Remove manually typed numbering only when it is being replaced by genuine Word numbering.
4. Convert visual-only constructs to native Word structures when required.
5. Do not copy all direct formatting from the source document into the template.

### C. When creating a document from plain text or Markdown

1. Determine semantic structure first.
2. Insert content into the template using the template's styles.
3. Apply native heading numbering and list structures.
4. Insert figures/tables with native captions when required.
5. Use Word fields for TOC, page numbers, caption sequences, and cross-references when applicable.

## New-format workflow

The built-in default format lives at:

`templates/generic-formal/template.docx`

A new format should normally be added as:

`templates/<template-name>/template.docx`

Optionally add:

`templates/<template-name>/profile.yaml`

The profile is only for semantic mappings and explicit overrides that cannot be reliably inferred from the Word file.

Do NOT create another long Markdown file that restates font sizes, spacing, margins, and numbering already encoded in the Word template.

The user may also upload a template ad hoc. It does not need to be installed permanently into this folder for the skill to use it.

## Semantic mapping

When a template contains custom style names, determine which styles correspond to semantic roles such as:

- document title;
- abstract title/body;
- Heading 1 / Heading 2 / Heading 3 / Heading 4;
- body text;
- figure caption;
- table caption;
- figure/table note;
- TOC 1 / TOC 2 / TOC 3;
- references;
- acknowledgements;
- appendix heading/body.

Prefer existing template styles even when their names are not standard English Word names.

## Native Word requirements

Follow `references/word-native-structures.md`.

At minimum:

- heading hierarchy must use actual heading styles or template-equivalent outline styles;
- numbered headings must use a real multilevel numbering definition;
- chapter-start behavior should use `pageBreakBefore` when appropriate;
- body lists must use Word list/numbering structures;
- figure/table captions must use caption styles and `SEQ`-based fields or equivalent native fields;
- TOC must use a Word TOC field;
- page numbers must use Word PAGE fields;
- cross-references should use REF/PAGEREF or equivalent native fields when required;
- section breaks should be used only for true section-level changes.

## No silent degradation

Never silently replace a requested native feature with a visual imitation.

Examples of unacceptable degradation:

- multilevel heading numbering -> typed text such as `1.2.3`;
- figure caption field -> typed `图1-1`;
- table caption field -> typed `表2-3`;
- TOC field -> manually typed TOC;
- page number field -> typed number;
- page-break-before -> empty paragraphs;
- body list -> typed bullets/numbers when an editable Word list is required.

If the runtime cannot implement a required native feature correctly, report that limitation instead of pretending the document is fully compliant.

## Validation gate

Before final delivery, run structural validation according to `references/validation.md`.

If scripts are available, use:

- `scripts/inspect_docx.py` to inspect a template or generated document;
- `scripts/validate_docx.py` to validate required Word-native structures.

A final document must not be treated as compliant only because its rendered appearance looks correct.

## Output

Default final artifact: `.docx`.

Do not use Markdown as the final deliverable unless the user explicitly asks for Markdown.

When the user asks only for a reusable template system or skill configuration, provide the skill/template files rather than generating a content document.
