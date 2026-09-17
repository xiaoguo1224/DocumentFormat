# Template Contract

## 1. Source of truth

A Word `.docx` / `.dotx` template is the authoritative format source whenever one is supplied.

Do not duplicate its visual rules into separate Markdown files unless the rule cannot be represented or reliably inferred from Word.

## 2. What to preserve from a template

Preserve whenever present:

- `word/styles.xml`
- `word/numbering.xml`
- `word/settings.xml`
- `word/theme/*`
- `word/fontTable.xml`
- section properties (`sectPr`)
- headers and footers
- page-number configuration
- TOC fields and TOC styles
- caption styles and sequence fields
- table styles
- custom style names and based-on relationships
- keep-with-next / keep-together / widow-control / page-break-before behavior

Prefer modifying content around those structures rather than reconstructing them.

## 3. How to add a reusable format

Recommended directory:

```text
templates/<template-name>/
├── template.docx
└── profile.yaml        # optional
```

`template.docx` should contain the real Word styles and native structures.

`profile.yaml` should be small. It exists only to declare semantic mappings or rules that are ambiguous from the template itself.

## 4. What belongs in profile.yaml

Good profile values:

- semantic mapping of custom style names;
- which heading level starts a new page;
- whether figure/table numbering includes chapter number;
- caption labels (`图`, `表`, `式`, `Figure`, `Table`);
- numbering separator (`-`, `.`, etc.);
- TOC depth;
- whether references/acknowledgements/appendices start on new pages;
- body-list styles when template names are ambiguous;
- special section names that should map to heading levels or TOC entries.

Avoid copying these into YAML unless they truly need to override the template:

- font name;
- font size;
- margins;
- line spacing;
- paragraph spacing;
- direct indentation;
- bold/italic;
- table border details.

Those should normally remain inside the Word template.

## 5. Ad-hoc templates

A user-supplied template does not need to be copied into the skill folder.

The skill may inspect and use the uploaded template directly for that task.

If the user later wants it reusable, save it under `templates/<name>/template.docx` and optionally add a minimal profile.

## 6. Existing content document + template

When both exist:

- content comes from the content document;
- formatting system comes from the template;
- user instructions override both.

Do not blindly copy source paragraph formatting into the target template.

Map source semantics to target styles instead.

## 7. Template examples versus rules

Sample text inside a template may exist only to demonstrate formatting.

Do not treat sample wording as required content unless the user or template explicitly marks it as fixed wording.

Repeated structural behavior is stronger evidence than one-off direct formatting.
