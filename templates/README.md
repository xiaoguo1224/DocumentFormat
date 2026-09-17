# Templates

## Default common format

When the user does not provide another format template, use:

```text
templates/generic-formal/
├── template.docx
└── profile.yaml
```

`template.docx` is the actual reusable Word format baseline. It contains the native Word structures rather than only visual formatting, including:

- A4 page setup and the common margins;
- Chinese/English abstract styles;
- TOC title and TOC field;
- Heading 1-4 styles bound to one real multilevel numbering definition;
- `Heading 1` chapter page-break-before behavior;
- body paragraph style;
- native chapter-based figure/table `SEQ` caption examples;
- native `REF` cross-reference examples;
- front-matter Roman page numbering and body Arabic page-number restart;
- reference, acknowledgement, appendix, figure/table note, and formula styles;
- a three-line-table example.

`profile.yaml` stores only semantic requirements that cannot be represented reliably by the Word template alone, such as keyword punctuation/count limits, continuation-table rules, appendix numbering conventions, equation-editor requirements, and validation requirements.

## Add another reusable format

Create only one new directory:

```text
templates/
└── my-format/
    ├── template.docx
    └── profile.yaml     # optional
```

The new `template.docx` overrides the default common format for that task. Do not copy the generic profile unless a rule really applies to the new format; use a small profile containing only ambiguous semantic mappings or explicit overrides.

A user-supplied `.docx` / `.dotx` may also be used directly without permanently adding it to this folder.

## Recommended template preparation

Prepare real Word/WPS mechanisms in the template whenever possible:

- Heading 1-N styles;
- heading multilevel numbering;
- paragraph spacing and fonts;
- page-break-before where required;
- body/list styles;
- caption styles and `SEQ` fields;
- TOC field and TOC styles;
- headers/footers and `PAGE` fields;
- section breaks and page-number restarts;
- table styles.

A template with Word-native structures is preferred over one that only demonstrates the same appearance with ordinary text.
