# Template Contract

## Source of truth

A supplied `.docx`/`.dotx` is the authoritative visual-format source. Preserve styles, numbering, settings, theme/font table, section properties, headers/footers, page numbering, TOC/caption fields/styles, table styles, custom style names/relationships, and pagination controls whenever present.

Do not duplicate visual rules into Markdown/YAML when Word already represents them.

## Reusable format

```text
templates/<template-name>/
├── template.docx
└── profile.yaml        # optional, small, semantic
```

Good profile values include semantic style mappings, section-title aliases, heading depth/new-page behavior, caption labels/chapter level/separator, TOC depth, bibliography numbering grammar (`[%1]`, TAB suffix), and other rules that cannot be inferred reliably.

Font names/sizes, margins, line/paragraph spacing, ordinary indentation, bold/italic, and border details normally belong in `template.docx`, not YAML.

## Content + template

Content comes from the content document; formatting semantics come from the target template/profile; explicit user instructions override both. Map source semantics to target styles rather than copying source direct formatting.

## Template examples

Sample text demonstrates formatting unless explicitly marked fixed. Repeated native structure is stronger evidence than one-off direct formatting.

## Safe section transfer

Page geometry/page-number settings may be transferred from corresponding template sections. Do not blindly copy header/footer relationship ids between packages; relationships must be cloned correctly or the template itself should be used as the base document.
