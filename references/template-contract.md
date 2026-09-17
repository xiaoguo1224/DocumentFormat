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

Font names/sizes, margins, line/paragraph spacing, ordinary indentation, bold/italic, alignment, table borders, and header/footer layout normally belong in `template.docx`, not YAML.

## Template without a profile

A custom template must remain usable without a companion YAML profile. Infer semantics conservatively from Word-native evidence in this order:

1. outline levels and numbering bindings for Heading 1-N;
2. style usage in template sample content;
3. conventional/custom semantic names such as `正文`, `图题`, `表题`, `参考文献`;
4. safe generic fallbacks only when the template provides no stronger evidence.

Do not silently force generic style names such as `Heading 1`, `Body Text Generic`, `Figure Caption`, or `References` onto a custom template.

## Content + template

Content comes from the content document; formatting semantics come from the target template/profile; explicit user instructions override both. Map source semantics to target styles rather than copying source direct formatting.

Source semantic recognition should prefer Word outline levels and native structure over English style-name regexes. A source style named `章标题` with outline level 0 is still a level-1 heading.

## Template examples

Sample text demonstrates formatting unless explicitly marked fixed. Repeated native structure and repeated style usage are stronger evidence than one-off direct formatting.

## Section, header, and footer transfer

When front matter and body require different page-number formats, a one-section source document may be split at the first body Heading 1 so the front matter and body can inherit separate section rules.

Transfer section-native properties from the matching template section, including page size, margins, columns, document grid, page-number format/restart, and first-page behavior. Match front/body sections by their page-number format when possible rather than simply applying the template's last section to every source section.

Clone header/footer stories through package relationships, including relationship-backed images and hyperlinks. Do not copy raw `r:id` values between packages. Copy odd/default, first-page, and even-page stories when the template actually contains content, and preserve `evenAndOddHeaders`/`titlePg` behavior.

## Captions, page numbers, and tables

Caption alignment must come from the caption style/template. Do not force center alignment in formatter code.

If a PAGE field must be synthesized because the template does not already provide one, its story and alignment come from `pagination.position`; existing template page-number fields win.

Do not force every table's first row to repeat unless the active profile requires continued-table header repetition. Do not overwrite every table-cell paragraph with body style. Apply table/table-body styles only when the active template/profile explicitly maps them.
