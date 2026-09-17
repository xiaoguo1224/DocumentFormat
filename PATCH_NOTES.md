# DocumentFormat fix overlay

Base reviewed: `xiaoguo1224/DocumentFormat` master at `d8dcd301023dd38dd52a0f416f943e4a86e9602b`.

## Fixes

- Bibliography entries are normalized to native Word list numbering: `[%1]` + real TAB suffix.
- Manually typed `[1]`, `[2]`, `1.`, `1、` prefixes are removed once native numbering is installed.
- Every bibliography entry receives a stable `_FmtRefNNNN` bookmark.
- Body citations become `REF _FmtRefNNNN \\n \\h` fields, so citations reference the paragraph number rather than bibliography paragraph text.
- Reference repair is idempotent for existing bibliography bookmarks/native REF fields.
- Reference-section state now exits correctly at acknowledgements/appendices; later body text no longer inherits `References` style.
- Appendix body detection no longer mistakes ordinary appendix body text for an appendix heading.
- Formatter and validator are driven by `profile.yaml` semantic mappings instead of hard-coded generic style names/caption labels.
- Added safe template section geometry transfer (page size/margins/columns/page-number settings; header/footer relationship ids are not blindly copied).
- Inspector now reports reusable visual style metadata and section/page geometry in addition to native structures.
- Cross-reference repair keeps the existing formula-number repair path and skips destructive rewriting of complex paragraphs.
- Added `.gitignore`, dependency metadata, and regression tests.

## Regression tests

`pytest -q` -> `6 passed`.

Covered: native bibliography numbering, paragraph-number REF switch, section-state exit, repair idempotency, custom template/profile style mappings, and template page geometry transfer.

## Repository cleanup

After applying this overlay, remove tracked IDE/backup artifacts if present:

- `.idea/`
- `templates/generic-formal/template.docx.bak`
