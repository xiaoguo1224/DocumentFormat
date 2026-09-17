# DocumentFormat template-runtime repair

Base reviewed: `master` at `e5c779a0998c55c4f174a6187103ea4cd7dd451b`.

## Fixes

- Custom templates can infer semantic heading/body/reference style mappings when no profile is supplied.
- Source headings are recognized by Word outline level, not only `Heading N` / `标题N` style names.
- One-section source documents can be split into front matter + body at the first body Heading 1 when page-number formats differ.
- Section mapping selects template sections by front/body page-number format instead of applying the last template section everywhere.
- Page size/margins/columns/doc-grid/page-number restart/first-page behavior are transferred from the matching template section.
- Default, first-page, and even-page headers/footers are cloned with valid package relationships; relationship-backed images/hyperlinks are preserved.
- A deliberately blank target header/footer clears source leftovers.
- `evenAndOddHeaders` / `mirrorMargins` settings are preserved.
- Caption alignment is inherited from the active template caption style; formatter code no longer forces center alignment.
- PAGE field synthesis is controlled by `pagination.position` and does not force bottom-center.
- Table first-row repetition is controlled by the profile; table-cell paragraphs are no longer globally overwritten with body style.
- Custom-template inference detects TOC/PAGE presence, PAGE location, section page-number formats, caption labels, heading page-break-before, and native/visual numeric bibliography evidence.
- Native bibliography numbering is only enforced when the active template/profile requires it; visual `[1]` reference samples can be upgraded to native numbering automatically.
- Validator accepts `--template`, uses inferred semantics, checks front/body page-number formats and template header/footer presence.

## Verification

- `pytest -q tests/test_template_runtime.py` -> `9 passed`
- `python -m py_compile scripts/*.py` -> PASS
- Full custom-template/no-profile pipeline -> PASS, including native bibliography `REF ... \\n \\h`
- Header image relationship cloning -> PASS

## Repository cleanup included by apply scripts

The apply scripts remove old one-time patch artifacts if still tracked:

- `PATCH_NOTES.md`
- `apply_fix.ps1`
- `apply_fix.sh`
- `templates/generic-formal/template.docx.bak`
