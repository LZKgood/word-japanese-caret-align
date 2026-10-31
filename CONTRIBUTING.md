# Contributing

Thanks for considering a contribution. This project is a Word VBA macro; small,
well-tested changes are preferred over large rewrites.

## Before you start

- Read [README.md](README.md), [docs/DESIGN.md](docs/DESIGN.md), and
  [docs/FAQ.md](docs/FAQ.md).
- For security issues, follow [SECURITY.md](SECURITY.md) instead of opening a
  public issue.
- By contributing, you agree that your contribution is licensed under the
  project's [MIT License](LICENSE) and that you have the right to submit it.

## Reporting bugs

Use the bug report issue template. Include:

- Word version (File → Account/About Word) and Windows version;
- macro version (`VERSION` / release tag);
- which source you used: English `.bas` import or Chinese copy-paste file;
- exact steps, expected result, and actual result;
- the error number and the highlighted VBA line, if any;
- whether all old modules were removed and only one copy of the macro exists;
- a minimal synthetic document structure (no real personal notes).

If the bug involves the annotation grid, state whether the target line was
created by v1.0 or by an older indent-style version.

## Development setup

There is no build step. The deliverables are source files:

- `src/CaretDownAligned.bas` — ASCII/English version, directly importable into
  the VBA editor. Must start with:

  ```vba
  Attribute VB_Name = "CaretDownAlignedModule"
  Option Explicit
  ```

- `src/CaretDownAligned.zh-CN.vba.txt` — UTF-8 (no BOM) Chinese copy-paste
  version. It must **not** contain `Attribute VB_Name` or start with
  `Option Explicit` (the paste target may already have one). It should include
  the same logic as the `.bas`, with Chinese comments and messages.

Keep the two versions logically identical. A change is not complete until both
are updated.

## Coding rules and known traps

- Use late-bound `Object` variables for Word objects (for example
  `Dim sel As Object`). This avoids compile-time "method or data member not
  found" errors in mixed Word/WPS environments.
- Numeric Word enum values are used in this code (`1`, `5`, `7`, `12`, etc.)
  with comments. Do not replace them with named constants unless you have
  tested the target Word versions.
- **Never use `Paragraph.Index`.** Word's `Paragraph` object does not expose an
  `Index` property. Locate the current paragraph by comparing
  `Paragraphs(i).Range.Start` with the current paragraph's `Range.Start`.
- **Never split the Japanese paragraph at the caret.** Do not insert a
  paragraph mark or line break at the caret; this pushes the rest of the
  sentence to the next line. Create/navigate paragraphs around the current
  paragraph instead.
- **Do not insert a paragraph at the end of the document with
  `Range(End, End).InsertAfter`.** That path raised runtime error 4608. Use
  `Paragraph.Range.InsertParagraphAfter` and move the caret into the new
  paragraph.
- The annotation line is a full-width-space grid. Do not "clean up" trailing
  spaces; they are the horizontal anchors.
- Overtype mode is intentional. Do not silently switch it off after
  positioning, and do not force it on for ordinary paragraphs.
- Avoid external dependencies, network calls, file writes outside the active
  document, and `Shell`/`WScript`/`CreateObject` calls.

## Encoding and line endings

- English `.bas`: ASCII only, CRLF line endings. The import path must not
  depend on a non-English code page.
- Chinese `.vba.txt`: UTF-8 without BOM, CRLF preferred, for manual copy-paste.
- Do not save a Chinese `.bas` as UTF-8 with BOM for direct import; VBA import
  may fail or garble. If you need a GBK test copy, keep it out of the repository.

## Testing

Test in a **copy** of a Word document, with all old macro modules removed.
Minimum matrix:

1. Japanese sentence + existing empty next paragraph.
2. Japanese sentence + no next paragraph (single paragraph document).
3. Japanese sentence + next paragraph containing ordinary text.
4. First annotation on the grid line, then a second annotation to the left
   (Overtype path) and to the right (Insert path).
5. Full-width Japanese/Chinese annotations; note any half-width behavior.
6. Different font sizes and page margins.
7. Old indent-style annotation line detection.
8. Undo behavior after a jump.
9. Save/reopen the document; ensure the layout persists.

Record the Word version, Windows version, and results in the PR.

## Static validation

Run:

```powershell
pwsh -File scripts/validate.ps1
```

The script checks source presence, the `.bas` header, forbidden regression
patterns, and basic documentation links. It does not compile VBA; a successful
run is not a substitute for testing in Word.

## Pull requests

- Keep the change focused.
- Update both source variants and any affected docs.
- Update [CHANGELOG.md](CHANGELOG.md) for user-visible changes.
- Fill in the pull request template, including test evidence.
- Do not commit documents, logs, screenshots with personal content, secrets, or
  generated binaries.

## Commit messages

Use short, imperative subject lines, for example:

```text
fix: avoid Paragraph.Index runtime error
docs: clarify overtype behavior
```

## Questions

For usage questions, open a discussion or issue with a minimal example. For
security concerns, use the private channel in [SECURITY.md](SECURITY.md).
