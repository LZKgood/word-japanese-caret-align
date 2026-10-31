# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/).

## [Unreleased]

- Nothing yet.

## [1.0.0] - 2026-10-31

### Added

- `Alt + Down` vertical caret alignment for Japanese intensive-listening notes.
- Full-width-space annotation grid, allowing multiple annotations on the same
  line under different words.
- Automatic Overtype mode when inserting before existing annotations, so
  existing text is not pushed away.
- Automatic creation of an annotation line when no next paragraph exists.
- Support for reusing an existing empty paragraph or an existing grid line.
- Detection of old indent-style annotation lines with a migration prompt.
- Chinese copy-paste source and English importable `.bas`.
- Documentation: installation, FAQ, design notes, publishing checklist.
- Static validation script and GitHub Actions workflow.
- MIT license, privacy notice, disclaimer, security policy, and contribution
  guidelines.

### Fixed

- Runtime error **4608** ("Value out of range") when the document had no next
  paragraph. The old code inserted a paragraph with `InsertAfter` at the
  document end; it now uses `Paragraph.Range.InsertParagraphAfter` and moves
  the caret into the new paragraph.
- Runtime error **438** ("Method or data member not found"). The old code used
  `Paragraph.Index`, which does not exist in Word. The macro now locates the
  current paragraph by comparing `Range.Start` values in the `Paragraphs`
  collection.
- Incorrect caret placement after an annotation line already contained text:
  the macro now distinguishes grid lines from ordinary paragraphs and positions
  the caret at the closest character boundary.
- Annotation insertion before existing text no longer shifts the line; the
  macro enables Word's Overtype mode when needed.

### Tested

- Microsoft Word 16.0 on Windows.
- Existing empty next paragraph.
- No next paragraph (annotation line created automatically).
- Next paragraph containing ordinary text (dedicated annotation line inserted).
- Second annotation on the same line, including the Overtype replacement path.
- Packaged English `.bas` import and Chinese source import (GBK test copy).

### Known limitations

- Normal single-column text only; tables, text boxes, columns, and vertical
  writing are not supported.
- Full-width annotations are required for exact column alignment; half-width
  Latin characters can shift by half a column.
- A very long annotation may wrap and break the grid; use a table for heavy
  layouts.
- WPS and Word for Mac are not officially tested.

## Development history

- **0.1** — Initial indent-based caret alignment experiment.
- **0.2** — Fixed first-line indent behavior and discovered the 4608 document-end
  insertion bug.
- **1.0.0** — Redesigned same-line annotations around a full-width-space grid,
  added Overtype protection, fixed 438/4608, packaged for GitHub.
