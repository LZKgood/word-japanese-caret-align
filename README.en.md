<p align="center">
  <img src="assets/banner.svg" alt="Word Japanese listening caret alignment" width="100%">
</p>

# Word Japanese Listening · Caret Vertical Alignment (CaretDownAligned)

[中文](README.md) | [English](README.en.md)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Word 2016+](https://img.shields.io/badge/Word-2016%2B-blue.svg)](docs/INSTALL.md)
[![VBA](https://img.shields.io/badge/VBA-macro-green.svg)](src/CaretDownAligned.bas)
[![No telemetry](https://img.shields.io/badge/telemetry-none-brightgreen.svg)](PRIVACY.md)
[![Version](https://img.shields.io/badge/version-1.0.0-informational.svg)](CHANGELOG.md)

A Microsoft Word VBA macro for Japanese intensive-listening notes. Put the caret after any word in a Japanese sentence and press `Alt + Down`; the caret drops to the next line at the same horizontal column, without splitting the sentence or pushing the remaining text down.

> This macro modifies document layout (inserts full-width spaces and toggles Overtype mode). Back up your documents first, and read [DISCLAIMER.md](DISCLAIMER.md) and [PRIVACY.md](PRIVACY.md).

---

## Demo

```text
先生と学生が話しています。
        ^ caret here, after 学生
Press Alt+Down

先生と学生が話しています。
            学生
            ^ caret aligned under 学生
```

The same annotation line can hold multiple words:

```text
先生と学生が話しています。
老师  学生  正在说话。
```

## Features

- **Vertical caret alignment**: `Alt + Down` places the caret under the current word on the next line.
- **No line splitting**: it never inserts a paragraph break at the caret, so the rest of the Japanese sentence stays put.
- **Multiple annotations on one line**: a full-width-space grid line is created; later jumps return to the same line at the new word's column.
- **Overwrite protection**: when text already exists to the right, the macro enables Word's Overtype mode so existing annotations are not pushed away.
- **Works without an existing next line**: creates an annotation line after the Japanese paragraph.
- **Local only**: no network calls, no telemetry, no data collection.
- **Auditable**: complete VBA source, no third-party components.

## Requirements

- Windows 10 / 11
- Microsoft Word 2016 / 2019 / 2021 / Microsoft 365 desktop
- VBA macros enabled
- Normal single-column main text; tables, text boxes, columns, and vertical writing are not supported
- Full-width Chinese/Japanese annotations recommended; half-width Latin characters may shift by half a column
- WPS and Word for Mac are not officially tested

Tested with **Microsoft Word 16.0 on Windows** in these scenarios: existing empty next line, no next line, normal next paragraph, and a second annotation on the same line.

## Install

### Option A: import the English `.bas` (recommended)

1. Download [`src/CaretDownAligned.bas`](src/CaretDownAligned.bas).
2. In Word, press `Alt + F11`.
3. `File` → `Import File…`, choose `CaretDownAligned.bas`.
4. `Debug` → `Compile VBAProject`.
5. Assign a shortcut:
   - `File` → `Options` → `Customize Ribbon` → `Keyboard shortcuts: Customize…`
   - Category: `Macros`; select `CaretDownAlignedFinal`
   - Press `Alt + Down`; click `Assign`
   - Save in `Normal.dotm` (global) or the current document.
6. Save the template/document.

The English `.bas` uses English UI messages. Its logic is identical to the Chinese version.

### Option B: paste the Chinese code

1. Open [`src/CaretDownAligned.zh-CN.vba.txt`](src/CaretDownAligned.zh-CN.vba.txt).
2. In the VBA editor, insert a new module.
3. Keep any existing `Option Explicit` line at the top; paste the code below it.
4. Compile, then assign `Alt + Down` to `CaretDownAlignedFinal`.

The Chinese file is UTF-8 text for copy-paste. A direct `.bas` import on a non-Chinese Windows may garble Chinese strings, which is why the importable file is ASCII English.

See [docs/INSTALL.md](docs/INSTALL.md) for macro security, update, and uninstall steps.

## Usage

1. Type a Japanese sentence:
   ```text
   先生と学生が話しています。
   ```
2. Put the caret after the word you want to annotate (for example, after `学生`), and press `Alt + Down`.
3. On the first press, an annotation line made of full-width spaces is created. Type the annotation.
4. Go back to the Japanese line, put the caret after another word (`先生`), and press `Alt + Down` again. The caret returns to the same annotation line at that word's column. If there is already text to the right, the macro switches to Overtype mode.
5. Press `Insert` when done to return to insert mode.

To align to the **left edge** of a word, put the caret before its first character; to align to the **right edge**, put it after the last character. The macro aligns exactly to the caret column.

### Upgrading from pre-1.0

Pre-1.0 versions used a first-line indent to create the annotation line. That line cannot be reused; it has no columns to the left of its first character. When the macro detects one, it asks you to delete it. Copy the old text, delete the entire line, call `Alt + Down` again from the Japanese line, and paste/type the annotation.

## How it works

Word's flowing text has no virtual columns: an empty line only has a caret position at its paragraph start (controlled by indentation) or between existing characters. The macro therefore:

1. Measures the caret's horizontal position relative to the text boundary.
2. Creates an annotation line filled with full-width spaces (`U+3000`). Each space is one full-width character wide, forming a character grid.
3. Moves the caret to the closest character boundary at the target position.
4. Enables Overtype when text exists to the right, so typing replaces padding spaces instead of shifting annotations.
5. Never inserts a paragraph break at the caret; it only navigates/creates paragraphs around it.

See [docs/DESIGN.md](docs/DESIGN.md) for details and trade-offs.

## Limitations

| Scenario | Support | Notes |
|---|---|---|
| Normal single-column text | ✅ | Main use case |
| Multiple annotations on one line | ✅ | Full-width text recommended |
| Existing / missing next line | ✅ | Reuses or creates an annotation line |
| Half-width Latin annotations | ⚠️ | May shift by half a column |
| Tables, text boxes, columns | ❌ | The macro beeps and exits |
| WPS / Word for Mac | ⚠️ | Not officially tested |
| Vertical writing | ❌ | Different coordinate semantics |
| Changing annotation font size | ⚠️ | Grid column width changes |
| Very long annotations | ⚠️ | May wrap; use a table for heavy layout |
| Old indent-style annotation lines | ❌ | Delete and regenerate |

The annotation line contains many full-width spaces as anchors. Do not delete them unless you no longer need alignment. For heavy or publication-quality layouts, a two-row table (original / annotation, one column per word) is more robust.

## FAQ

- **438 "Method or data member not found"**: pre-1.0 used `Paragraph.Index`, which does not exist in Word. v1.0 locates the paragraph by iterating `Paragraphs` and comparing `Range.Start`. Delete all old modules and keep only the new code.
- **4608 "Value out of range"**: the old version inserted a paragraph at the document end with `InsertAfter`; v1.0 uses `InsertParagraphAfter`.
- **Caret does not land under the word**: make sure the target line is a v1.0 grid line and start from the Japanese line, not the annotation line. Old indent-style lines must be recreated.
- **Two macros with the same name**: remove the old module and rebind `CaretDownAlignedFinal`.

More: [docs/FAQ.md](docs/FAQ.md).

## Repository layout

```text
src/     VBA source (ASCII .bas + UTF-8 Chinese copy-paste file)
docs/    install, FAQ, design, publishing
assets/  banner
scripts/ static validation
.github/ issue/PR templates and CI
```

## Contributing

Issues and pull requests are welcome. Read [CONTRIBUTING.md](CONTRIBUTING.md) and [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md) first. Report security issues privately per [SECURITY.md](SECURITY.md); do not paste sensitive documents into issues.

## License and notices

- Code: [MIT License](LICENSE).
- Chinese legal summary (license/disclaimer/privacy/trademarks): [docs/LEGAL.zh-CN.md](docs/LEGAL.zh-CN.md).
- Risks, backup advice, and trademark notice: [DISCLAIMER.md](DISCLAIMER.md) and [NOTICE.md](NOTICE.md).
- No data collection or network access: [PRIVACY.md](PRIVACY.md).
- Not affiliated with Microsoft. Microsoft Word is a trademark of Microsoft Corporation.

Version `1.0.0` · [CHANGELOG.md](CHANGELOG.md)
