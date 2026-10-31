# Privacy Policy / 隐私说明

**CaretDownAligned** is a local Microsoft Word VBA macro. It does not collect,
transmit, or sell personal data.

## What the macro accesses

To align the caret, the macro reads from the active Word document:

- the current selection and its horizontal position;
- the font name and font size at the caret;
- paragraph start positions and paragraph formatting;
- short excerpts of the active paragraph (to detect the full-width-space
  annotation grid and any existing annotation text to the right of the caret);
- whether the selection is in a table, a text box, or the main text story.

It modifies only the active document:

- inserts full-width spaces (`U+3000`) to build the annotation grid;
- sets paragraph indentation/alignment for the annotation line;
- toggles Word's Overtype mode;
- writes a short message to Word's status bar.

## Network and telemetry

- No network requests.
- No analytics, telemetry, crash reporting, or usage tracking.
- No document text is sent to the author or any third party.
- No data is stored outside the active Word document.

## Repository and GitHub

The source code is hosted on GitHub. Browsing the repository, opening issues,
or downloading files is subject to GitHub's own privacy policy. The macro
itself does not contact GitHub.

## Issues and sample documents

If you report a bug, do not attach real study notes, personal documents, or
other sensitive content. Use a minimal synthetic example and redact any
personal information.

## Contact

For privacy questions, open a non-sensitive issue or use the private security
reporting channel described in [SECURITY.md](SECURITY.md).

Last updated: 2026-10-31.
