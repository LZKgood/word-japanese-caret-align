# Security Policy

## Supported versions

| Version | Supported |
|---|---|
| 1.0.x | ✅ |
| < 1.0 | ❌ (development builds, known defects) |

## Reporting a vulnerability

Please **do not** open a public issue for a security-sensitive report. Use
GitHub's private vulnerability reporting for this repository:

1. Open the repository's **Security** tab.
2. Choose **Advisories** → **Report a vulnerability**.
3. Include the macro version, Word version, Windows version, a description,
   reproduction steps, and any relevant VBA line numbers.
4. Do not include real personal documents or sensitive study notes. Use a
   minimal synthetic example.

If private advisories are not enabled, contact the maintainers through the
repository's listed contact method. Never post exploit details or sensitive
documents publicly before a fix is available.

## Scope

This project is a local Word VBA macro. It:

- makes no network requests;
- collects no telemetry;
- modifies only the active document (full-width spaces, paragraph formatting,
  Overtype mode, status bar).

Security-relevant issues include, but are not limited to:

- unintended file/document modification beyond the documented behavior;
- macro code that executes external commands, accesses the network, or writes
  outside the active document;
- code injection through document content, fields, bookmarks, or templates;
- supply-chain issues in the repository's scripts or GitHub Actions.

Out of scope:

- Word or Windows vulnerabilities (report them to Microsoft);
- the general risk of enabling macros in Word;
- data loss caused by user edits, lack of backups, or misuse;
- WPS/Word for Mac compatibility issues that are not security bugs.

## Macro safety

Only run macros you trust and have reviewed. The source is open; inspect it,
test with a backup copy, and follow your organization's macro policy. The
project does not ask you to enable "Trust access to the VBA project object
model" for normal manual installation and use.

## Disclosure

The maintainers will acknowledge a report as soon as practical, investigate,
and coordinate a fix and disclosure. Please allow reasonable time before
public disclosure.

Last updated: 2026-10-31.
