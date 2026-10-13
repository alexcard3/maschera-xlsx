# Changelog

All notable changes are listed here. The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/)
and versions follow [Semantic Versioning](https://semver.org/).
Le modifiche importanti sono elencate qui, in inglese.

## [0.9.2] - 2026-10-13

### Added
- ⚠ "Version in development" note in the header: suggestions can be wrong, check the result; link to this
  repository (the page's only link, opened only by a click, no data). Themes can remove it with
  `"avviso_sviluppo": false`.

## [0.9.1] - 2026-10-13

### Added
- Experiment `experiments/rizzo-pii`: rizzo-pii as a local name-suggestion engine for Maschera, on invented data,
  with the prediction committed before the run.
- README: pointer to rizzo-pii for PDF and text documents.

### Fixed
- Single capitalised words under a header that is not about names (for example cities under «Città», ledger
  reasons under «Causale») or in the header row are no longer proposed as names. Block layouts without a real
  header keep the previous behaviour.

## [0.9.0] - 2026-10-13

First public release (beta).

### Added
- Single-file HTML tool that runs offline in the browser, with a Content-Security-Policy that blocks network access
  from the code.
- Preview of every sheet, including hidden ones, with likely names of people and companies already marked.
- Marking by category with clicks, mouse drag, whole rows and columns, keyboard, or names typed by hand.
- Automatic rules for Italian tax codes, VAT numbers, IBANs, e-mails and phone numbers.
- Anonymization inside the original `.xlsx` package: formatting, formulas, charts, pivots, comments, data
  validation, print headers and number formats are kept. Author, Office user names, folder paths with a user name,
  macros and printer settings are removed.
- Final check before download: real values, personal paths and unreadable parts block the formatted output, with a
  values-only fallback.
- Restore of the real names in text answers and in spreadsheets modified by an AI assistant (rows, columns, new
  sheets, comments, charts); codes written in a different form are reported.
- Stable codes across files by loading previous correspondences.
- Themes for `build.py`: logo, colours, font and texts.
- Release workflow with SHA-256 checksums and build attestation.

[0.9.2]: https://github.com/alexcard3/maschera-xlsx/releases/tag/v0.9.2
[0.9.1]: https://github.com/alexcard3/maschera-xlsx/releases/tag/v0.9.1
[0.9.0]: https://github.com/alexcard3/maschera-xlsx/releases/tag/v0.9.0
