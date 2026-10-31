![Banner](banner.png)

# dso-spreadsheet-sanitizer
Removes macros and hidden sheets from spreadsheet files (e.g., xlsx, ods, csv) to prevent malicious code execution and data exfiltration. - Focused on Tools for sanitizing and obfuscating sensitive data within text files and structured data formats

## Install
`git clone https://github.com/ShadowStrikeHQ/dso-spreadsheet-sanitizer`

## Usage
`./dso-spreadsheet-sanitizer [params]`

## Parameters
- `--remove-macros`: Remove VBA macros from XLSX files: the VBA project part is deleted, its content type override is removed, the workbook content type is changed to the non macro spreadsheet type, and the VBA project relationship is removed. Macro removal is not supported for ODS files.
- `--remove-hidden-sheets`: Remove hidden sheets (XLSX) or hidden tables (ODS), including the sheet part, its relationship, and its content type override. If a hidden sheet cannot be removed, the run fails at exit 1.
- `--overwrite`: Overwrite the output file if it exists. This consents to replacing an existing output file only; it does not permit in-place work. An output path that resolves to the input file is refused at exit 1, even with `--overwrite`, and the input is read fully into memory before any output is opened.

A run that removes or changes nothing (no flags set, nothing removed, the ODS macro case, or a CSV with no formula cells) exits 1 with a plain line and no Successfully sanitized line. CSV handling keeps every row, including rows with empty cells, and neutralises cells beginning with =, +, or @ by prefixing a single quote.

## License
Copyright (c) ShadowStrikeHQ
