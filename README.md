# Right Click XLSX2CSV2XLSX

Two-way CSV and XLSX conversion from the Windows right-click menu.

One PowerShell file, about 53 KB, and no dependencies: no Excel, no Python, no
Node, no PowerShell module. An XLSX file is only a ZIP archive of XML documents,
so the script reads and writes the OpenXML package directly.

## Menu entries

| File type | Entry | Result |
| --- | --- | --- |
| `.csv`, `.tsv` | Convert to Excel (.xlsx) | A formatted workbook |
| `.xlsx` | Convert to CSV (;) | Semicolon-separated CSV |
| `.xlsx` | Convert to CSV (,) | Comma-separated CSV |

## Install

Download `RightClickXlsx2Csv2Xlsx.ps1`, put it wherever you want to keep it, and
run:

```powershell
powershell -ExecutionPolicy Bypass -File .\RightClickXlsx2Csv2Xlsx.ps1 -Install
```

The installer asks which language the menu entries should use. English and French
are available, and English is the default. Pass `-Language en` or `-Language fr`
to skip the prompt, which is what you want for an unattended install.

Everything is written under `HKCU`, so no administrator rights are needed and no
other user account is touched.

To remove everything:

```powershell
powershell -ExecutionPolicy Bypass -File .\RightClickXlsx2Csv2Xlsx.ps1 -Uninstall
```

Installing writes one companion file next to the script,
`RightClickXlsx2Csv2Xlsx.launcher.vbs`. It exists only so Explorer does not flash
a console window on every conversion, it is about 30 lines, and `-Uninstall`
deletes it. Do not move the folder after installing, because the registry points
at both paths. If you do move it, run `-Install` again.

Macro-enabled workbooks are not registered by default. Add them with
`-XlsxExtensions .xlsx,.xlsm`.

On the command line the converter accepts more than the menu registers: `.xlsm`
reads like `.xlsx`, and `.txt` is treated as a CSV. Those two are deliberately
left out of the registry, because claiming file types people use for other
things is not this tool's business.

## Use

Right-click one or more files and pick an entry. On Windows 11 these live under
**Show more options**, the inherited menu, which `Shift+F10` opens directly. The
short Windows 11 menu only accepts signed MSIX packages, which would mean an
installer and a code-signing certificate, the opposite of the goal here.

Output lands next to the source file. An existing file of the same name is never
overwritten: a numbered suffix is added instead.

`-Delimiter` is not sent by the menu when converting a CSV, so the separator of
the file being read is always detected. The two CSV entries send it only because
there the separator is a choice about the file being written.

## Command line

The direction comes from the file extension, so there is nothing to specify:

```powershell
.\RightClickXlsx2Csv2Xlsx.ps1 sales.csv
.\RightClickXlsx2Csv2Xlsx.ps1 report.xlsx
.\RightClickXlsx2Csv2Xlsx.ps1 *.csv -Force
.\RightClickXlsx2Csv2Xlsx.ps1 report.xlsx -Delimiter "," -Sheet "Q3 results"
```

One object is emitted per converted file, so the script composes in a pipeline.

| Parameter | Effect |
| --- | --- |
| `-Delimiter` | The separator on the CSV side, either way. Overrides detection when reading, sets the output separator when writing, defaulting to a semicolon |
| `-Encoding` | `auto`, `utf8`, `utf8bom`, `ansi`, `unicode`. Reading a CSV, `auto` detects. Writing one, `auto` means UTF-8 with a byte order mark |
| `-OutFile` | Explicit output path, single file and single sheet. Needs `-Force` to replace an existing file, and never writes onto the source |
| `-Force` | Overwrite existing output |
| `-Open` | Open the result when done |
| `-Language` | `en` or `fr`, for messages and boolean words |
| `-Version` | Print the tool name and version, then exit |
| `-AsText` | CSV to XLSX: write everything as text, no number or date conversion |
| `-NoHeader` | CSV to XLSX: no header formatting |
| `-Sheet` | XLSX to CSV: sheet name or 1-based index |
| `-DecimalSeparator` | XLSX to CSV: `auto`, `dot` or `comma` |
| `-Raw` | XLSX to CSV: keep stored values, leaving dates as serial numbers |

A workbook with one sheet produces one CSV. A workbook with several sheets
produces one CSV per sheet, each suffixed with the sheet name, so nothing is
silently dropped.

## What the conversion does

### CSV to XLSX

**Separator detection** counts `;`, `,`, tab and `|` outside quotes on the first
non-empty line. The Excel `sep=;` convention on a leading line is honoured and
that line is consumed.

**Encoding detection** looks for a UTF-8, UTF-16 LE or UTF-16 BE byte order mark.
Without one, a sample is decoded as strict UTF-8; if that succeeds the file is
treated as UTF-8, otherwise as the system ANSI code page.

**CSV parsing** is delegated to the .NET `TextFieldParser`, which handles quoted
fields, doubled quotes, and separators or line breaks inside a field.

**Cell typing** turns integers and decimals into numbers. A decimal comma is
accepted when the comma is not the field separator. Values with a leading zero
such as `0012345` stay text, which preserves reference numbers, postal codes and
phone numbers. Past 15 digits a value stays text, because a double no longer
holds it exactly.

**Dates** are recognised in ISO `yyyy-MM-dd` and `yyyy/MM/dd` form, plus the short
date pattern of the current Windows locale, with or without a time. Dates before
1 March 1900 stay text, since Excel shifts serial numbers before that point.

**Formatting** gives the header row bold text, a frozen pane, an autofilter, and
column widths fitted to the content between 8 and 60 characters.

### XLSX to CSV

Shared strings, inline strings, cached formula results, booleans and error values
such as `#DIV/0!` are all resolved to what Excel displays.

**Dates** are the interesting part. A numeric cell is only a date because its
number format says so, so the style table is read and both built-in and custom
format codes are classified as date, time or both. Matching serial numbers are
rendered with the short date pattern of the current locale. Excel counts a day
that never existed, 29 February 1900, and that offset is reproduced, so output
matches Excel cell for cell.

**Decimal separator** follows Excel's own behaviour under `auto`: a comma when the
field separator is a semicolon and the locale uses a decimal comma, a dot
otherwise.

**Blank rows** inside the used range are preserved, so row numbers still line up
with the original sheet.

Output is UTF-8 with a byte order mark by default, which is what makes Excel
reopen the file with the right encoding.

## Technical notes

Both directions stream, in two passes over the input. The first pass measures
dimensions, the second writes. Memory use therefore does not depend on file size.
Output goes to a temporary file that is renamed at the end, so an interrupted run
never leaves a truncated result behind.

Writing uses `inlineStr` rather than a shared string table. Files are slightly
larger, but writing stays streaming and stateless.

Control characters XML forbids are stripped, and `<`, `>`, `&` and quotes are
escaped. On the way back, XML entities and Excel's `_xHHHH_` escapes are decoded.
A line whose quotes are never closed is kept verbatim in a single text cell, so no
data is lost.

The reader's inner loop picks cell attributes apart with `IndexOf` instead of
regexes or helper calls. That is what takes a 100,000-row workbook from 195
seconds down to 18.

The launcher hands file paths to PowerShell through environment variables rather
than on the command line. `WshShell.Run` expands `%NAME%` inside the command
string it is given, so a file called `Q1 %USERNAME% report.csv` would otherwise
arrive with the percent section replaced and fail as not found. Environment
variables are inherited verbatim, the command string contains no percent sign at
all, and no caller-supplied text ever reaches a parser.

Measured on a Windows 11 laptop, 100,000 rows by 8 columns:

| Direction | Input | Time | Output |
| --- | --- | --- | --- |
| CSV to XLSX | 7.9 MB | 5.9 s | 3.7 MB |
| XLSX to CSV | 3.7 MB | 18.0 s | 7.8 MB |

A real 1,923-row by 23-column export round-trips through both directions in about
a second with all 44,229 cells identical.

## Known limits

Writing produces a single sheet, named after the CSV file, with no formulas,
conditional formatting or structured tables. Reading covers XLSX only; XLSB and
XLS are binary formats and are rejected with a clear message.

A CSV line whose quotes are never closed swallows the following lines, which is
inherent to the format.

The short date pattern follows the Windows locale in both directions. A CSV
written as `MM/dd/yyyy` and read on a machine set to `dd/MM/yyyy` will be
misread, so use `-AsText` or ISO dates in that case.

## Changelog

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and
this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
`-Version` prints the version the script reports.

### 1.0.0 - 2026-09-11

Initial release.

- Two-way conversion between CSV and XLSX, in a single PowerShell file with no
  external dependencies. The OpenXML package is read and written directly.
- Right-click menu entries: Convert to Excel (.xlsx) on `.csv` and `.tsv`,
  Convert to CSV (;) and Convert to CSV (,) on `.xlsx`.
- `-Install` and `-Uninstall`, writing only under `HKCU` so no administrator
  rights are needed. Uninstall removes every verb it finds, whatever extensions
  were registered.
- Menu and message language selectable at install time, English or French.
- Separator detection over a whole logical CSV record, covering `;`, `,`, tab
  and `|`, plus the Excel `sep=` convention.
- Encoding detection by byte order mark, falling back to a strict UTF-8 test and
  then to the system ANSI code page.
- Number and date typing on the way in, with leading zeros, long digit strings
  and pre-1900 dates kept as text.
- Date, time and elapsed-time number formats resolved on the way out, including
  Excel's 29 February 1900 offset and the 1904 calendar.
- Shared strings, inline strings, cached formula results, booleans and error
  values resolved to what Excel displays.
- One CSV per sheet for multi-sheet workbooks, so nothing is dropped silently.
- Streaming in both directions, in two passes, so memory use does not track file
  size. Output is written to a temporary file and renamed.

## Requirements

Windows with PowerShell 5.1, which ships with Windows 10 and 11. Nothing else.

## Licence

GNU Affero General Public License, version 3 or later. The full text is in
[LICENSE](LICENSE).

In short: you may use, study, modify and redistribute this tool freely, and any
modified version you distribute, or make available to users over a network, must
be offered under the same licence with its source. It comes with absolutely no
warranty.

`SPDX-License-Identifier: AGPL-3.0-or-later`
