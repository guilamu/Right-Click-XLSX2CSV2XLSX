# Right Click XLSX2CSV2XLSX - two-way CSV and XLSX conversion for Windows.
# Copyright (C) 2026 Guillaume Taurel
#
# This program is free software: you can redistribute it and/or modify it under
# the terms of the GNU Affero General Public License as published by the Free
# Software Foundation, either version 3 of the License, or (at your option) any
# later version.
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU Affero General Public License for more
# details.
#
# You should have received a copy of the GNU Affero General Public License along
# with this program. If not, see <https://www.gnu.org/licenses/>.
#
# SPDX-License-Identifier: AGPL-3.0-or-later

<#
.SYNOPSIS
    Right Click XLSX2CSV2XLSX - two-way CSV and XLSX conversion, in one file,
    with no external dependencies.

.DESCRIPTION
    An XLSX file is only a ZIP archive of XML documents, so this script reads and
    writes the OpenXML package directly. It needs no Excel, no Python and no
    PowerShell module. Both directions stream, so memory use does not grow with
    file size.

    The direction is taken from the file extension: .csv and .tsv become .xlsx,
    and .xlsx becomes .csv.

    Run it with -Install to add the entries to the Windows right-click menu.

.PARAMETER Path
    One or more files. Wildcards are accepted.

.PARAMETER Install
    Adds the right-click menu entries and writes the launcher next to this
    script. Everything goes under HKCU, so no administrator rights are needed.

.PARAMETER Uninstall
    Removes the menu entries and deletes the generated launcher.

.PARAMETER Language
    Language of menu entries and messages: en or fr. When -Install is used
    without it, an interactive session asks, defaulting to English.

.PARAMETER Delimiter
    The separator on the CSV side, whichever way the conversion runs. Reading a
    CSV, it overrides detection. Writing a CSV, it sets the output separator and
    defaults to a semicolon.

.PARAMETER Encoding
    auto, utf8, utf8bom, ansi or unicode. Reading a CSV, auto detects the
    encoding. Writing a CSV, auto means UTF-8 with a byte order mark, which is
    what makes Excel reopen the file correctly.

.PARAMETER OutFile
    Explicit output path. Valid for a single source file and a single sheet.

.PARAMETER AsText
    CSV to XLSX only. Writes every cell as text, with no number or date
    conversion.

.PARAMETER NoHeader
    CSV to XLSX only. Skips header formatting: bold, autofilter, frozen pane.

.PARAMETER Sheet
    XLSX to CSV only. Sheet to export, as a name or a 1-based index. When
    omitted, a multi-sheet workbook produces one CSV per sheet.

.PARAMETER DecimalSeparator
    XLSX to CSV only. auto, dot or comma. auto uses a comma when the field
    separator is a semicolon and the current culture uses a decimal comma, which
    is what Excel itself does.

.PARAMETER Raw
    XLSX to CSV only. Writes stored values as they are, leaving dates as serial
    numbers.

.PARAMETER Force
    Overwrites an existing output file instead of picking a new name.

.PARAMETER Open
    Opens each converted file when done.

.PARAMETER Gui
    Reports failures in a message box instead of on the console. Used by the
    launcher, which runs without a visible window.

.PARAMETER Version
    Prints the tool name and version, then exits.

.EXAMPLE
    .\RightClickXlsx2Csv2Xlsx.ps1 -Install

.EXAMPLE
    .\RightClickXlsx2Csv2Xlsx.ps1 sales.csv

.EXAMPLE
    .\RightClickXlsx2Csv2Xlsx.ps1 report.xlsx -Delimiter ","

.EXAMPLE
    .\RightClickXlsx2Csv2Xlsx.ps1 *.csv -Force
#>
[CmdletBinding()]
param(
    # ValueFromRemainingArguments is what lets Explorer hand over several bare
    # paths at once. The parameter is object[] rather than string[] on purpose:
    # with string[], PowerShell 5.1 flattens an array passed as a single
    # argument into one space-joined string. Both shapes are normalised below.
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [object[]] $Path,

    [switch] $Install,
    [switch] $Uninstall,

    [ValidateSet('en', 'fr')]
    [string] $Language,

    [string] $Delimiter,

    [ValidateSet('auto', 'utf8', 'utf8bom', 'ansi', 'unicode')]
    [string] $Encoding = 'auto',

    [string] $OutFile,

    [switch] $AsText,
    [switch] $NoHeader,

    [string] $Sheet,

    [ValidateSet('auto', 'dot', 'comma')]
    [string] $DecimalSeparator = 'auto',

    [switch] $Raw,
    [switch] $Force,
    [switch] $Open,
    [switch] $Gui,
    [switch] $Version,

    # Set by the generated launcher: take the file list from the environment
    # rather than from the command line. See Write-Launcher for why.
    [switch] $FromLauncher,

    [string[]] $CsvExtensions  = @('.csv', '.tsv'),
    [string[]] $XlsxExtensions = @('.xlsx')
)

$ErrorActionPreference = 'Stop'

$ToolName = 'Right Click XLSX2CSV2XLSX'

# Keep in step with the changelog in README.md.
$ToolVersion = '1.0.0'

# ---------------------------------------------------------------------------
# User-facing strings. Code and comments stay in English.
# ---------------------------------------------------------------------------
$Strings = @{
    en = @{
        MenuToXlsx  = 'Convert to Excel (.xlsx)'
        MenuToCsvSc = 'Convert to CSV (;)'
        MenuToCsvC  = 'Convert to CSV (,)'
        Usage       = 'Usage: RightClickXlsx2Csv2Xlsx.ps1 <file.csv|file.xlsx> [...]   |   -Install   |   -Uninstall'
        Licence     = 'Free software under the GNU AGPL v3, with absolutely no warranty. See the LICENSE file.'
        NoAssembly  = 'Could not load the required .NET assemblies: {0}'
        NotFound    = 'File not found.'
        Unknown     = "Unsupported file type '{0}'. Expected {1} or {2}."
        SingleOut   = '-OutFile can only be used with a single source file and a single sheet.'
        NoFreeName  = "Could not find a free output name for '{0}'."
        OutExists   = "'{0}' already exists. Use -Force to overwrite it."
        OutIsSource = 'The output path is the source file itself. Refusing to overwrite it.'
        NotXlsx     = 'Not a readable XLSX package (no workbook part found).'
        NoSheet     = "Sheet '{0}' was not found in the workbook."
        NoSheets    = 'The workbook contains no worksheet.'
        Xlsb        = 'XLSB and XLS are binary formats and are not supported. Save as XLSX first.'
        DlgTitle    = 'Right Click XLSX2CSV2XLSX'
        DlgFailed   = 'Conversion failed for {0} file(s):'
        Installed   = 'Installed: {0}'
        Removed     = 'Removed: {0}'
        DoneInstall = 'Done. Right-click a file, then "Show more options" on Windows 11 (or press Shift+F10).'
        DoneRemove  = 'Context menu entries removed.'
        AskLang     = 'Menu language:'
        AskEn       = '  [1] English (default)'
        AskFr       = '  [2] French / Francais'
        AskPrompt   = 'Choice'
        WroteLauncher = 'Launcher written: {0}'
    }
    fr = @{
        MenuToXlsx  = 'Convertir en Excel (.xlsx)'
        MenuToCsvSc = 'Convertir en CSV (;)'
        MenuToCsvC  = 'Convertir en CSV (,)'
        Usage       = 'Utilisation : RightClickXlsx2Csv2Xlsx.ps1 <fichier.csv|fichier.xlsx> [...]   |   -Install   |   -Uninstall'
        Licence     = 'Logiciel libre sous licence GNU AGPL v3, sans aucune garantie. Voir le fichier LICENSE.'
        NoAssembly  = 'Impossible de charger les assemblies .NET requises : {0}'
        NotFound    = 'Fichier introuvable.'
        Unknown     = "Type de fichier non pris en charge '{0}'. Attendu {1} ou {2}."
        SingleOut   = "-OutFile ne peut etre utilise qu'avec un seul fichier source et une seule feuille."
        NoFreeName  = "Impossible de trouver un nom de sortie libre pour '{0}'."
        OutExists   = "'{0}' existe deja. Utilisez -Force pour l'ecraser."
        OutIsSource = "Le chemin de sortie est le fichier source lui-meme. Ecrasement refuse."
        NotXlsx     = "Ce fichier n'est pas un paquet XLSX lisible (partie workbook absente)."
        NoSheet     = "La feuille '{0}' est introuvable dans le classeur."
        NoSheets    = 'Le classeur ne contient aucune feuille de calcul.'
        Xlsb        = 'Les fichiers XLSB et XLS sont des formats binaires non pris en charge. Enregistrez-les en XLSX.'
        DlgTitle    = 'Right Click XLSX2CSV2XLSX'
        DlgFailed   = 'La conversion a echoue pour {0} fichier(s) :'
        Installed   = 'Installe : {0}'
        Removed     = 'Retire : {0}'
        DoneInstall = 'Termine. Clic droit sur un fichier, puis "Afficher plus d''options" sous Windows 11 (ou Maj+F10).'
        DoneRemove  = 'Entrees du menu contextuel supprimees.'
        AskLang     = 'Langue du menu :'
        AskEn       = '  [1] Anglais / English (defaut)'
        AskFr       = '  [2] Francais'
        AskPrompt   = 'Choix'
        WroteLauncher = 'Lanceur ecrit : {0}'
    }
}

$EffectiveLanguage = $Language
if (-not $EffectiveLanguage) { $EffectiveLanguage = 'en' }
$L = $Strings[$EffectiveLanguage]

try {
    Add-Type -AssemblyName Microsoft.VisualBasic
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
} catch {
    throw ($L.NoAssembly -f $_.Exception.Message)
}

$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$Cul = [System.Globalization.CultureInfo]::CurrentCulture

# What the converter accepts on the command line. These are wider than the
# extensions registered in the right-click menu: .xlsm reads like .xlsx, and a
# .txt export is just a CSV with another name. Registering them by default
# would hijack file types people use for other things, so the menu sticks to
# $CsvExtensions and $XlsxExtensions.
$ReadableXlsx = @($XlsxExtensions + '.xlsm' | Select-Object -Unique)
$ReadableCsv  = @($CsvExtensions  + '.txt'  | Select-Object -Unique)

# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------
function Get-FreePath([string] $dir, [string] $base, [string] $ext) {
    $candidate = [System.IO.Path]::Combine($dir, $base + $ext)
    if ($Force -or -not (Test-Path -LiteralPath $candidate)) { return $candidate }
    for ($i = 1; $i -lt 1000; $i++) {
        $candidate = [System.IO.Path]::Combine($dir, $base + ' (' + $i + ')' + $ext)
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    throw ($L.NoFreeName -f $base)
}

# -OutFile names the destination outright, so it must not become a way around
# the overwrite rules that Get-FreePath enforces everywhere else. Writing onto
# the source is refused even with -Force: the conversion would replace the file
# being read with a different format and lose the original.
function Resolve-OutFilePath([string] $source, [string] $requested) {
    $dest = [System.IO.Path]::GetFullPath($requested)
    $src = [System.IO.Path]::GetFullPath($source)
    if ([string]::Equals($src, $dest, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw $L.OutIsSource
    }
    if ((-not $Force) -and (Test-Path -LiteralPath $dest)) {
        throw ($L.OutExists -f $dest)
    }
    return $dest
}

function ConvertTo-SafeFileName([string] $name) {
    foreach ($bad in [System.IO.Path]::GetInvalidFileNameChars()) { $name = $name.Replace($bad, '_') }
    return $name.Trim()
}

# ===========================================================================
#  CSV  ->  XLSX
# ===========================================================================

$RxInt   = [regex] '^[+-]?[0-9]{1,15}$'
$RxDec   = [regex] '^[+-]?(?:[0-9]{1,15}\.[0-9]{1,15}|\.[0-9]{1,15}|[0-9]{1,15}\.)$'
$RxDecC  = [regex] '^[+-]?[0-9]{1,15},[0-9]{1,15}$'
$RxLead0 = [regex] '^[+-]?0[0-9]'
$RxCtrl  = [regex] '[\x00-\x08\x0B\x0C\x0E-\x1F]'
$MinDate = [datetime] '1900-03-01'

$sdp = $Cul.DateTimeFormat.ShortDatePattern
$DateFmts = [string[]] @('yyyy-MM-dd', 'yyyy/MM/dd', $sdp)
$TimeFmts = [string[]] @(
    'yyyy-MM-dd HH:mm', 'yyyy-MM-dd HH:mm:ss',
    'yyyy-MM-ddTHH:mm', 'yyyy-MM-ddTHH:mm:ss',
    ($sdp + ' HH:mm'), ($sdp + ' HH:mm:ss')
)

function Get-ColumnName([int] $index) {
    $n = ''
    while ($index -gt 0) {
        $r = ($index - 1) % 26
        $n = [string] [char] (65 + $r) + $n
        $index = [int] (($index - $r - 1) / 26)
    }
    return $n
}

function Resolve-CsvEncoding([string] $file, [string] $pref) {
    switch ($pref) {
        'utf8'    { return (New-Object System.Text.UTF8Encoding($false)) }
        'utf8bom' { return (New-Object System.Text.UTF8Encoding($true)) }
        'ansi'    { return [System.Text.Encoding]::Default }
        'unicode' { return [System.Text.Encoding]::Unicode }
    }

    $fs = [System.IO.File]::OpenRead($file)
    try {
        $head = New-Object byte[] 4
        $n = $fs.Read($head, 0, 4)
        if ($n -ge 3 -and $head[0] -eq 0xEF -and $head[1] -eq 0xBB -and $head[2] -eq 0xBF) {
            return (New-Object System.Text.UTF8Encoding($true))
        }
        if ($n -ge 2 -and $head[0] -eq 0xFF -and $head[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode }
        if ($n -ge 2 -and $head[0] -eq 0xFE -and $head[1] -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode }

        # No BOM: test whether a sample decodes as strict UTF-8.
        $fs.Position = 0
        $size = [int] [Math]::Min($fs.Length, 4MB)
        if ($size -eq 0) { return (New-Object System.Text.UTF8Encoding($false)) }
        $sample = New-Object byte[] $size
        $read = $fs.Read($sample, 0, $size)

        # Drop a multi-byte sequence cut in half at the end of the sample.
        $end = $read
        if ($read -lt $fs.Length) {
            $k = 0
            while ($end -gt 0 -and $k -lt 4 -and ($sample[$end - 1] -band 0x80) -ne 0) { $end--; $k++ }
        }
        try {
            $strict = New-Object System.Text.UTF8Encoding($false, $true)
            [void] $strict.GetString($sample, 0, $end)
            return (New-Object System.Text.UTF8Encoding($false))
        } catch {
            return [System.Text.Encoding]::Default
        }
    } finally {
        $fs.Dispose()
    }
}

function Resolve-CsvDelimiter([string] $file, $enc, [ref] $skipFirstLine) {
    # Sniffing must see a whole logical record, not one physical line. A first
    # field like "Title<newline>continued" would otherwise hide every separator
    # on the line that follows it, and detection would fall back to the locale.
    $line = $null
    $sr = New-Object System.IO.StreamReader($file, $enc, $true)
    try {
        $sb = New-Object System.Text.StringBuilder 512
        $firstPhysical = $null
        $inQuote = $false
        for ($i = 0; $i -lt 64; $i++) {
            $l = $sr.ReadLine()
            if ($null -eq $l) { break }
            if (($sb.Length -eq 0) -and (-not $inQuote) -and ($l.Trim().Length -eq 0)) { continue }
            if ($null -eq $firstPhysical) { $firstPhysical = $l }
            if ($sb.Length -gt 0) { [void] $sb.Append("`n") }
            [void] $sb.Append($l)
            foreach ($ch in $l.ToCharArray()) { if ($ch -eq '"') { $inQuote = -not $inQuote } }
            if (-not $inQuote) { break }
            if ($sb.Length -gt 65536) { break }
        }
        if ($sb.Length -gt 0) {
            # Quotes that never close mean the file is malformed rather than
            # multi-line. Fall back to the first physical line in that case.
            if ($inQuote) { $line = $firstPhysical } else { $line = $sb.ToString() }
        }
    } finally { $sr.Dispose() }

    if ($null -eq $line) { return ',' }

    # Excel convention: a leading "sep=;" line sets the separator.
    if ($line -match '^sep=(.)\s*$') {
        $skipFirstLine.Value = $true
        return $Matches[1]
    }

    $best = $null
    $bestCount = 0
    foreach ($cand in @(';', ',', "`t", '|')) {
        $c = [char] $cand
        $count = 0
        $inQuote = $false
        foreach ($ch in $line.ToCharArray()) {
            if ($ch -eq '"') { $inQuote = -not $inQuote }
            elseif ((-not $inQuote) -and ($ch -eq $c)) { $count++ }
        }
        if ($count -gt $bestCount) { $bestCount = $count; $best = $cand }
    }
    if ($null -eq $best) { return $Cul.TextInfo.ListSeparator }
    return $best
}

function New-CsvParser([string] $file, $enc, [string] $delim, [bool] $skipFirstLine) {
    $sr = New-Object System.IO.StreamReader($file, $enc, $true)
    if ($skipFirstLine) { [void] $sr.ReadLine() }
    $p = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($sr)
    $p.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $p.SetDelimiters([string[]] @($delim))
    $p.HasFieldsEnclosedInQuotes = $true
    $p.TrimWhiteSpace = $false
    return $p
}

function Add-ZipTextEntry($zip, [string] $name, [string] $content) {
    $entry = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $w = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        $w.Write($content)
        $w.Flush()
        $w.Dispose()
    } finally { $stream.Dispose() }
}

function Get-SheetName([string] $baseName) {
    $n = $baseName
    foreach ($bad in @('[', ']', ':', '*', '?', '/', '\')) { $n = $n.Replace($bad, '_') }
    $n = $n.Trim([char] 39).Trim()
    if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
    if ($n.Length -eq 0) { $n = 'Sheet1' }
    return $n
}

function Convert-CsvToXlsx([string] $source, [string] $destination) {

    $enc = Resolve-CsvEncoding $source $Encoding
    $skipFirst = $false
    $sniffed = Resolve-CsvDelimiter $source $enc ([ref] $skipFirst)
    $delim = $Delimiter
    if ([string]::IsNullOrEmpty($delim)) { $delim = $sniffed }
    $allowDecComma = ($delim -ne ',')

    # --- Pass 1: dimensions and column widths ------------------------------
    $rowCount = 0
    $colCount = 0
    $widths = New-Object 'System.Collections.Generic.List[int]'
    $p = New-CsvParser $source $enc $delim $skipFirst
    try {
        while (-not $p.EndOfData) {
            try { $fields = $p.ReadFields() }
            catch [Microsoft.VisualBasic.FileIO.MalformedLineException] { $fields = @($p.ErrorLine) }
            if ($null -eq $fields) { continue }
            $rowCount++
            if ($fields.Count -gt $colCount) { $colCount = $fields.Count }
            for ($i = 0; $i -lt $fields.Count; $i++) {
                while ($widths.Count -le $i) { $widths.Add(0) }
                $len = $fields[$i].Length
                if ($len -gt $widths[$i]) { $widths[$i] = $len }
            }
        }
    } finally { $p.Dispose() }

    $hasHeader = (-not $NoHeader) -and ($rowCount -ge 2)
    $span = [Math]::Max($colCount, 1)
    $colNames = New-Object string[] $span
    for ($i = 0; $i -lt $span; $i++) { $colNames[$i] = Get-ColumnName ($i + 1) }
    $lastCol = $colNames[$span - 1]

    # --- Static parts of the OpenXML package -------------------------------
    $sheetName = [System.Security.SecurityElement]::Escape(
        (Get-SheetName ([System.IO.Path]::GetFileNameWithoutExtension($source))))

    $contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'

    $rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

    $wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'

    $styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="22" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'

    $workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="' + $sheetName + '" sheetId="1" r:id="rId1"/></sheets></workbook>'

    # --- Worksheet head and tail -------------------------------------------
    $head = New-Object System.Text.StringBuilder 4096
    [void] $head.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void] $head.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    if ($rowCount -gt 0) {
        [void] $head.Append('<dimension ref="A1:').Append($lastCol).Append($rowCount).Append('"/>')
    }
    [void] $head.Append('<sheetViews><sheetView tabSelected="1" workbookViewId="0">')
    if ($hasHeader) {
        [void] $head.Append('<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>')
    }
    [void] $head.Append('</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>')
    if ($colCount -gt 0) {
        [void] $head.Append('<cols>')
        for ($i = 0; $i -lt $colCount; $i++) {
            $w = $widths[$i] + 2
            if ($w -lt 8) { $w = 8 }
            if ($w -gt 60) { $w = 60 }
            [void] $head.Append('<col min="').Append($i + 1).Append('" max="').Append($i + 1).Append('" width="').Append($w).Append('" customWidth="1"/>')
        }
        [void] $head.Append('</cols>')
    }
    [void] $head.Append('<sheetData>')

    $tail = New-Object System.Text.StringBuilder 256
    [void] $tail.Append('</sheetData>')
    if ($hasHeader) {
        [void] $tail.Append('<autoFilter ref="A1:').Append($lastCol).Append($rowCount).Append('"/>')
    }
    [void] $tail.Append('</worksheet>')

    # --- Write the package --------------------------------------------------
    $destDir = [System.IO.Path]::GetDirectoryName($destination)
    $tmp = [System.IO.Path]::Combine($destDir, '~' + [System.IO.Path]::GetFileName($destination) + '.tmp')
    $fs = [System.IO.File]::Open($tmp, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    $zip = $null
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)

        Add-ZipTextEntry $zip '[Content_Types].xml'        $contentTypes
        Add-ZipTextEntry $zip '_rels/.rels'                $rootRels
        Add-ZipTextEntry $zip 'xl/workbook.xml'            $workbook
        Add-ZipTextEntry $zip 'xl/_rels/workbook.xml.rels' $wbRels
        Add-ZipTextEntry $zip 'xl/styles.xml'              $styles

        $entry = $zip.CreateEntry('xl/worksheets/sheet1.xml', [System.IO.Compression.CompressionLevel]::Optimal)
        $es = $entry.Open()
        $sw = New-Object System.IO.StreamWriter($es, (New-Object System.Text.UTF8Encoding($false)), 65536)
        try {
            $sw.Write($head.ToString())

            # --- Pass 2: the data --------------------------------------------
            $sb = New-Object System.Text.StringBuilder 8192
            $r = 0
            $p = New-CsvParser $source $enc $delim $skipFirst
            try {
                while (-not $p.EndOfData) {
                    try { $fields = $p.ReadFields() }
                    catch [Microsoft.VisualBasic.FileIO.MalformedLineException] { $fields = @($p.ErrorLine) }
                    if ($null -eq $fields) { continue }
                    $r++
                    $isHeaderRow = ($hasHeader -and $r -eq 1)

                    [void] $sb.Clear()
                    [void] $sb.Append('<row r="').Append($r).Append('">')

                    for ($i = 0; $i -lt $fields.Count; $i++) {
                        $v = $fields[$i]
                        if ($null -eq $v -or $v.Length -eq 0) { continue }
                        $ref = $colNames[$i] + $r

                        $num = $null
                        $serial = $null
                        $dstyle = 0

                        if ((-not $isHeaderRow) -and (-not $AsText)) {
                            $t = $v.Trim()
                            if ($t.Length -gt 0) {
                                $c0 = $t[0]
                                if ([char]::IsDigit($c0) -or $c0 -eq '-' -or $c0 -eq '+' -or $c0 -eq '.') {
                                    if (-not $RxLead0.IsMatch($t)) {
                                        $d = 0.0
                                        if ($RxInt.IsMatch($t) -or $RxDec.IsMatch($t)) {
                                            if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, $Inv, [ref] $d)) { $num = $d }
                                        } elseif ($allowDecComma -and $RxDecC.IsMatch($t)) {
                                            if ([double]::TryParse($t.Replace(',', '.'), [System.Globalization.NumberStyles]::Float, $Inv, [ref] $d)) { $num = $d }
                                        }
                                    }
                                    if (($null -eq $num) -and [char]::IsDigit($c0)) {
                                        $dt = [datetime]::MinValue
                                        if ([datetime]::TryParseExact($t, $DateFmts, $Cul, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) {
                                            if ($dt -ge $MinDate) { $serial = $dt.ToOADate(); $dstyle = 2 }
                                        } elseif ([datetime]::TryParseExact($t, $TimeFmts, $Cul, [System.Globalization.DateTimeStyles]::None, [ref] $dt)) {
                                            if ($dt -ge $MinDate) { $serial = $dt.ToOADate(); $dstyle = 3 }
                                        }
                                    }
                                }
                            }
                        }

                        if ($null -ne $num) {
                            [void] $sb.Append('<c r="').Append($ref).Append('"><v>').Append($num.ToString('R', $Inv)).Append('</v></c>')
                        } elseif ($null -ne $serial) {
                            [void] $sb.Append('<c r="').Append($ref).Append('" s="').Append($dstyle).Append('"><v>').Append($serial.ToString('R', $Inv)).Append('</v></c>')
                        } else {
                            $x = $v
                            if ($RxCtrl.IsMatch($x)) { $x = $RxCtrl.Replace($x, '') }
                            $x = [System.Security.SecurityElement]::Escape($x)
                            [void] $sb.Append('<c r="').Append($ref).Append('"')
                            if ($isHeaderRow) { [void] $sb.Append(' s="1"') }
                            [void] $sb.Append(' t="inlineStr"><is><t')
                            if ([char]::IsWhiteSpace($v[0]) -or [char]::IsWhiteSpace($v[$v.Length - 1])) {
                                [void] $sb.Append(' xml:space="preserve"')
                            }
                            [void] $sb.Append('>').Append($x).Append('</t></is></c>')
                        }
                    }

                    [void] $sb.Append('</row>')
                    $sw.Write($sb.ToString())
                }
            } finally { $p.Dispose() }

            $sw.Write($tail.ToString())
            $sw.Flush()
        } finally {
            $sw.Dispose()
            $es.Dispose()
        }
    } catch {
        if ($null -ne $zip) { $zip.Dispose(); $zip = $null }
        $fs.Dispose()
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        throw
    } finally {
        if ($null -ne $zip) { $zip.Dispose() }
        $fs.Dispose()
    }

    if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Force }
    [System.IO.File]::Move($tmp, $destination)

    return [pscustomobject] @{
        Source    = $source
        Output    = $destination
        Sheet     = $sheetName
        Rows      = $rowCount
        Columns   = $colCount
        Delimiter = $delim
        Encoding  = $enc.WebName
    }
}

# ===========================================================================
#  XLSX  ->  CSV
# ===========================================================================

$RxOpts   = [System.Text.RegularExpressions.RegexOptions] 'Singleline, Compiled'
$RxCell   = [regex]::new('<c\b([^>]*?)(?:/>|>(.*?)</c>)', $RxOpts)
$RxText   = [regex]::new('<t\b[^>]*(?<!/)>(.*?)</t>', $RxOpts)
$RxVal    = [regex]::new('<v\b[^>]*(?<!/)>(.*?)</v>', $RxOpts)
$RxRPh    = [regex]::new('<rPh\b.*?</rPh>', 'Singleline')
$RxNumEnt = [regex]::new('&#(x?)([0-9A-Fa-f]+);')
$RxEscSeq = [regex]::new('_x([0-9A-Fa-f]{4})_')

# Undoes XML entity encoding plus Excel's _xHHHH_ escaping of control chars.
function Expand-XmlText([string] $s) {
    if ($s.Length -eq 0) { return $s }
    if ($s.IndexOf('&') -ge 0) {
        if ($s.IndexOf('&#') -ge 0) {
            $s = $RxNumEnt.Replace($s, {
                param($m)
                if ($m.Groups[1].Value -eq 'x') { [string][char][Convert]::ToInt32($m.Groups[2].Value, 16) }
                else { [string][char][int] $m.Groups[2].Value }
            })
        }
        $s = $s.Replace('&lt;', '<').Replace('&gt;', '>').Replace('&quot;', '"').Replace('&apos;', "'")
        $s = $s.Replace('&amp;', '&')   # must come last
    }
    if ($s.IndexOf('_x') -ge 0) {
        # _x005F_ is Excel's escape for a literal underscore introducing a
        # sequence, so it is parked aside before the real sequences are decoded.
        $park = [string][char] 1
        $s = $s.Replace('_x005F_', $park)
        $s = $RxEscSeq.Replace($s, { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) })
        $s = $s.Replace($park, '_')
    }
    return $s
}

function Get-ColumnIndex([string] $cellRef) {
    $n = 0
    foreach ($ch in $cellRef.ToCharArray()) {
        if ($ch -ge 'A' -and $ch -le 'Z') { $n = $n * 26 + ([int][char] $ch - 64) }
        elseif ($ch -ge 'a' -and $ch -le 'z') { $n = $n * 26 + ([int][char] $ch - 96) }
        else { break }
    }
    return $n
}

function New-XmlReaderOn($stream) {
    $st = New-Object System.Xml.XmlReaderSettings
    $st.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $st.IgnoreComments = $true
    $st.CloseInput = $true
    return [System.Xml.XmlReader]::Create($stream, $st)
}

function Get-ZipEntry($zip, [string] $name) {
    $e = $zip.GetEntry($name)
    if ($null -ne $e) { return $e }
    # Some writers use a leading slash or different casing.
    foreach ($c in $zip.Entries) {
        if ($c.FullName.TrimStart('/') -ieq $name) { return $c }
    }
    return $null
}

function Read-ZipXmlDocument($zip, [string] $name) {
    $e = Get-ZipEntry $zip $name
    if ($null -eq $e) { return $null }
    $s = $e.Open()
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null
        $doc.Load($s)
        return $doc
    } finally { $s.Dispose() }
}

function Select-Local($node, [string] $path) {
    # Namespace-agnostic selection, so both strict and transitional OOXML work.
    $xp = ($path -split '/' | Where-Object { $_ } | ForEach-Object { "*[local-name()='$_']" }) -join '/'
    return $node.SelectNodes($xp)
}

function Read-SharedStrings($zip) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    $e = Get-ZipEntry $zip 'xl/sharedStrings.xml'
    if ($null -eq $e) { return $list }
    $s = $e.Open()
    $xr = New-XmlReaderOn $s
    try {
        $sb = New-Object System.Text.StringBuilder 256
        $advance = $true
        while ($true) {
            if ($advance) { if (-not $xr.Read()) { break } } else { $advance = $true }
            if ($xr.NodeType -eq [System.Xml.XmlNodeType]::Element -and $xr.LocalName -eq 'si') {
                $xml = $xr.ReadOuterXml()     # already positioned on the next node
                $advance = $false
                if ($xml.IndexOf('<rPh') -ge 0) { $xml = $RxRPh.Replace($xml, '') }
                [void] $sb.Clear()
                foreach ($m in $RxText.Matches($xml)) { [void] $sb.Append($m.Groups[1].Value) }
                $list.Add((Expand-XmlText $sb.ToString()))
            }
        }
    } finally { $xr.Dispose() }
    return $list
}

# Returns 0 = plain, 1 = date, 2 = time, 3 = date and time.
function Get-FormatKind([string] $code) {
    if ([string]::IsNullOrEmpty($code)) { return 0 }
    $c = $code
    # An elapsed-time token such as [h] or [m] means time whatever else the code
    # holds. Recording that first keeps a bare [m], elapsed minutes, from being
    # read as a month by the fallback at the end of this function.
    $elapsed = [regex]::IsMatch($c, '\[(h+|m+|s+)\]', 'IgnoreCase')
    $c = [regex]::Replace($c, '\[(h+|m+|s+)\]', '$1', 'IgnoreCase')
    $c = [regex]::Replace($c, '\[[^\]]*\]', '')
    $c = [regex]::Replace($c, '"[^"]*"', '')
    $c = [regex]::Replace($c, '\\.', '')
    $c = $c.ToLowerInvariant()
    $hasDate = ($c.IndexOf('y') -ge 0) -or ($c.IndexOf('d') -ge 0)
    $hasTime = $elapsed -or ($c.IndexOf('h') -ge 0) -or ($c.IndexOf('s') -ge 0)
    if ((-not $hasDate) -and (-not $hasTime) -and $c.IndexOf('m') -ge 0) { $hasDate = $true }
    if ($hasDate -and $hasTime) { return 3 }
    if ($hasDate) { return 1 }
    if ($hasTime) { return 2 }
    return 0
}

function Get-BuiltinFormatKind([int] $id) {
    if ($id -ge 14 -and $id -le 17) { return 1 }
    if ($id -ge 18 -and $id -le 21) { return 2 }
    if ($id -eq 22) { return 3 }
    if ($id -ge 27 -and $id -le 36) { return 1 }
    if ($id -ge 45 -and $id -le 47) { return 2 }
    if ($id -ge 50 -and $id -le 58) { return 1 }
    return 0
}

# Maps each cell style index to a format kind.
function Read-StyleKinds($zip) {
    $kinds = New-Object 'System.Collections.Generic.List[int]'
    $doc = Read-ZipXmlDocument $zip 'xl/styles.xml'
    if ($null -eq $doc) { return $kinds }

    $custom = @{}
    foreach ($n in (Select-Local $doc 'styleSheet/numFmts/numFmt')) {
        $id = [int] $n.GetAttribute('numFmtId')
        $custom[$id] = Get-FormatKind $n.GetAttribute('formatCode')
    }
    foreach ($xf in (Select-Local $doc 'styleSheet/cellXfs/xf')) {
        $idText = $xf.GetAttribute('numFmtId')
        $id = 0
        if ($idText) { [void] [int]::TryParse($idText, [ref] $id) }
        if ($custom.ContainsKey($id)) { $kinds.Add($custom[$id]) }
        else { $kinds.Add((Get-BuiltinFormatKind $id)) }
    }
    return $kinds
}

function Get-WorkbookSheets($zip) {
    $doc = Read-ZipXmlDocument $zip 'xl/workbook.xml'
    if ($null -eq $doc) { return $null }

    $rels = @{}
    $relDoc = Read-ZipXmlDocument $zip 'xl/_rels/workbook.xml.rels'
    if ($null -ne $relDoc) {
        foreach ($r in (Select-Local $relDoc 'Relationships/Relationship')) {
            $rels[$r.GetAttribute('Id')] = $r.GetAttribute('Target')
        }
    }

    $date1904 = $false
    foreach ($pr in (Select-Local $doc 'workbook/workbookPr')) {
        $v = $pr.GetAttribute('date1904')
        if ($v -eq '1' -or $v -eq 'true') { $date1904 = $true }
    }

    $sheets = @()
    foreach ($s in (Select-Local $doc 'workbook/sheets/sheet')) {
        $rid = $null
        foreach ($a in $s.Attributes) { if ($a.LocalName -eq 'id') { $rid = $a.Value } }
        $target = $null
        if ($rid -and $rels.ContainsKey($rid)) { $target = $rels[$rid] }
        if (-not $target) { continue }
        $target = $target -replace '^/xl/', '' -replace '^/', ''
        if ($target -notmatch '^xl/') { $target = 'xl/' + $target }
        $sheets += [pscustomobject] @{ Name = $s.GetAttribute('name'); Target = $target }
    }
    return [pscustomobject] @{ Sheets = $sheets; Date1904 = $date1904 }
}

function Format-SerialDate([double] $serial, [int] $kind, [bool] $d1904) {
    if ($kind -eq 2) {
        $frac = $serial - [Math]::Floor($serial)
        return ([timespan]::FromDays($frac)).ToString('hh\:mm\:ss', $Inv)
    }
    $s = $serial
    if ($d1904) { $s += 1462 }
    $dt = $null
    # Excel counts a day that never existed, 1900-02-29. From serial 61 on, the
    # .NET OLE automation date lines up exactly; below 60 it is off by one.
    if ($s -ge 61 -and $s -le 2958465) { $dt = [datetime]::FromOADate($s) }
    elseif ($s -ge 1 -and $s -lt 60) { $dt = [datetime]::FromOADate($s + 1) }
    if ($null -eq $dt) { return $null }
    if ($kind -eq 3) { return $dt.ToString($Cul.DateTimeFormat.ShortDatePattern + ' HH:mm:ss', $Cul) }
    return $dt.ToString($Cul.DateTimeFormat.ShortDatePattern, $Cul)
}

function Export-Sheet($zip, [string] $target, [string] $destination, $shared, $styleKinds, [bool] $d1904) {

    $entry = Get-ZipEntry $zip $target
    if ($null -eq $entry) { throw ($L.NoSheet -f $target) }

    $delim = $Delimiter
    if ([string]::IsNullOrEmpty($delim)) { $delim = ';' }

    $useComma = $false
    if ($DecimalSeparator -eq 'comma') { $useComma = $true }
    elseif ($DecimalSeparator -eq 'auto') {
        $useComma = ($delim -ne ',') -and ($Cul.NumberFormat.NumberDecimalSeparator -eq ',')
    }
    $boolTrue = 'TRUE'; $boolFalse = 'FALSE'
    if ($EffectiveLanguage -eq 'fr') { $boolTrue = 'VRAI'; $boolFalse = 'FAUX' }

    # --- Pass 1: how wide is the sheet -------------------------------------
    # The dimension element is authoritative when present and sane; otherwise
    # every cell reference is scanned.
    $maxCol = 0
    $s1 = $entry.Open()
    $xr = New-XmlReaderOn $s1
    try {
        while ($xr.Read()) {
            if ($xr.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            if ($xr.LocalName -eq 'dimension') {
                $ref = $xr.GetAttribute('ref')
                if ($ref) {
                    $last = $ref.Split(':')[-1]
                    $ci = Get-ColumnIndex $last
                    if ($ci -gt 0 -and $ci -le 16384) { $maxCol = $ci; break }
                }
            } elseif ($xr.LocalName -eq 'sheetData') {
                break
            }
        }
        if ($maxCol -eq 0) {
            while ($xr.Read()) {
                if ($xr.NodeType -eq [System.Xml.XmlNodeType]::Element -and $xr.LocalName -eq 'c') {
                    $r = $xr.GetAttribute('r')
                    if ($r) {
                        $ci = Get-ColumnIndex $r
                        if ($ci -gt $maxCol) { $maxCol = $ci }
                    }
                }
            }
        }
    } finally { $xr.Dispose() }

    # --- Pass 2: write the rows ---------------------------------------------
    $encPref = $Encoding
    if ($encPref -eq 'auto') { $encPref = 'utf8bom' }
    $encObj = switch ($encPref) {
        'utf8bom' { New-Object System.Text.UTF8Encoding($true) }
        'utf8'    { New-Object System.Text.UTF8Encoding($false) }
        'ansi'    { [System.Text.Encoding]::Default }
        'unicode' { New-Object System.Text.UnicodeEncoding($false, $true) }
    }

    $destDir = [System.IO.Path]::GetDirectoryName($destination)
    $tmp = [System.IO.Path]::Combine($destDir, '~' + [System.IO.Path]::GetFileName($destination) + '.tmp')
    $sharedCount = $shared.Count
    $kindCount = $styleKinds.Count
    $rowsOut = 0

    $sw = New-Object System.IO.StreamWriter($tmp, $false, $encObj, 65536)
    try {
        $sw.NewLine = "`r`n"
        $s2 = $entry.Open()
        $xr = New-XmlReaderOn $s2
        try {
            $cells = New-Object string[] ([Math]::Max($maxCol, 1))
            $sb = New-Object System.Text.StringBuilder 8192
            $expected = 1
            $advance = $true

            while ($true) {
                if ($advance) { if (-not $xr.Read()) { break } } else { $advance = $true }
                if ($xr.NodeType -ne [System.Xml.XmlNodeType]::Element -or $xr.LocalName -ne 'row') { continue }

                $rowIdx = $expected
                $rAttr = $xr.GetAttribute('r')
                if ($rAttr) { [void] [int]::TryParse($rAttr, [ref] $rowIdx) }
                $xml = $xr.ReadOuterXml()
                $advance = $false

                # Rows Excel never wrote still exist between two used rows.
                while ($expected -lt $rowIdx) {
                    $sw.WriteLine([string]::Join($delim, (New-Object string[] $cells.Length)))
                    $expected++
                    $rowsOut++
                }

                [Array]::Clear($cells, 0, $cells.Length)
                $lastCol = 0

                # This loop runs once per cell in the workbook, so attributes are
                # picked apart with IndexOf rather than helper calls or regexes.
                foreach ($m in $RxCell.Matches($xml)) {
                    $attrs = $m.Groups[1].Value

                    $col = 0
                    $ai = $attrs.IndexOf('r="')
                    if ($ai -ge 0) {
                        $ai += 3
                        while ($ai -lt $attrs.Length) {
                            $ch = $attrs[$ai]
                            if ($ch -ge 'A' -and $ch -le 'Z') { $col = $col * 26 + ([int][char] $ch - 64); $ai++ }
                            else { break }
                        }
                    }
                    if ($col -eq 0) { $col = $lastCol + 1 }
                    $lastCol = $col
                    if ($col -gt $cells.Length) { continue }

                    $type = ''
                    $ai = $attrs.IndexOf('t="')
                    if ($ai -ge 0) {
                        $aj = $attrs.IndexOf('"', $ai + 3)
                        if ($aj -gt 0) { $type = $attrs.Substring($ai + 3, $aj - $ai - 3) }
                    }

                    $inner = $m.Groups[2].Value
                    $out = $null

                    if ($type -eq 'inlineStr') {
                        [void] $sb.Clear()
                        foreach ($t in $RxText.Matches($inner)) { [void] $sb.Append($t.Groups[1].Value) }
                        $out = $sb.ToString()
                        if ($out.IndexOf('&') -ge 0 -or $out.IndexOf('_x') -ge 0) { $out = Expand-XmlText $out }
                    } else {
                        $rawValue = $null
                        $vi = $inner.IndexOf('<v>')
                        if ($vi -ge 0) {
                            $ve = $inner.IndexOf('</v>', $vi + 3)
                            if ($ve -ge 0) { $rawValue = $inner.Substring($vi + 3, $ve - $vi - 3) }
                        }
                        if ($null -eq $rawValue) {
                            $vm = $RxVal.Match($inner)
                            if (-not $vm.Success) { continue }
                            $rawValue = $vm.Groups[1].Value
                        }

                        if ($type -eq 's') {
                            $idx = 0
                            if ([int]::TryParse($rawValue, [ref] $idx) -and $idx -ge 0 -and $idx -lt $sharedCount) {
                                $out = $shared[$idx]
                            } else { $out = '' }
                        } elseif ($type -eq 'b') {
                            if ($rawValue -eq '1') { $out = $boolTrue } else { $out = $boolFalse }
                        } elseif ($type -eq 'str' -or $type -eq 'e' -or $type -eq 'd') {
                            # 'str' is a cached formula string, 'e' an error value
                            # such as #DIV/0!, 'd' an ISO date in strict OOXML.
                            $out = $rawValue
                            if ($out.IndexOf('&') -ge 0 -or $out.IndexOf('_x') -ge 0) { $out = Expand-XmlText $out }
                        } else {
                            # Numeric cell: a date number format decides how it reads.
                            $kind = 0
                            if ((-not $Raw) -and $kindCount -gt 0) {
                                $ai = $attrs.IndexOf('s="')
                                if ($ai -ge 0) {
                                    $ai += 3
                                    $si = 0
                                    $seen = $false
                                    while ($ai -lt $attrs.Length) {
                                        $ch = $attrs[$ai]
                                        if ($ch -ge '0' -and $ch -le '9') { $si = $si * 10 + ([int][char] $ch - 48); $ai++; $seen = $true }
                                        else { break }
                                    }
                                    if ($seen -and $si -lt $kindCount) { $kind = $styleKinds[$si] }
                                }
                            }
                            $d = 0.0
                            if ([double]::TryParse($rawValue, [System.Globalization.NumberStyles]::Float, $Inv, [ref] $d)) {
                                if ($kind -ne 0) { $out = Format-SerialDate $d $kind $d1904 }
                                if ($null -eq $out) {
                                    $out = $d.ToString('R', $Inv)
                                    if ($useComma) { $out = $out.Replace('.', ',') }
                                }
                            } else {
                                $out = Expand-XmlText $rawValue
                            }
                        }
                    }

                    $cells[$col - 1] = $out
                }

                # --- Assemble one CSV record ---------------------------------
                [void] $sb.Clear()
                for ($i = 0; $i -lt $cells.Length; $i++) {
                    if ($i -gt 0) { [void] $sb.Append($delim) }
                    $f = $cells[$i]
                    if ($null -eq $f -or $f.Length -eq 0) { continue }
                    $quote = ($f.IndexOf($delim) -ge 0) -or ($f.IndexOf('"') -ge 0) -or
                             ($f.IndexOf("`n") -ge 0) -or ($f.IndexOf("`r") -ge 0) -or
                             [char]::IsWhiteSpace($f[0]) -or [char]::IsWhiteSpace($f[$f.Length - 1])
                    if ($quote) { [void] $sb.Append('"').Append($f.Replace('"', '""')).Append('"') }
                    else { [void] $sb.Append($f) }
                }
                $sw.WriteLine($sb.ToString())
                $rowsOut++
                $expected = $rowIdx + 1
            }
        } finally { $xr.Dispose() }
        $sw.Flush()
    } catch {
        $sw.Dispose()
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        throw
    } finally {
        $sw.Dispose()
    }

    if (Test-Path -LiteralPath $destination) { Remove-Item -LiteralPath $destination -Force }
    [System.IO.File]::Move($tmp, $destination)

    return [pscustomobject] @{
        Output    = $destination
        Rows      = $rowsOut
        Columns   = $cells.Length
        Delimiter = $delim
        Encoding  = $encObj.WebName
    }
}

function Convert-XlsxToCsv([string] $source) {
    if ($source -match '\.(xlsb|xls)$') { throw $L.Xlsb }

    $zip = [System.IO.Compression.ZipFile]::OpenRead($source)
    try {
        $wb = Get-WorkbookSheets $zip
        if ($null -eq $wb) { throw $L.NotXlsx }
        if ($wb.Sheets.Count -eq 0) { throw $L.NoSheets }

        $wanted = $wb.Sheets
        if ($Sheet) {
            $idx = 0
            if ([int]::TryParse($Sheet, [ref] $idx)) {
                if ($idx -lt 1 -or $idx -gt $wb.Sheets.Count) { throw ($L.NoSheet -f $Sheet) }
                $wanted = @($wb.Sheets[$idx - 1])
            } else {
                $wanted = @($wb.Sheets | Where-Object { $_.Name -eq $Sheet })
                if ($wanted.Count -eq 0) { throw ($L.NoSheet -f $Sheet) }
            }
        }

        if ($OutFile -and $wanted.Count -gt 1) { throw $L.SingleOut }

        $shared = Read-SharedStrings $zip
        $kinds = Read-StyleKinds $zip
        $dir = [System.IO.Path]::GetDirectoryName($source)
        $base = [System.IO.Path]::GetFileNameWithoutExtension($source)

        $out = @()
        foreach ($sh in $wanted) {
            if ($OutFile) {
                $dest = Resolve-OutFilePath $source $OutFile
            } elseif ($wanted.Count -gt 1) {
                $dest = Get-FreePath $dir (ConvertTo-SafeFileName ($base + ' - ' + $sh.Name)) '.csv'
            } else {
                $dest = Get-FreePath $dir $base '.csv'
            }
            $res = Export-Sheet $zip $sh.Target $dest $shared $kinds $wb.Date1904
            $res | Add-Member -NotePropertyName Source -NotePropertyValue $source
            $res | Add-Member -NotePropertyName Sheet -NotePropertyValue $sh.Name
            $out += $res
        }
        return $out
    } finally { $zip.Dispose() }
}

# ===========================================================================
#  Right-click menu integration
# ===========================================================================

$ScriptPath   = $MyInvocation.MyCommand.Path
$ScriptDir    = Split-Path -Parent $ScriptPath
$ScriptLeaf   = Split-Path -Leaf $ScriptPath
$LauncherPath = Join-Path $ScriptDir ([System.IO.Path]::GetFileNameWithoutExtension($ScriptLeaf) + '.launcher.vbs')

# Verb key names are prefixed so the semicolon entry sorts above the comma one:
# Explorer lists verbs in registry key order.
$Verbs = @(
    @{ Key = 'RcX2C2X1ToXlsx';     Side = 'csv';  Args = '';               Label = 'MenuToXlsx' }
    @{ Key = 'RcX2C2X1ToCsvSemi';  Side = 'xlsx'; Args = '/sep:semicolon'; Label = 'MenuToCsvSc' }
    @{ Key = 'RcX2C2X2ToCsvComma'; Side = 'xlsx'; Args = '/sep:comma';     Label = 'MenuToCsvC' }
)

function Get-VerbKey([string] $ext, [string] $verb) {
    return "HKCU:\Software\Classes\SystemFileAssociations\$ext\shell\$verb"
}

# The launcher exists only so Explorer does not flash a console window on every
# conversion. It is generated here so the tool ships as a single file.
function Write-Launcher {
    $template = @'
' Generated by @@SCRIPT@@ -Install. Removed by -Uninstall.
' Runs the converter through PowerShell without flashing a console window.
'
' File paths and the script path travel in environment variables, never inside
' the command string. WshShell.Run expands %NAME% in that string, so a file
' called "Q1 %USERNAME% report.csv" would otherwise reach PowerShell with the
' percent section replaced and fail as not found. Environment variables are
' inherited verbatim, and the command string below contains no percent sign at
' all. It also means no caller-supplied text ever reaches the parser.
Option Explicit

Dim fso, shell, script, cmd, files, lang, sep, i, arg

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

script = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "@@SCRIPT@@")
If Not fso.FileExists(script) Then
    MsgBox "Script not found:" & vbCrLf & script, vbCritical, "@@TITLE@@"
    WScript.Quit 1
End If

lang = "en"
sep = ""
files = ""

For i = 0 To WScript.Arguments.Count - 1
    arg = WScript.Arguments(i)
    If LCase(Left(arg, 6)) = "/lang:" Then
        lang = LCase(Mid(arg, 7))
    ElseIf LCase(Left(arg, 5)) = "/sep:" Then
        sep = LCase(Mid(arg, 6))
    Else
        If files <> "" Then files = files & vbLf
        files = files & arg
    End If
Next

If files = "" Then WScript.Quit 0
If lang <> "fr" Then lang = "en"

shell.Environment("Process")("RCX2C2X_SCRIPT") = script
shell.Environment("Process")("RCX2C2X_FILES") = files

cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass" _
    & " -Command ""& $env:RCX2C2X_SCRIPT -Gui -FromLauncher -Language " & lang

' Only the XLSX verbs send /sep:, because there the separator is a choice about
' the file being written. The CSV verb must send nothing at all, so the script
' keeps detecting the separator of the file it is reading.
If sep <> "" Then
    Select Case sep
        Case "comma"
            cmd = cmd & " -Delimiter ','"
        Case "tab"
            cmd = cmd & " -Delimiter ([char]9)"
        Case Else
            cmd = cmd & " -Delimiter ';'"
    End Select
End If

WScript.Quit shell.Run(cmd & """", 0, True)
'@
    $vbs = $template.Replace('@@SCRIPT@@', $ScriptLeaf).Replace('@@TITLE@@', $ToolName)
    [System.IO.File]::WriteAllText($LauncherPath, $vbs, (New-Object System.Text.UTF8Encoding($false)))
}

function Invoke-Install {
    # --- Language choice -----------------------------------------------------
    $lang = $Language
    if (-not $lang) {
        $lang = 'en'
        if ([Environment]::UserInteractive) {
            $en = $Strings['en']
            Write-Host ''
            Write-Host $en.AskLang -ForegroundColor Cyan
            Write-Host $en.AskEn
            Write-Host $en.AskFr
            $answer = ''
            try { $answer = Read-Host $en.AskPrompt } catch { $answer = '' }
            if ($answer -and $answer.Trim() -match '^(2|fr|francais|french)$') { $lang = 'fr' }
            Write-Host ''
        }
    }
    $T = $Strings[$lang]

    Write-Launcher
    Write-Host ($T.WroteLauncher -f $LauncherPath) -ForegroundColor DarkGray

    # Prefer Excel's own icon when Excel is installed.
    $icon = "$env:SystemRoot\System32\imageres.dll,-102"
    $appPaths = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'
    if (Test-Path $appPaths) {
        $exe = (Get-ItemProperty -Path $appPaths -Name '(default)' -ErrorAction SilentlyContinue).'(default)'
        if ($exe -and (Test-Path $exe)) { $icon = "$exe,0" }
    }

    foreach ($ext in @($CsvExtensions + $XlsxExtensions | Select-Object -Unique)) {
        $side = 'xlsx'
        if ($CsvExtensions -contains $ext) { $side = 'csv' }

        foreach ($v in $Verbs) {
            if ($v.Side -ne $side) { continue }

            $key = Get-VerbKey $ext $v.Key
            New-Item -Path $key -Force | Out-Null
            New-Item -Path "$key\command" -Force | Out-Null

            $extra = $v.Args
            if ($extra) { $extra = $extra + ' ' }
            $command = '"{0}\System32\wscript.exe" "{1}" {2}/lang:{3} "%1"' -f $env:SystemRoot, $LauncherPath, $extra, $lang

            Set-ItemProperty -Path $key -Name '(default)'        -Value $T[$v.Label]
            Set-ItemProperty -Path $key -Name 'Icon'             -Value $icon
            Set-ItemProperty -Path $key -Name 'MultiSelectModel' -Value 'Player'
            Set-ItemProperty -Path "$key\command" -Name '(default)' -Value $command

            Write-Host ($T.Installed -f "$ext  $($T[$v.Label])") -ForegroundColor Green
        }
    }

    Write-Host ''
    Write-Host $T.DoneInstall -ForegroundColor Cyan
}

function Invoke-Uninstall {
    $T = $Strings['en']

    # Sweep every registered file type rather than the current -CsvExtensions and
    # -XlsxExtensions. An install that added .xlsm must still be fully removed by
    # a plain -Uninstall, which would otherwise never look at that extension.
    $root = 'HKCU:\Software\Classes\SystemFileAssociations'
    if (Test-Path $root) {
        foreach ($extKey in @(Get-ChildItem -Path $root -ErrorAction SilentlyContinue)) {
            $shell = Join-Path $extKey.PSPath 'shell'
            if (-not (Test-Path $shell)) { continue }
            foreach ($v in $Verbs) {
                $key = Join-Path $shell $v.Key
                if (Test-Path $key) {
                    Remove-Item -Path $key -Recurse -Force
                    Write-Host ($T.Removed -f "$($extKey.PSChildName) $($v.Key)") -ForegroundColor Yellow
                }
            }
            # Leave the shell key behind only if something else still uses it.
            if (-not @(Get-ChildItem -Path $shell -ErrorAction SilentlyContinue)) {
                Remove-Item -Path $shell -Force -ErrorAction SilentlyContinue
            }
        }
    }
    if (Test-Path -LiteralPath $LauncherPath) {
        Remove-Item -LiteralPath $LauncherPath -Force
        Write-Host ($T.Removed -f $LauncherPath) -ForegroundColor Yellow
    }
    Write-Host $T.DoneRemove -ForegroundColor Green
}

# ===========================================================================
#  Entry point
# ===========================================================================

if ($Version) {
    Write-Host "$ToolName $ToolVersion"
    exit 0
}

if ($Uninstall) { Invoke-Uninstall; exit 0 }
if ($Install)   { Invoke-Install;   exit 0 }

if ($FromLauncher) {
    $fromEnv = $env:RCX2C2X_FILES
    if ($fromEnv) { $Path = @($fromEnv -split "`r?`n" | Where-Object { $_.Trim().Length -gt 0 }) }
}

if ((-not $Path) -or $Path.Count -eq 0) {
    Write-Host $L.Usage
    Write-Host ''
    Write-Host "$ToolName $ToolVersion  Copyright (C) 2026 Guillaume Taurel" -ForegroundColor DarkGray
    Write-Host $L.Licence -ForegroundColor DarkGray
    exit 2
}

# Flatten whatever shape the caller used: bare arguments, an array, or a mix.
$requested = @()
foreach ($p in $Path) {
    if (($p -is [System.Collections.IEnumerable]) -and ($p -isnot [string])) {
        foreach ($q in $p) { $requested += [string] $q }
    } else {
        $requested += [string] $p
    }
}

$targets = @()
foreach ($item in $requested) {
    if ([string]::IsNullOrWhiteSpace($item)) { continue }
    $matched = $null
    try { $matched = @(Resolve-Path -Path $item -ErrorAction Stop) } catch { $matched = $null }
    if ($matched) {
        foreach ($rp in $matched) {
            if (Test-Path -LiteralPath $rp.Path -PathType Leaf) { $targets += $rp.Path }
        }
    } else {
        $targets += $item   # the main loop will report the error
    }
}

if ($OutFile -and $targets.Count -gt 1) { throw $L.SingleOut }

$results = @()
$failures = @()

foreach ($src in $targets) {
    try {
        if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { throw $L.NotFound }
        $full = (Get-Item -LiteralPath $src).FullName
        $ext = [System.IO.Path]::GetExtension($full).ToLowerInvariant()

        if ($ReadableXlsx -contains $ext) {
            $res = Convert-XlsxToCsv $full
            $results += $res
            if ($Open) { foreach ($r in $res) { Start-Process -FilePath $r.Output } }
        } elseif ($ReadableCsv -contains $ext) {
            if ($OutFile) { $dest = Resolve-OutFilePath $full $OutFile }
            else { $dest = Get-FreePath ([System.IO.Path]::GetDirectoryName($full)) ([System.IO.Path]::GetFileNameWithoutExtension($full)) '.xlsx' }
            $res = Convert-CsvToXlsx $full $dest
            $results += $res
            if ($Open) { Start-Process -FilePath $res.Output }
        } else {
            throw ($L.Unknown -f $ext, ($ReadableCsv -join ', '), ($ReadableXlsx -join ', '))
        }
    } catch {
        $failures += ([System.IO.Path]::GetFileName($src) + ': ' + $_.Exception.Message)
    }
}

if ($Gui) {
    if ($failures.Count -gt 0) {
        Add-Type -AssemblyName System.Windows.Forms
        $msg = ($L.DlgFailed -f $failures.Count) + "`r`n`r`n" + ($failures -join "`r`n")
        [void] [System.Windows.Forms.MessageBox]::Show($msg, $L.DlgTitle, 'OK', 'Error')
    }
} else {
    $results
    foreach ($f in $failures) { Write-Error $f -ErrorAction Continue }
}

if ($failures.Count -gt 0) { exit 1 }
exit 0
