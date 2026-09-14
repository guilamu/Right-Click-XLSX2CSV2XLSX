# ============================================================================
# Suite de tests pour RightClickXlsx2Csv2Xlsx.ps1
# Harnais INDEPENDANT : le XLSX produit est relu par un lecteur maison
# (XmlDocument + XPath local-name()), pas par le script teste, pour eviter
# qu'un bogue dans un sens soit masque par le bogue symetrique dans l'autre.
#
# Usage :
#   pwsh -File .\Tests\Invoke-ConversionTests.ps1 [-Engine powershell.exe]
#   pwsh -File .\Tests\Invoke-ConversionTests.ps1 -Engine pwsh -Filter 'xlsx2csv'
#
# Ne modifie rien au script teste. Tout se passe dans %TEMP%.
# ============================================================================
[CmdletBinding()]
param(
    [string] $Engine = 'powershell.exe',   # moteur qui execute le script teste
    [string] $Filter = '',                 # regex sur le nom du test
    [switch] $Keep                         # conserve le dossier de travail
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ScriptUnderTest = Join-Path $PSScriptRoot '..\RightClickXlsx2Csv2Xlsx.ps1'
$ScriptUnderTest = [System.IO.Path]::GetFullPath($ScriptUnderTest)
if (-not (Test-Path -LiteralPath $ScriptUnderTest)) { throw "Script teste introuvable : $ScriptUnderTest" }

$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$Cul = [System.Globalization.CultureInfo]::CurrentCulture
$Sdp = $Cul.DateTimeFormat.ShortDatePattern

$Work = Join-Path $env:TEMP ('rcx2c2x-tests-' + [guid]::NewGuid().ToString('n').Substring(0, 12))
New-Item -ItemType Directory -Path $Work | Out-Null

# L'option -Encoding ansi du script vise la page de codes ANSI sous les deux
# moteurs, y compris sous pwsh 7 ou Encoding.Default vaut UTF-8 : les attentes
# des tests d'encodage sont donc identiques quel que soit le moteur.
Write-Host "Moteur fils : $Engine" -ForegroundColor DarkCyan

$script:Results = @()
$script:CurrentTest = ''
$script:CurrentDir = $Work

# Tests qui documentent un bogue CONFIRME du script : ils doivent echouer.
# S'ils se mettent a passer, le bogue est corrige (XPASS) et la carte saute.
# BUG-DIM, BUG-ELAPSED et BUG-SHARED1 sont corriges en 1.0.1 : aucune entree.
$script:KnownBugs = @{}

# ---------------------------------------------------------------------------
# Assertions
# ---------------------------------------------------------------------------
function Test-Case([string] $Name, [scriptblock] $Body) {
    if ($Filter -and $Name -notmatch $Filter) { return }
    $script:CurrentTest = $Name
    $dir = Join-Path $Work ($Name -replace '[^\w-]', '_')
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $script:CurrentDir = $dir
    Push-Location $dir
    try {
        & $Body
        if ($script:KnownBugs.ContainsKey($Name)) {
            $script:Results += [pscustomobject]@{ Test = $Name; Result = 'XPASS'; Detail = 'bogue connu apparemment corrige : ' + $script:KnownBugs[$Name] }
            Write-Host "XPASS $Name : bogue connu corrige ? $($script:KnownBugs[$Name])" -ForegroundColor Magenta
        } else {
            $script:Results += [pscustomobject]@{ Test = $Name; Result = 'PASS'; Detail = '' }
            Write-Host "PASS  $Name" -ForegroundColor Green
        }
    } catch {
        if ($script:KnownBugs.ContainsKey($Name)) {
            $script:Results += [pscustomobject]@{ Test = $Name; Result = 'XFAIL'; Detail = $script:KnownBugs[$Name] + '  |  ' + $_.Exception.Message }
            Write-Host "XFAIL $Name : $($script:KnownBugs[$Name])" -ForegroundColor Yellow
        } else {
            $script:Results += [pscustomobject]@{ Test = $Name; Result = 'FAIL'; Detail = $_.Exception.Message }
            Write-Host "FAIL  $Name : $($_.Exception.Message)" -ForegroundColor Red
        }
    } finally {
        Pop-Location
    }
}

function Assert-True($Cond, [string] $Msg) {
    if (-not $Cond) { throw "Assertion : $Msg" }
}

function Assert-Equal($Expected, $Actual, [string] $Msg) {
    if (-not [object]::Equals($Expected, $Actual)) {
        throw "Assertion : $Msg. Attendu=[$Expected] Obtenu=[$Actual]"
    }
}

function Assert-Contains([string] $Haystack, [string] $Needle, [string] $Msg) {
    # La console replie les lignes longues (Out-String, largeur 80) : on
    # normalise les blancs des deux cotes avant de chercher.
    $h = ($Haystack -replace '\s+', ' ')
    $n = ($Needle -replace '\s+', ' ')
    if (-not $h.Contains($n)) {
        throw "Assertion : $Msg. Absent=[$n] dans [$h]"
    }
}

# ---------------------------------------------------------------------------
# Execution du script teste dans un processus fils
# ---------------------------------------------------------------------------
function Invoke-Converter([string[]] $Arguments) {
    $out = & $Engine -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $ScriptUnderTest @Arguments 2>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $out }
}

# ---------------------------------------------------------------------------
# Ecriture de fichiers CSV avec encodage explicite
# ---------------------------------------------------------------------------
# NB harnais : il tourne sous pwsh 7, ou [Encoding]::Default vaut UTF-8.
# Pour ecrire de vrais octets ANSI on utilise iso-8859-1 (identique a cp1252
# pour les caracteres des fixtures : e avec accent, a avec circonflexe...).
$script:EncUtf8NoBom = New-Object System.Text.UTF8Encoding($false)
$script:EncUtf8Bom   = New-Object System.Text.UTF8Encoding($true)
$script:EncAnsi      = [System.Text.Encoding]::GetEncoding('iso-8859-1')
$script:EncUtf16Le   = New-Object System.Text.UnicodeEncoding($false, $true)

function Write-CsvFile([string] $Path, [string] $Content, $Enc) {
    if ($null -eq $Enc) { $Enc = $script:EncUtf8Bom }
    [System.IO.File]::WriteAllText($Path, $Content, $Enc)
}

# ---------------------------------------------------------------------------
# Lecteur XLSX independent (XmlDocument, XPath local-name())
# Retourne : Sheets[] = @{ Name; Cells = @{ ref -> @{R;T;S;V} }; RowCount; HasPane; HasAutoFilter; Dimension }
#            Styles = liste des numFmtId par index de xf ; NumFmts = customs ; Date1904 ; Shared
# ---------------------------------------------------------------------------
function Read-XlsxPackage([string] $Path) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $getXml = {
            param([string] $EntryName)
            $e = $zip.GetEntry($EntryName)
            if ($null -eq $e) { return $null }
            $st = $e.Open()
            try {
                $doc = New-Object System.Xml.XmlDocument
                $doc.XmlResolver = $null
                $doc.Load($st)
                return $doc
            } finally { $st.Dispose() }
        }

        $wb = & $getXml 'xl/workbook.xml'
        if ($null -eq $wb) { throw "Pas de workbook.xml dans $Path" }

        $rels = @{}
        $relDoc = & $getXml 'xl/_rels/workbook.xml.rels'
        if ($null -ne $relDoc) {
            foreach ($r in $relDoc.SelectNodes("/*[local-name()='Relationships']/*[local-name()='Relationship']")) {
                $rels[$r.GetAttribute('Id')] = $r.GetAttribute('Target')
            }
        }

        $date1904 = $false
        foreach ($pr in $wb.SelectNodes("/*[local-name()='workbook']/*[local-name()='workbookPr']")) {
            $v = $pr.GetAttribute('date1904')
            if ($v -eq '1' -or $v -eq 'true') { $date1904 = $true }
        }

        $shared = New-Object 'System.Collections.Generic.List[string]'
        $ssd = & $getXml 'xl/sharedStrings.xml'
        if ($null -ne $ssd) {
            foreach ($si in $ssd.SelectNodes("/*[local-name()='sst']/*[local-name()='si']")) {
                $t = ''
                foreach ($n in $si.SelectNodes(".//*[local-name()='t']")) { $t += $n.InnerText }
                $shared.Add($t)
            }
        }

        $styles = New-Object 'System.Collections.Generic.List[int]'
        $numFmts = @{}
        $sty = & $getXml 'xl/styles.xml'
        if ($null -ne $sty) {
            foreach ($n in $sty.SelectNodes("//*[local-name()='numFmt']")) {
                $numFmts[[int] $n.GetAttribute('numFmtId')] = $n.GetAttribute('formatCode')
            }
            foreach ($xf in $sty.SelectNodes("//*[local-name()='cellXfs']/*[local-name()='xf']")) {
                $id = 0
                [void] [int]::TryParse($xf.GetAttribute('numFmtId'), [ref] $id)
                $styles.Add($id)
            }
        }

        $sheets = @()
        foreach ($s in $wb.SelectNodes("/*[local-name()='workbook']/*[local-name()='sheets']/*[local-name()='sheet']")) {
            $rid = $null
            foreach ($a in $s.Attributes) { if ($a.LocalName -eq 'id') { $rid = $a.Value } }
            $target = $rels[$rid]
            $target = $target -replace '^/xl/', '' -replace '^/', ''
            if ($target -notmatch '^xl/') { $target = 'xl/' + $target }
            $sd = & $getXml $target
            if ($null -eq $sd) { throw "Feuille '$target' absente du paquet" }

            $cells = @{}
            $rowCount = 0
            foreach ($row in $sd.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
                $rowCount++
                foreach ($c in $row.SelectNodes("./*[local-name()='c']")) {
                    $t = $c.GetAttribute('t')
                    $v = $null
                    if ($t -eq 'inlineStr') {
                        $v = ''
                        foreach ($n in $c.SelectNodes("./*[local-name()='is']//*[local-name()='t']")) { $v += $n.InnerText }
                    } else {
                        $vn = $c.SelectSingleNode("./*[local-name()='v']")
                        if ($null -ne $vn) { $v = $vn.InnerText }
                        if ($t -eq 's' -and $null -ne $v) { $v = $shared[[int] $v] }
                    }
                    $cells[$c.GetAttribute('r')] = [pscustomobject]@{
                        R = $c.GetAttribute('r'); T = $t; S = $c.GetAttribute('s'); V = $v
                    }
                }
            }
            $dim = $null
            foreach ($d in $sd.SelectNodes("/*[local-name()='worksheet']/*[local-name()='dimension']")) { $dim = $d.GetAttribute('ref') }
            $sheets += [pscustomobject]@{
                Name          = $s.GetAttribute('name')
                Cells         = $cells
                RowCount      = $rowCount
                HasPane       = ($sd.OuterXml.Contains('<pane '))
                HasAutoFilter = ($null -ne $sd.SelectSingleNode("/*[local-name()='worksheet']/*[local-name()='autoFilter']"))
                Dimension     = $dim
            }
        }
        return [pscustomobject]@{
            Sheets = $sheets; Styles = $styles; NumFmts = $numFmts
            Date1904 = $date1904; Shared = $shared
        }
    } finally { $zip.Dispose() }
}

function Get-Cell($Pkg, [int] $SheetIdx, [string] $Ref) {
    $c = $Pkg.Sheets[$SheetIdx].Cells[$Ref]
    return $c
}

function Assert-NumCell($Cell, [double] $Expected, [string] $What) {
    Assert-True ($null -ne $Cell) "$What : cellule absente"
    Assert-True ($Cell.T -ne 'inlineStr' -and $Cell.T -ne 's') "$What : devrait etre numerique, t='$($Cell.T)' v='$($Cell.V)'"
    $d = 0.0
    Assert-True ([double]::TryParse($Cell.V, [System.Globalization.NumberStyles]::Float, $Inv, [ref] $d)) "$What : valeur '$($Cell.V)' non numerique"
    Assert-True ([math]::Abs($d - $Expected) -lt 1e-9) "$What : attendu $Expected, obtenu $($Cell.V)"
}

function Assert-TextCell($Cell, [string] $Expected, [string] $What) {
    Assert-True ($null -ne $Cell) "$What : cellule absente"
    Assert-True ($Cell.T -eq 'inlineStr' -or $Cell.T -eq 's') "$What : devrait etre texte, t='$($Cell.T)' v='$($Cell.V)'"
    Assert-Equal $Expected $Cell.V $What
}

# ---------------------------------------------------------------------------
# Lecteur CSV maison (pour verifier la sortie XLSX->CSV)
# ---------------------------------------------------------------------------
function Read-CsvAll([string] $Path, [string] $Delim) {
    $sr = New-Object System.IO.StreamReader($Path, $true)
    $p = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($sr)
    $p.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $p.SetDelimiters([string[]] @($Delim))
    $p.HasFieldsEnclosedInQuotes = $true
    $p.TrimWhiteSpace = $false
    $rows = @()
    while (-not $p.EndOfData) {
        $f = $p.ReadFields()
        if ($null -eq $f) { $rows += , @() } else { $rows += , @($f) }
    }
    $p.Dispose()
    # Le pipeline aplatirait le tableau de tableaux : la virgule le protege.
    return , $rows
}

# ---------------------------------------------------------------------------
# Constructeur de fixtures XLSX (paquets faits main pour tester la lecture)
# Sheets : @{ Name; Rows = @( @{ R = int (optionnel); Cells = @( @{R;T;S;V} ) } ); Dimension = string | $null (absent si non fourni -> calcule) ; OmitDimension }
# Shared : chaines XML BRUTES (deja encodees) pour sharedStrings.xml
# NumFmts : @( @{Id;Code} ) ; Xfs : int[] des numFmtId par index de style
# ---------------------------------------------------------------------------
function Get-ColLetter([int] $index) {
    $n = ''
    while ($index -gt 0) {
        $r = ($index - 1) % 26
        $n = [string] [char] (65 + $r) + $n
        $index = [int] (($index - $r - 1) / 26)
    }
    return $n
}

function Add-ZipEntry($zip, [string] $Name, [string] $Content) {
    $entry = $zip.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $w = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        $w.Write($Content)
        $w.Flush()
        $w.Dispose()
    } finally { $stream.Dispose() }
}

function New-XlsxFixture {
    param(
        [string] $Path,
        [array]  $Sheets,
        [string[]] $Shared = @(),
        [array]  $NumFmts = @(),
        [int[]]  $Xfs = @(0),
        [bool]   $Date1904 = $false,
        [switch] $SkipWorkbook
    )
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        if (-not $SkipWorkbook) {
            $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>'
            Add-ZipEntry $zip '[Content_Types].xml' $ct
            Add-ZipEntry $zip '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

            $wbXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            if ($Date1904) { $wbXml += '<workbookPr date1904="1"/>' }
            $wbXml += '<sheets>'
            for ($i = 0; $i -lt $Sheets.Count; $i++) {
                $wbXml += '<sheet name="' + [System.Security.SecurityElement]::Escape($Sheets[$i].Name) + '" sheetId="' + ($i + 1) + '" r:id="rId' + ($i + 1) + '"/>'
            }
            $wbXml += '</sheets></workbook>'
            Add-ZipEntry $zip 'xl/workbook.xml' $wbXml

            $relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            for ($i = 0; $i -lt $Sheets.Count; $i++) {
                $relsXml += '<Relationship Id="rId' + ($i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + ($i + 1) + '.xml"/>'
            }
            $relsXml += '<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
            Add-ZipEntry $zip 'xl/_rels/workbook.xml.rels' $relsXml

            # styles.xml
            $styXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            if ($NumFmts.Count -gt 0) {
                $styXml += '<numFmts count="' + $NumFmts.Count + '">'
                foreach ($nf in $NumFmts) {
                    $styXml += '<numFmt numFmtId="' + $nf.Id + '" formatCode="' + [System.Security.SecurityElement]::Escape($nf.Code) + '"/>'
                }
                $styXml += '</numFmts>'
            }
            $styXml += '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>'
            $styXml += '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            $styXml += '<cellXfs count="' + $Xfs.Count + '">'
            foreach ($x in $Xfs) { $styXml += '<xf numFmtId="' + $x + '" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' }
            $styXml += '</cellXfs></styleSheet>'
            Add-ZipEntry $zip 'xl/styles.xml' $styXml

            if ($Shared.Count -gt 0) {
                $ssXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + $Shared.Count + '" uniqueCount="' + $Shared.Count + '">'
                foreach ($s in $Shared) { $ssXml += '<si><t>' + $s + '</t></si>' }
                $ssXml += '</sst>'
                Add-ZipEntry $zip 'xl/sharedStrings.xml' $ssXml
            }

            for ($i = 0; $i -lt $Sheets.Count; $i++) {
                $sh = $Sheets[$i]
                $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                $maxRow = 0; $maxCol = 0
                foreach ($row in $sh.Rows) {
                    $ri = 0
                    if ($row.ContainsKey('R')) { $ri = [int] $row.R }
                    foreach ($c in $row.Cells) {
                        if ($c.ContainsKey('R') -and $c.R) {
                            if ($c.R -match '^([A-Z]+)([0-9]+)$') {
                                $ci = 0
                                foreach ($ch in $Matches[1].ToCharArray()) { $ci = $ci * 26 + ([int][char] $ch - 64) }
                                if ($ci -gt $maxCol) { $maxCol = $ci }
                                if ([int] $Matches[2] -gt $maxRow) { $maxRow = [int] $Matches[2] }
                            }
                        }
                    }
                    if ($ri -gt $maxRow) { $maxRow = $ri }
                }
                if ($sh.ContainsKey('Dimension') -and $null -ne $sh.Dimension -and $sh.Dimension -ne '') {
                    $xml += '<dimension ref="' + $sh.Dimension + '"/>'
                } elseif ((-not $sh.ContainsKey('OmitDimension') -or -not $sh.OmitDimension) -and $maxRow -gt 0) {
                    if ($maxCol -eq 0) { $maxCol = 1 }
                    $xml += '<dimension ref="A1:' + (Get-ColLetter $maxCol) + $maxRow + '"/>'
                }
                $xml += '<sheetData>'
                $rowNum = 0
                foreach ($row in $sh.Rows) {
                    $rowNum++
                    $ri = $rowNum
                    if ($row.ContainsKey('R')) { $ri = [int] $row.R }
                    $xml += '<row r="' + $ri + '">'
                    $colNum = 0
                    foreach ($c in $row.Cells) {
                        $colNum++
                        $ref = $null
                        if ($c.ContainsKey('R') -and $c.R) { $ref = $c.R } else { $ref = (Get-ColLetter $colNum) + $ri }
                        $attrs = ' r="' + $ref + '"'
                        if ($c.ContainsKey('S') -and $null -ne $c.S -and [int] $c.S -gt 0) { $attrs += ' s="' + $c.S + '"' }
                        $t = ''
                        if ($c.ContainsKey('T')) { $t = $c.T }
                        if ($t) { $attrs += ' t="' + $t + '"' }
                        $v = ''
                        if ($c.ContainsKey('V')) { $v = [string] $c.V }
                        if ($t -eq 'inlineStr') {
                            $xml += '<c' + $attrs + '><is><t>' + [System.Security.SecurityElement]::Escape($v) + '</t></is></c>'
                        } elseif ($t -eq 's') {
                            $xml += '<c' + $attrs + '><v>' + $v + '</v></c>'
                        } elseif ($null -eq $c.V -or $v -eq '') {
                            $xml += '<c' + $attrs + '/>'
                        } else {
                            $xml += '<c' + $attrs + '><v>' + [System.Security.SecurityElement]::Escape($v) + '</v></c>'
                        }
                    }
                    $xml += '</row>'
                }
                $xml += '</sheetData></worksheet>'
                Add-ZipEntry $zip ('xl/worksheets/sheet' + ($i + 1) + '.xml') $xml
            }
        } else {
            Add-ZipEntry $zip 'dummy.txt' 'ceci n est pas un classeur'
        }
    } finally {
        $zip.Dispose()
        $fs.Dispose()
    }
}

# Serial OADate attendu
function OAD([string] $Iso) { return ([datetime]::ParseExact($Iso, 'yyyy-MM-dd HH:mm:ss', $Inv)).ToOADate() }

# Rendu date attendu cote XLSX->CSV (meme regle que le script : culture courante)
function Render-Date([double] $Serial) {
    $dt = [datetime]::FromOADate($Serial)
    return $dt.ToString($Sdp, $Cul)
}
function Render-DateTime([double] $Serial) {
    $dt = [datetime]::FromOADate($Serial)
    return $dt.ToString($Sdp + ' HH:mm:ss', $Cul)
}

# ############################################################################
#  TESTS CSV -> XLSX
# ############################################################################

Test-Case 'csv2xlsx-types-et-entete' {
    $csv = Join-Path $script:CurrentDir 'basic.csv'
    Write-CsvFile $csv ("Nom;Entier;DecPoint;DecVirgule;Zero;Long16;Date;DateHeure;Vide;Texte`r`nfoo;42;3.14;2,5;0012345;1234567890123456;2024-06-15;2024-06-15 13:45:30;;bar")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'basic.xlsx')
    Assert-Equal 'basic' $pkg.Sheets[0].Name 'nom de feuille = nom du fichier'
    Assert-Equal 2 $pkg.Sheets[0].RowCount '2 lignes'
    Assert-Equal 'A1:J2' $pkg.Sheets[0].Dimension 'dimension'
    Assert-True $pkg.Sheets[0].HasPane 'volet fige sur entete'
    Assert-True $pkg.Sheets[0].HasAutoFilter 'autofiltre sur entete'
    Assert-TextCell (Get-Cell $pkg 0 'A1') 'Nom' 'entete texte'
    Assert-Equal '1' (Get-Cell $pkg 0 'A1').S 'style gras entete'
    Assert-TextCell (Get-Cell $pkg 0 'A2') 'foo' 'A2'
    Assert-NumCell  (Get-Cell $pkg 0 'B2') 42 'entier'
    Assert-NumCell  (Get-Cell $pkg 0 'C2') 3.14 'decimal point'
    Assert-NumCell  (Get-Cell $pkg 0 'D2') 2.5 'decimal virgule (separateur ;)'
    Assert-TextCell (Get-Cell $pkg 0 'E2') '0012345' 'zero initial conserve en texte'
    Assert-TextCell (Get-Cell $pkg 0 'F2') '1234567890123456' '16 chiffres reste texte'
    Assert-NumCell  (Get-Cell $pkg 0 'G2') (OAD '2024-06-15 00:00:00') 'date ISO -> serie'
    Assert-Equal '2' (Get-Cell $pkg 0 'G2').S 'style date (numFmtId 14)'
    Assert-NumCell  (Get-Cell $pkg 0 'H2') (OAD '2024-06-15 13:45:30') 'date-heure -> serie'
    Assert-Equal '3' (Get-Cell $pkg 0 'H2').S 'style date-heure (numFmtId 22)'
    Assert-True ($null -eq (Get-Cell $pkg 0 'I2')) 'champ vide -> cellule absente'
    Assert-TextCell (Get-Cell $pkg 0 'J2') 'bar' 'J2'
}

Test-Case 'csv2xlsx-virgule-separateur' {
    $csv = Join-Path $script:CurrentDir 'virgule.csv'
    # "2,5" entre guillemets avec separateur virgule : doit rester TEXTE (virgule = separateur)
    Write-CsvFile $csv ("Nom,Entier,DecPoint,Quoted`r`nfoo,42,3.14,""2,5""")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'virgule.xlsx')
    Assert-NumCell  (Get-Cell $pkg 0 'B2') 42 'entier'
    Assert-NumCell  (Get-Cell $pkg 0 'C2') 3.14 'decimal point'
    Assert-TextCell (Get-Cell $pkg 0 'D2') '2,5' 'virgule decimale ambigue reste texte'
}

Test-Case 'csv2xlsx-tab-tsv' {
    $csv = Join-Path $script:CurrentDir 'tab.tsv'
    Write-CsvFile $csv ("a`tb`tc`r`n1`t2`t3")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'tab.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'C1') 'c' '3 colonnes detectees (tab)'
    Assert-NumCell  (Get-Cell $pkg 0 'C2') 3 'C2'
}

Test-Case 'csv2xlsx-pipe' {
    $csv = Join-Path $script:CurrentDir 'pipe.csv'
    Write-CsvFile $csv ("a|b|c`r`n1|2|3")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'pipe.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'C1') 'c' 'pipe detecte'
}

Test-Case 'csv2xlsx-directive-sep' {
    $csv = Join-Path $script:CurrentDir 'sep.csv'
    Write-CsvFile $csv ("sep=;`r`na;b`r`n1;2")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'sep.xlsx')
    Assert-Equal 2 $pkg.Sheets[0].RowCount 'la ligne sep= est consommee'
    Assert-TextCell (Get-Cell $pkg 0 'A1') 'a' 'A1'
    Assert-NumCell  (Get-Cell $pkg 0 'B2') 2 'B2'
}

Test-Case 'csv2xlsx-champs-quotes-multiligne' {
    $csv = Join-Path $script:CurrentDir 'q.csv'
    Write-CsvFile $csv ("H1;H2;H3`r`n""a;b"";""c""""d"";""multi`nline""")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'q.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A2') 'a;b' 'separateur dans guillemets'
    Assert-TextCell (Get-Cell $pkg 0 'B2') 'c"d' 'guillemets doubles'
    Assert-TextCell (Get-Cell $pkg 0 'C2') ("multi`nline") 'saut de ligne dans champ'
}

Test-Case 'csv2xlsx-encodages-lecture' {
    foreach ($case in @(
        @{ N = 'utf8nobom'; Enc = $script:EncUtf8NoBom },
        @{ N = 'utf8bom';   Enc = $script:EncUtf8Bom },
        @{ N = 'ansi';      Enc = $script:EncAnsi },
        @{ N = 'utf16';     Enc = $script:EncUtf16Le }
    )) {
        $csv = Join-Path $script:CurrentDir ($case.N + '.csv')
        Write-CsvFile $csv ("Nom;Ville`r`nRemi;Oeuvre-sur-Seine") $case.Enc
        $r = Invoke-Converter @($csv)
        Assert-Equal 0 $r.ExitCode "$($case.N) exit code ($($r.Output))"
        $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir ($case.N + '.xlsx'))
        Assert-TextCell (Get-Cell $pkg 0 'A2') 'Remi' "$($case.N) valeur accentuee"
        Assert-TextCell (Get-Cell $pkg 0 'B2') 'Oeuvre-sur-Seine' "$($case.N) B2"
    }
    # Variante avec accents REELS (non ASCII) pour utf8/ansi
    $csv = Join-Path $script:CurrentDir 'utf8accents.csv'
    [System.IO.File]::WriteAllText($csv, "Pr`u{e9}nom;Ville`r`nRen`u{e9};Ch`u{e2}teau", $script:EncUtf8NoBom)
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "utf8 accents exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'utf8accents.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A1') "Pr`u{e9}nom" 'utf8 sans BOM : entete accentuee'
    Assert-TextCell (Get-Cell $pkg 0 'B2') "Ch`u{e2}teau" 'utf8 sans BOM : valeur accentuee'

    $csv = Join-Path $script:CurrentDir 'ansiaccents.csv'
    [System.IO.File]::WriteAllText($csv, "Pr`u{e9}nom;Ville`r`nRen`u{e9};Ch`u{e2}teau", $script:EncAnsi)
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "ansi accents exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'ansiaccents.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'B2') "Ch`u{e2}teau" 'ANSI detecte : valeur accentuee'
}

Test-Case 'csv2xlsx-encodage-force' {
    # Fichier UTF-8 SANS BOM relu de force en ANSI -> mojibake attendu
    $csv = Join-Path $script:CurrentDir 'f.csv'
    [System.IO.File]::WriteAllText($csv, "Nom`r`n`u{e9}t`u{e9}", $script:EncUtf8NoBom)
    $r = Invoke-Converter @($csv, '-Encoding', 'ansi')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'f.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A2') ($script:EncAnsi.GetString($script:EncUtf8NoBom.GetBytes("`u{e9}t`u{e9}"))) 'override ANSI produit du mojibake controle'
}

Test-Case 'csv2xlsx-astext' {
    $csv = Join-Path $script:CurrentDir 't.csv'
    Write-CsvFile $csv ("A;B`r`n42;2024-06-15")
    $r = Invoke-Converter @($csv, '-AsText')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 't.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A2') '42' 'nombre reste texte avec -AsText'
    Assert-TextCell (Get-Cell $pkg 0 'B2') '2024-06-15' 'date reste texte avec -AsText'
}

Test-Case 'csv2xlsx-noheader' {
    $csv = Join-Path $script:CurrentDir 'h.csv'
    Write-CsvFile $csv ("42;abc`r`n7;def")
    $r = Invoke-Converter @($csv, '-NoHeader')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'h.xlsx')
    Assert-True (-not $pkg.Sheets[0].HasPane) 'pas de volet fige'
    Assert-True (-not $pkg.Sheets[0].HasAutoFilter) 'pas d autofiltre'
    Assert-NumCell (Get-Cell $pkg 0 'A1') 42 'ligne 1 typee nombre sans -NoHeader implicite'
    Assert-True ((Get-Cell $pkg 0 'B1').S -ne '1') 'pas de style entete'
}

Test-Case 'csv2xlsx-fichier-vide' {
    $csv = Join-Path $script:CurrentDir 'vide.csv'
    [System.IO.File]::WriteAllBytes($csv, [byte[]] @())
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'vide.xlsx')
    Assert-Equal 0 $pkg.Sheets[0].RowCount 'aucune ligne'
}

Test-Case 'csv2xlsx-ligne-unique' {
    $csv = Join-Path $script:CurrentDir 'u.csv'
    Write-CsvFile $csv "2024-06-15;123"
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'u.xlsx')
    Assert-True (-not $pkg.Sheets[0].HasAutoFilter) 'pas d autofiltre avec 1 seule ligne'
    Assert-True (-not $pkg.Sheets[0].HasPane) 'pas de volet fige avec 1 seule ligne'
    # Comportement epingle : une ligne unique n est pas une entete, les
    # valeurs SONT typees (pas de notion d entete avec rowCount < 2).
    Assert-NumCell (Get-Cell $pkg 0 'A1') (OAD '2024-06-15 00:00:00') 'ligne unique typee comme donnee'
    Assert-NumCell (Get-Cell $pkg 0 'B1') 123 'B1 typee'
}

Test-Case 'csv2xlsx-champs-vides-et-irreguliers' {
    $csv = Join-Path $script:CurrentDir 'j.csv'
    Write-CsvFile $csv ("a;b;;`r`n;;;c`r`n1;2;3;4;5")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'j.xlsx')
    Assert-Equal 'A1:E3' $pkg.Sheets[0].Dimension 'dimension sur la ligne la plus large'
    Assert-True ($null -eq (Get-Cell $pkg 0 'C1')) 'C1 vide'
    Assert-True ($null -eq (Get-Cell $pkg 0 'D1')) 'D1 vide (trailing)'
    Assert-TextCell (Get-Cell $pkg 0 'D2') 'c' 'D2'
    Assert-NumCell  (Get-Cell $pkg 0 'E3') 5 'E3'
}

Test-Case 'csv2xlsx-lignes-vides-ignorees' {
    # TextFieldParser ignore les lignes vides : le comportement est epingle ici.
    $csv = Join-Path $script:CurrentDir 'b.csv'
    Write-CsvFile $csv ("a;b`r`n`r`n1;2")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'b.xlsx')
    Assert-Equal 2 $pkg.Sheets[0].RowCount 'ligne vide physique ignoree'
    Assert-NumCell (Get-Cell $pkg 0 'A2') 1 'A2'
}

Test-Case 'csv2xlsx-caracteres-xml' {
    $csv = Join-Path $script:CurrentDir 'x.csv'
    Write-CsvFile $csv ("H`r`n""<a & """"b""""> suite 'c'""")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'x.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A2') "<a & ""b""> suite 'c'" 'echappement XML'
}

Test-Case 'csv2xlsx-caracteres-controle' {
    $csv = Join-Path $script:CurrentDir 'c.csv'
    [System.IO.File]::WriteAllText($csv, "H`r`na" + [char] 1 + "b" + [char] 7 + "c", $script:EncUtf8Bom)
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'c.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A2') 'abc' 'caracteres de controle retires'
}

Test-Case 'csv2xlsx-dates-limites' {
    $csv = Join-Path $script:CurrentDir 'd.csv'
    Write-CsvFile $csv ("H1;H2;H3;H4;H5`r`n1899-12-31;1900-02-28;1900-03-01;2024-02-30;15/06/2024")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'd.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A2') '1899-12-31' 'avant 1900 : texte'
    Assert-TextCell (Get-Cell $pkg 0 'B2') '1900-02-28' 'avant le 01/03/1900 : texte'
    Assert-NumCell  (Get-Cell $pkg 0 'C2') (OAD '1900-03-01 00:00:00') '01/03/1900 : date (serie 61)'
    Assert-TextCell (Get-Cell $pkg 0 'D2') '2024-02-30' 'date invalide : texte'
    # date locale dd/MM/yyyy (depend de la culture courante)
    $dt = [datetime]::MinValue
    if ([datetime]::TryParseExact('15/06/2024', @($Sdp), $Cul, 'None', [ref] $dt)) {
        Assert-NumCell (Get-Cell $pkg 0 'E2') $dt.ToOADate() 'date locale typee'
    } else {
        Assert-TextCell (Get-Cell $pkg 0 'E2') '15/06/2024' 'date locale non reconnue sur cette culture'
    }
}

Test-Case 'csv2xlsx-guillemets-non-fermes' {
    $csv = Join-Path $script:CurrentDir 'm.csv'
    Write-CsvFile $csv ("a;b`r`n""unclosed;rest`r`nmore;lines")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "ligne malformee : conversion reussie quand meme ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'm.xlsx')
    Assert-True ($pkg.Sheets[0].RowCount -ge 2) 'au moins 2 lignes conservees'
    $c = Get-Cell $pkg 0 'A2'
    Assert-True ($null -ne $c -and $c.V.Contains('unclosed')) 'ligne malformee conservee verbatim'
}

Test-Case 'csv2xlsx-espaces-preserves' {
    $csv = Join-Path $script:CurrentDir 'w.csv'
    Write-CsvFile $csv ("H1;H2`r`n"" lead"";""trail """)
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'w.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A2') ' lead' 'espace de tete'
    Assert-TextCell (Get-Cell $pkg 0 'B2') 'trail ' 'espace de fin'
}

Test-Case 'csv2xlsx-outfile-et-protections' {
    $csv = Join-Path $script:CurrentDir 's.csv'
    Write-CsvFile $csv ("a;b`r`n1;2")
    $out = Join-Path $script:CurrentDir 'dest.xlsx'
    $r = Invoke-Converter @($csv, '-OutFile', $out)
    Assert-Equal 0 $r.ExitCode "OutFile premiere ecriture ($($r.Output))"
    Assert-True (Test-Path -LiteralPath $out) 'OutFile cree'
    $r = Invoke-Converter @($csv, '-OutFile', $out)
    Assert-Equal 1 $r.ExitCode 'OutFile existant refuse sans -Force'
    Assert-Contains $r.Output 'already exists' 'message OutExists'
    $r = Invoke-Converter @($csv, '-OutFile', $out, '-Force')
    Assert-Equal 0 $r.ExitCode "OutFile + Force ($($r.Output))"
    $r = Invoke-Converter @($csv, '-OutFile', $csv, '-Force')
    Assert-Equal 1 $r.ExitCode 'OutFile == source refuse meme avec -Force'
    Assert-Contains $r.Output 'itself' 'message OutIsSource'
}

Test-Case 'csv2xlsx-nom-libre-automatique' {
    $csv = Join-Path $script:CurrentDir 'f.csv'
    Write-CsvFile $csv ("a`r`n1")
    $r = Invoke-Converter @($csv); Assert-Equal 0 $r.ExitCode "conv 1 ($($r.Output))"
    $r = Invoke-Converter @($csv); Assert-Equal 0 $r.ExitCode "conv 2 ($($r.Output))"
    $r = Invoke-Converter @($csv); Assert-Equal 0 $r.ExitCode "conv 3 ($($r.Output))"
    Assert-True (Test-Path -LiteralPath (Join-Path $script:CurrentDir 'f.xlsx')) 'f.xlsx'
    Assert-True (Test-Path -LiteralPath (Join-Path $script:CurrentDir 'f (1).xlsx')) 'f (1).xlsx'
    Assert-True (Test-Path -LiteralPath (Join-Path $script:CurrentDir 'f (2).xlsx')) 'f (2).xlsx'
}

Test-Case 'csv2xlsx-extension-txt-et-nom-feuille' {
    $csv = Join-Path $script:CurrentDir 't[e]st.txt'
    Write-CsvFile $csv ("a;b`r`n1;2")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode ".txt accepte ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 't[e]st.xlsx')
    Assert-Equal 't_e_st' $pkg.Sheets[0].Name 'crochets du nom de fichier neutralises dans le nom de feuille'
}

Test-Case 'csv2xlsx-delimiter-override' {
    $csv = Join-Path $script:CurrentDir 'o.csv'
    Write-CsvFile $csv ("a,b,c`r`n1,2,3")
    $r = Invoke-Converter @($csv, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'o.xlsx')
    Assert-Equal 'A1:A2' $pkg.Sheets[0].Dimension 'avec -Delimiter ; la ligne entiere est un seul champ'
    Assert-TextCell (Get-Cell $pkg 0 'A1') 'a,b,c' 'champ unique'
}

Test-Case 'csv2xlsx-nombres-signes-et-formes' {
    $csv = Join-Path $script:CurrentDir 'n.csv'
    Write-CsvFile $csv ("H1;H2;H3;H4;H5;H6;H7`r`n-12;+7;-3.5;.5;5.;1e5; 42 ")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 'n.xlsx')
    Assert-NumCell  (Get-Cell $pkg 0 'A2') -12 'negatif'
    Assert-NumCell  (Get-Cell $pkg 0 'B2') 7 '+7'
    Assert-NumCell  (Get-Cell $pkg 0 'C2') -3.5 'negatif decimal'
    Assert-NumCell  (Get-Cell $pkg 0 'D2') 0.5 '.5'
    Assert-NumCell  (Get-Cell $pkg 0 'E2') 5 '5.'
    Assert-TextCell (Get-Cell $pkg 0 'F2') '1e5' 'notation scientifique non typee'
    Assert-NumCell  (Get-Cell $pkg 0 'G2') 42 'espaces autour du nombre : trime et type'
}

Test-Case 'csv2xlsx-heure-seule-reste-texte' {
    $csv = Join-Path $script:CurrentDir 't.csv'
    Write-CsvFile $csv ("H`r`n13:45")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $pkg = Read-XlsxPackage (Join-Path $script:CurrentDir 't.xlsx')
    Assert-TextCell (Get-Cell $pkg 0 'A2') '13:45' 'heure seule non typee'
}

# ############################################################################
#  TESTS XLSX -> CSV
# ############################################################################

Test-Case 'xlsx2csv-types-cellules' {
    $fx = Join-Path $script:CurrentDir 'types.xlsx'
    New-XlsxFixture $fx -Shared @('Bonjour', 'le;monde') -Xfs @(0) -Sheets @(
        @{ Name = 'S1'; Rows = @(
            @{ Cells = @(
                @{ T = 's'; V = '0' },                    # A1 shared string
                @{ T = 's'; V = '1' },                    # B1 shared string avec separateur
                @{ T = 'inlineStr'; V = 'enligne' },      # C1
                @{ T = 'b'; V = '1' },                    # D1 booleen
                @{ T = 'b'; V = '0' },                    # E1
                @{ T = 'e'; V = '#DIV/0!' },              # F1 erreur
                @{ T = 'str'; V = 'resformule' },         # G1 resultat de formule
                @{ T = '';  V = '3.5' },                  # H1 nombre
                @{ T = '';  V = '-42' }                   # I1
            ) }
        ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'types.csv') ';'
    Assert-Equal 1 $rows.Count '1 ligne'
    $line = $rows[0]
    Assert-Equal 'Bonjour' $line[0] 'shared string'
    Assert-Equal 'le;monde' $line[1] 'shared string avec separateur (quote)'
    Assert-Equal 'enligne' $line[2] 'inline string'
    Assert-Equal 'TRUE' $line[3] 'booleen vrai'
    Assert-Equal 'FALSE' $line[4] 'booleen faux'
    Assert-Equal '#DIV/0!' $line[5] 'valeur erreur'
    Assert-Equal 'resformule' $line[6] 'resultat de formule mis en cache'
    if ($Cul.NumberFormat.NumberDecimalSeparator -eq ',') {
        Assert-Equal '3,5' $line[7] 'decimal auto -> virgule (culture fr)'
    } else {
        Assert-Equal '3.5' $line[7] 'decimal auto -> point'
    }
    Assert-Equal '-42' $line[8] 'negatif'
}

Test-Case 'xlsx2csv-multi-feuilles' {
    $fx = Join-Path $script:CurrentDir 'multi.xlsx'
    New-XlsxFixture $fx -Sheets @(
        @{ Name = 'Premiere'; Rows = @( @{ Cells = @( @{ T = 'inlineStr'; V = 'un' } ) } ) },
        @{ Name = 'Seconde Feuille'; Rows = @( @{ Cells = @( @{ T = 'inlineStr'; V = 'deux' } ) } ) }
    )
    $r = Invoke-Converter @($fx)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $f1 = Join-Path $script:CurrentDir 'multi - Premiere.csv'
    $f2 = Join-Path $script:CurrentDir 'multi - Seconde Feuille.csv'
    Assert-True (Test-Path -LiteralPath $f1) 'un CSV par feuille (1)'
    Assert-True (Test-Path -LiteralPath $f2) 'un CSV par feuille (2)'
    $rows = Read-CsvAll $f2 ';'
    Assert-Equal 'deux' $rows[0][0] 'contenu feuille 2'

    # -Sheet par nom et par index
    $sub = Join-Path $script:CurrentDir 'sub'; New-Item -ItemType Directory -Path $sub | Out-Null
    Copy-Item $fx (Join-Path $sub 'm.xlsx')
    $r = Invoke-Converter @((Join-Path $sub 'm.xlsx'), '-Sheet', 'Seconde Feuille')
    Assert-Equal 0 $r.ExitCode "-Sheet nom ($($r.Output))"
    Assert-True (Test-Path -LiteralPath (Join-Path $sub 'm.csv')) 'sortie unique sans suffixe'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $sub 'm - Premiere.csv'))) 'pas de suffixe feuille'
    Remove-Item (Join-Path $sub 'm.csv')
    $r = Invoke-Converter @((Join-Path $sub 'm.xlsx'), '-Sheet', '2')
    Assert-Equal 0 $r.ExitCode "-Sheet index ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $sub 'm.csv') ';'
    Assert-Equal 'deux' $rows[0][0] 'index 2 = seconde feuille'

    $r = Invoke-Converter @((Join-Path $sub 'm.xlsx'), '-Sheet', 'Introuvable')
    Assert-Equal 1 $r.ExitCode '-Sheet nom inexistant -> echec'
    Assert-Contains $r.Output 'was not found' 'message NoSheet'
    $r = Invoke-Converter @((Join-Path $sub 'm.xlsx'), '-Sheet', '9')
    Assert-Equal 1 $r.ExitCode '-Sheet index hors limites -> echec'

    # -OutFile avec plusieurs feuilles -> refus
    $r = Invoke-Converter @($fx, '-OutFile', (Join-Path $script:CurrentDir 'x.csv'))
    Assert-Equal 1 $r.ExitCode '-OutFile multi-feuilles refuse'
    Assert-Contains $r.Output 'single sheet' 'message SingleOut'
}

Test-Case 'xlsx2csv-dates-et-raw' {
    $fx = Join-Path $script:CurrentDir 'dates.xlsx'
    $serialDate = OAD '2024-06-15 00:00:00'
    $serialDt   = OAD '2024-06-15 13:45:30'
    New-XlsxFixture $fx -Xfs @(0, 14, 22, 20) -Sheets @(
        @{ Name = 'S1'; Rows = @(
            @{ Cells = @(
                @{ T = ''; V = $serialDate.ToString('R', $Inv); S = 1 },   # date (numFmtId 14)
                @{ T = ''; V = $serialDt.ToString('R', $Inv);   S = 2 },   # date+heure (22)
                @{ T = ''; V = '0.625';                       S = 3 },     # heure (20) -> 15:00:00
                @{ T = ''; V = $serialDate.ToString('R', $Inv); S = 0 }    # meme serie sans style date
            ) }
        ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'dates.csv') ';'
    $line = $rows[0]
    Assert-Equal (Render-Date $serialDate) $line[0] 'date rendue selon la locale'
    Assert-Equal (Render-DateTime $serialDt) $line[1] 'date-heure rendue selon la locale'
    Assert-Equal '15:00:00' $line[2] 'heure seule rendue hh:mm:ss'
    Assert-Equal ($serialDate.ToString('R', $Inv).Replace('.', $(if ($Cul.NumberFormat.NumberDecimalSeparator -eq ',') { ',' } else { '.' }))) $line[3] 'sans style date : nombre brut'

    # -Raw : les series restent numeriques
    $sub = Join-Path $script:CurrentDir 'raw'; New-Item -ItemType Directory -Path $sub | Out-Null
    Copy-Item $fx (Join-Path $sub 'r.xlsx')
    $r = Invoke-Converter @((Join-Path $sub 'r.xlsx'), '-Delimiter', ';', '-Raw')
    Assert-Equal 0 $r.ExitCode "-Raw exit code ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $sub 'r.csv') ';'
    Assert-True ($rows[0][0].StartsWith(([string] [int] $serialDate))) '-Raw : serie brute, pas de date'
}

Test-Case 'xlsx2csv-decimal-separator' {
    $fx = Join-Path $script:CurrentDir 'd.xlsx'
    New-XlsxFixture $fx -Sheets @( @{ Name = 'S1'; Rows = @( @{ Cells = @( @{ T = ''; V = '3.5' } ) } ) } )
    $r = Invoke-Converter @($fx, '-Delimiter', ';', '-DecimalSeparator', 'dot')
    Assert-Equal 0 $r.ExitCode "dot ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'd.csv') ';'
    Assert-Equal '3.5' $rows[0][0] '-DecimalSeparator dot'
    Remove-Item (Join-Path $script:CurrentDir 'd.csv')
    $r = Invoke-Converter @($fx, '-Delimiter', ';', '-DecimalSeparator', 'comma')
    Assert-Equal 0 $r.ExitCode "comma ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'd.csv') ';'
    Assert-Equal '3,5' $rows[0][0] '-DecimalSeparator comma'
    Remove-Item (Join-Path $script:CurrentDir 'd.csv')
    # auto avec separateur virgule -> toujours point
    $r = Invoke-Converter @($fx, '-Delimiter', ',')
    Assert-Equal 0 $r.ExitCode "auto + virgule ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'd.csv') ','
    Assert-Equal '3.5' $rows[0][0] 'auto avec separateur , -> point'
}

Test-Case 'xlsx2csv-date1904' {
    $fx = Join-Path $script:CurrentDir 'd1904.xlsx'
    New-XlsxFixture $fx -Date1904 $true -Xfs @(0, 14) -Sheets @(
        @{ Name = 'S1'; Rows = @( @{ Cells = @( @{ T = ''; V = '0'; S = 1 }, @{ T = ''; V = '1462'; S = 1 } ) } ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'd1904.csv') ';'
    Assert-Equal ([datetime]::new(1904, 1, 1).ToString($Sdp, $Cul)) $rows[0][0] 'serie 0 en 1904 = 01/01/1904'
    Assert-Equal ([datetime]::FromOADate(1462 + 1462).ToString($Sdp, $Cul)) $rows[0][1] 'serie 1462 decalee de 1462 jours'
}

Test-Case 'xlsx2csv-lignes-vides-preservees' {
    $fx = Join-Path $script:CurrentDir 'rows.xlsx'
    New-XlsxFixture $fx -Sheets @(
        @{ Name = 'S1'; Rows = @(
            @{ R = 1; Cells = @( @{ R = 'A1'; T = 'inlineStr'; V = 'a' }, @{ R = 'B1'; T = 'inlineStr'; V = 'b' } ) },
            @{ R = 3; Cells = @( @{ R = 'A3'; T = 'inlineStr'; V = 'c' }, @{ R = 'B3'; T = 'inlineStr'; V = 'd' } ) },
            @{ R = 5; Cells = @( @{ R = 'A5'; T = 'inlineStr'; V = 'e' }, @{ R = 'B5'; T = 'inlineStr'; V = 'f' } ) }
        ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $text = [System.IO.File]::ReadAllText((Join-Path $script:CurrentDir 'rows.csv'))
    Assert-Equal ("a;b`r`n;`r`nc;d`r`n;`r`ne;f`r`n") $text 'lignes vides preservees a la bonne position'
}

Test-Case 'xlsx2csv-echappements-xml-et-xHHHH' {
    $fx = Join-Path $script:CurrentDir 'esc.xlsx'
    New-XlsxFixture $fx -Shared @(
        'a&lt;b&amp;c&quot;d&apos;e',   # 0 : entites XML classiques
        '&lt;',                         # 1 : '<' seul
        '_x005F__x0041_',               # 2 : texte litteral "_x0041_" tel que Excel le stocke (double prefixe)
        '_x005F_x0041_',                # 3 : texte litteral "_x0041_" tel que Excel le stocke (forme standard)
        'tab_x0009_ici',                # 4 : tabulation echappee
        '_x005F_'                       # 5 : texte litteral "_x005F_" ... non : "_x005F_" decode = "_"
    ) -Sheets @(
        @{ Name = 'S1'; Rows = @( @{ Cells = @(
            @{ T = 's'; V = '0' }, @{ T = 's'; V = '1' }, @{ T = 's'; V = '2' },
            @{ T = 's'; V = '3' }, @{ T = 's'; V = '4' }, @{ T = 's'; V = '5' }
        ) } ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'esc.csv') ';'
    $line = $rows[0]
    Assert-Equal 'a<b&c"d''e' $line[0] 'entites XML decodees'
    Assert-Equal '<' $line[1] '&lt;'
    # Decodage canonique en une passe : _x005F_ -> '_' puis '_x0041_' -> 'A'.
    # (Cette forme a double underscore n est de toute facon jamais produite par
    # Excel pour un litteral ; on epingle simplement le comportement.)
    Assert-Equal '_A' $line[2] 'forme non standard _x005F__x0041_'
    # Forme standard employee par Excel pour stocker le litteral "_x0041_".
    Assert-Equal '_x0041_' $line[3] 'litteral _x0041_ (forme standard Excel)'
    Assert-Equal ("tab" + [char] 9 + "ici") $line[4] 'tabulation _x0009_ decodee'
    Assert-Equal '_' $line[5] '_x005F_ seul = underscore'
}

Test-Case 'xlsx2csv-format-temps-ecoule' {
    $fx = Join-Path $script:CurrentDir 'elapsed.xlsx'
    New-XlsxFixture $fx -NumFmts @(
        @{ Id = 164; Code = '[h]:mm:ss' },
        @{ Id = 165; Code = '[mm]:ss' },
        @{ Id = 166; Code = '[s]' }
    ) -Xfs @(0, 164, 165, 166) -Sheets @(
        @{ Name = 'S1'; Rows = @( @{ Cells = @(
            @{ T = ''; V = '1.5';    S = 1 },   # 36 heures
            @{ T = ''; V = '0.0625'; S = 2 },   # 90 minutes
            @{ T = ''; V = '0.001';  S = 3 },   # 86,4 secondes
            @{ T = ''; V = '-0.5';   S = 1 }    # duree negative
        ) } ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'elapsed.csv') ';'
    Assert-Equal '36:00:00' $rows[0][0] 'format [h]:mm:ss : Excel affiche 36:00:00 pour 1,5 jour'
    Assert-Equal '90:00' $rows[0][1] 'format [mm]:ss : minutes totales'
    Assert-Equal '86' $rows[0][2] 'format [s] : secondes totales'
    Assert-Equal '-12:00:00' $rows[0][3] 'duree negative signee'
}

Test-Case 'xlsx2csv-dimension-erronee' {
    # dimension sous-estimee : la cellule C1 est hors dimension declaree
    $fx = Join-Path $script:CurrentDir 'dim.xlsx'
    New-XlsxFixture $fx -Sheets @(
        @{ Name = 'S1'; Dimension = 'A1:A1'; Rows = @(
            @{ R = 1; Cells = @( @{ R = 'A1'; T = ''; V = '1' }, @{ R = 'C1'; T = ''; V = '3' } ) }
        ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $text = [System.IO.File]::ReadAllText((Join-Path $script:CurrentDir 'dim.csv'))
    Assert-Equal ("1;;3`r`n") $text 'cellule au-dela de la dimension conservee'

    # Depassement sur une ligne tardive : l'export repart avec la largeur
    # mesuree, donc la ligne deja ecrite a elle aussi 3 champs.
    $sub = Join-Path $script:CurrentDir 'late'; New-Item -ItemType Directory -Path $sub | Out-Null
    $fx = Join-Path $sub 'late.xlsx'
    New-XlsxFixture $fx -Sheets @(
        @{ Name = 'S1'; Dimension = 'A1:A2'; Rows = @(
            @{ R = 1; Cells = @( @{ R = 'A1'; T = ''; V = '1' } ) },
            @{ R = 2; Cells = @( @{ R = 'A2'; T = ''; V = '2' }, @{ R = 'C2'; T = ''; V = '3' } ) }
        ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "ligne tardive exit code ($($r.Output))"
    $text = [System.IO.File]::ReadAllText((Join-Path $sub 'late.csv'))
    Assert-Equal ("1;;`r`n2;;3`r`n") $text 'toutes les lignes a la largeur reelle'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $sub '~late.csv.tmp'))) 'aucun fichier temporaire residuel'
}

Test-Case 'xlsx2csv-quoting-sortie' {
    $fx = Join-Path $script:CurrentDir 'q.xlsx'
    New-XlsxFixture $fx -Shared @('a;b', 'a"b', "multi`nline", ' lead ', 'plain') -Sheets @(
        @{ Name = 'S1'; Rows = @( @{ Cells = @(
            @{ T = 's'; V = '0' }, @{ T = 's'; V = '1' }, @{ T = 's'; V = '2' }, @{ T = 's'; V = '3' }, @{ T = 's'; V = '4' }
        ) } ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $text = [System.IO.File]::ReadAllText((Join-Path $script:CurrentDir 'q.csv'))
    $expected = '"a;b";"a""b";"multi' + "`n" + 'line";" lead ";plain' + "`r`n"
    Assert-Equal $expected $text 'quoting exact en sortie'
}

Test-Case 'xlsx2csv-encodages-ecriture' {
    $fx = Join-Path $script:CurrentDir 'e.xlsx'
    New-XlsxFixture $fx -Shared @("Ren`u{e9}", 'ok') -Sheets @(
        @{ Name = 'S1'; Rows = @( @{ Cells = @( @{ T = 's'; V = '0' }, @{ T = 's'; V = '1' } ) } ) }
    )
    $r = Invoke-Converter @($fx)
    Assert-Equal 0 $r.ExitCode "defaut ($($r.Output))"
    $bytes = [System.IO.File]::ReadAllBytes((Join-Path $script:CurrentDir 'e.csv'))
    Assert-True ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) 'defaut = UTF-8 BOM'
    Remove-Item (Join-Path $script:CurrentDir 'e.csv')

    $r = Invoke-Converter @($fx, '-Encoding', 'utf8')
    Assert-Equal 0 $r.ExitCode "utf8 ($($r.Output))"
    $bytes = [System.IO.File]::ReadAllBytes((Join-Path $script:CurrentDir 'e.csv'))
    Assert-True (-not ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB)) 'utf8 sans BOM'
    Remove-Item (Join-Path $script:CurrentDir 'e.csv')

    $r = Invoke-Converter @($fx, '-Encoding', 'ansi')
    Assert-Equal 0 $r.ExitCode "ansi ($($r.Output))"
    # 'ansi' ecrit dans la page de codes ANSI (cp1252 ici) sous les deux moteurs.
    $text = [System.IO.File]::ReadAllText((Join-Path $script:CurrentDir 'e.csv'), $script:EncAnsi)
    Assert-Contains $text "Ren`u{e9}" 'ansi relu en ANSI'
    Remove-Item (Join-Path $script:CurrentDir 'e.csv')

    $r = Invoke-Converter @($fx, '-Encoding', 'unicode')
    Assert-Equal 0 $r.ExitCode "unicode ($($r.Output))"
    $bytes = [System.IO.File]::ReadAllBytes((Join-Path $script:CurrentDir 'e.csv'))
    Assert-True ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) 'unicode = UTF-16 LE BOM'
}

Test-Case 'xlsx2csv-chaine-partagee-unique' {
    # Classeur avec EXACTEMENT une chaine partageee : la valeur doit sortir entiere.
    $fx = Join-Path $script:CurrentDir 'one.xlsx'
    New-XlsxFixture $fx -Shared @('Bonjour') -Sheets @(
        @{ Name = 'S1'; Rows = @( @{ Cells = @( @{ T = 's'; V = '0' }, @{ T = 'inlineStr'; V = 'x' } ) } ) }
    )
    $r = Invoke-Converter @($fx, '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'one.csv') ';'
    Assert-Equal 'Bonjour' $rows[0][0] 'chaine partagee unique lue en entier'
    Assert-Equal 'x' $rows[0][1] 'seconde cellule'
}

Test-Case 'xlsx2csv-feuille-vide' {
    $fx = Join-Path $script:CurrentDir 'empty.xlsx'
    New-XlsxFixture $fx -Sheets @( @{ Name = 'S1'; Rows = @(); OmitDimension = $true } )
    $r = Invoke-Converter @($fx)
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    Assert-True (Test-Path -LiteralPath (Join-Path $script:CurrentDir 'empty.csv')) 'csv cree'
    $bytes = [System.IO.File]::ReadAllBytes((Join-Path $script:CurrentDir 'empty.csv'))
    Assert-Equal 3 $bytes.Length 'csv vide = juste la BOM UTF-8'
    Assert-Equal '' ([System.IO.File]::ReadAllText((Join-Path $script:CurrentDir 'empty.csv'))) 'aucune ligne'
}

Test-Case 'xlsx2csv-language-fr-booleens' {
    $fx = Join-Path $script:CurrentDir 'b.xlsx'
    New-XlsxFixture $fx -Sheets @( @{ Name = 'S1'; Rows = @( @{ Cells = @( @{ T = 'b'; V = '1' }, @{ T = 'b'; V = '0' } ) } ) } )
    $r = Invoke-Converter @($fx, '-Delimiter', ';', '-Language', 'fr')
    Assert-Equal 0 $r.ExitCode "exit code ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'b.csv') ';'
    Assert-Equal 'VRAI' $rows[0][0] 'booleen vrai en francais'
    Assert-Equal 'FAUX' $rows[0][1] 'booleen faux en francais'
}

Test-Case 'xlsx2csv-outfile-et-nom-libre' {
    $fx = Join-Path $script:CurrentDir 'o.xlsx'
    New-XlsxFixture $fx -Sheets @( @{ Name = 'S1'; Rows = @( @{ Cells = @( @{ T = 'inlineStr'; V = 'x' } ) } ) } )
    $out = Join-Path $script:CurrentDir 'dest.csv'
    $r = Invoke-Converter @($fx, '-OutFile', $out)
    Assert-Equal 0 $r.ExitCode "OutFile ($($r.Output))"
    $r = Invoke-Converter @($fx, '-OutFile', $out)
    Assert-Equal 1 $r.ExitCode 'OutFile existant refuse'
    $r = Invoke-Converter @($fx, '-OutFile', $fx, '-Force')
    Assert-Equal 1 $r.ExitCode 'OutFile == source refuse'
    $r = Invoke-Converter @($fx)
    Assert-Equal 0 $r.ExitCode "auto ($($r.Output))"
    Assert-True (Test-Path -LiteralPath (Join-Path $script:CurrentDir 'o.csv')) 'o.csv'
    $r = Invoke-Converter @($fx)
    Assert-Equal 0 $r.ExitCode "auto 2 ($($r.Output))"
    Assert-True (Test-Path -LiteralPath (Join-Path $script:CurrentDir 'o (1).csv')) 'o (1).csv'
}

Test-Case 'xlsx2csv-formats-refuses' {
    # .xls/.xlsb recoivent le message dedie (enregistrer en XLSX), pas le generique.
    $fx = Join-Path $script:CurrentDir 'f.xls'
    [System.IO.File]::WriteAllText($fx, 'binaire')
    $r = Invoke-Converter @($fx)
    Assert-Equal 1 $r.ExitCode '.xls refuse'
    Assert-Contains $r.Output 'binary formats' 'message dedie pour .xls'
    Assert-True (-not $r.Output.Contains('Unsupported')) 'pas de message generique pour .xls'

    $fx = Join-Path $script:CurrentDir 'f.xlsb'
    [System.IO.File]::WriteAllText($fx, 'binaire')
    $r = Invoke-Converter @($fx)
    Assert-Equal 1 $r.ExitCode '.xlsb refuse'
    Assert-Contains $r.Output 'binary formats' 'message dedie pour .xlsb'
    Assert-True (-not $r.Output.Contains('Unsupported')) 'pas de message generique pour .xlsb'

    $fx = Join-Path $script:CurrentDir 'pasunzip.xlsx'
    [System.IO.File]::WriteAllText($fx, 'ceci n est pas un zip')
    $r = Invoke-Converter @($fx)
    Assert-Equal 1 $r.ExitCode 'zip invalide -> echec propre'

    $fx = Join-Path $script:CurrentDir 'sansworkbook.xlsx'
    New-XlsxFixture $fx -SkipWorkbook -Sheets @()
    $r = Invoke-Converter @($fx)
    Assert-Equal 1 $r.ExitCode 'zip sans workbook -> echec'
    Assert-Contains $r.Output 'workbook' 'message NotXlsx'
}

Test-Case 'xlsx2csv-xlsm-accepte' {
    $fx = Join-Path $script:CurrentDir 'macro.xlsm'
    New-XlsxFixture $fx -Sheets @( @{ Name = 'S1'; Rows = @( @{ Cells = @( @{ T = 'inlineStr'; V = 'ok' } ) } ) } )
    $r = Invoke-Converter @($fx)
    Assert-Equal 0 $r.ExitCode ".xlsm lu comme .xlsx ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $script:CurrentDir 'macro.csv') ';'
    Assert-Equal 'ok' $rows[0][0] 'contenu xlsm'
}

# ############################################################################
#  ALLER-RETOUR
# ############################################################################

Test-Case 'roundtrip-csv-xlsx-csv' {
    $dir1 = Join-Path $script:CurrentDir 'a'; New-Item -ItemType Directory -Path $dir1 | Out-Null
    $dir2 = Join-Path $script:CurrentDir 'b'; New-Item -ItemType Directory -Path $dir2 | Out-Null
    $csv = Join-Path $dir1 'rt.csv'
    Write-CsvFile $csv ("Nom;Entier;DecPoint;Zero;Long16;Date;DateHeure;Vide;Texte`r`nfoo;42;3.14;0012345;1234567890123456;2024-06-15;2024-06-15 13:45:30;;bar baz")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "csv->xlsx ($($r.Output))"
    Move-Item (Join-Path $dir1 'rt.xlsx') (Join-Path $dir2 'rt.xlsx')
    $r = Invoke-Converter @((Join-Path $dir2 'rt.xlsx'), '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "xlsx->csv ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $dir2 'rt.csv') ';'
    Assert-Equal 2 $rows.Count '2 lignes'
    $h = $rows[0]; $l = $rows[1]
    Assert-Equal 'Nom' $h[0] 'entete 1'
    Assert-Equal 'DateHeure' $h[6] 'entete 7'
    Assert-Equal 'foo' $l[0] 'texte'
    Assert-Equal '42' $l[1] 'entier'
    $expectedDec = $(if ($Cul.NumberFormat.NumberDecimalSeparator -eq ',') { '3,14' } else { '3.14' })
    Assert-Equal $expectedDec $l[2] 'decimal rendu selon la culture'
    Assert-Equal '0012345' $l[3] 'zero initial conserve'
    Assert-Equal '1234567890123456' $l[4] '16 chiffres conserve'
    Assert-Equal (Render-Date (OAD '2024-06-15 00:00:00')) $l[5] 'date aller-retour'
    Assert-Equal (Render-DateTime (OAD '2024-06-15 13:45:30')) $l[6] 'date-heure aller-retour'
    Assert-Equal '' $l[7] 'vide conserve'
    Assert-Equal 'bar baz' $l[8] 'texte avec espace'
}

Test-Case 'roundtrip-chaines-pieges' {
    $dir1 = Join-Path $script:CurrentDir 'a'; New-Item -ItemType Directory -Path $dir1 | Out-Null
    $dir2 = Join-Path $script:CurrentDir 'b'; New-Item -ItemType Directory -Path $dir2 | Out-Null
    $csv = Join-Path $dir1 'rt.csv'
    Write-CsvFile $csv ("H1;H2;H3;H4;H5`r`n""a;b"";""c""""d"";""multi`nline"";"" lead "";""<tag>&""")
    $r = Invoke-Converter @($csv)
    Assert-Equal 0 $r.ExitCode "csv->xlsx ($($r.Output))"
    Move-Item (Join-Path $dir1 'rt.xlsx') (Join-Path $dir2 'rt.xlsx')
    $r = Invoke-Converter @((Join-Path $dir2 'rt.xlsx'), '-Delimiter', ';')
    Assert-Equal 0 $r.ExitCode "xlsx->csv ($($r.Output))"
    $rows = Read-CsvAll (Join-Path $dir2 'rt.csv') ';'
    $l = $rows[1]
    Assert-Equal 'a;b' $l[0] 'separateur'
    Assert-Equal 'c"d' $l[1] 'guillemet'
    Assert-Equal ("multi`nline") $l[2] 'saut de ligne'
    Assert-Equal ' lead ' $l[3] 'espaces'
    Assert-Equal '<tag>&' $l[4] 'xml'
}

# ############################################################################
#  ERREURS GENERALES
# ############################################################################

Test-Case 'erreurs-fichiers' {
    $r = Invoke-Converter @((Join-Path $script:CurrentDir 'inexistant.csv'))
    Assert-Equal 1 $r.ExitCode 'fichier inexistant -> exit 1'
    Assert-Contains $r.Output 'not found' 'message NotFound'

    $doc = Join-Path $script:CurrentDir 'x.doc'
    [System.IO.File]::WriteAllText($doc, 'x')
    $r = Invoke-Converter @($doc)
    Assert-Equal 1 $r.ExitCode 'extension inconnue -> exit 1'
    Assert-Contains $r.Output 'Unsupported file type' 'message Unknown'

    $r = Invoke-Converter @()
    Assert-Equal 2 $r.ExitCode 'sans argument -> exit 2 (usage)'

    $r = Invoke-Converter @('-Version')
    Assert-Equal 0 $r.ExitCode '-Version -> exit 0'
    Assert-Contains $r.Output '1.0.1' 'version affichee'
}

Test-Case 'plusieurs-fichiers-en-une-passe' {
    $c1 = Join-Path $script:CurrentDir 'a.csv'
    $c2 = Join-Path $script:CurrentDir 'b.csv'
    $ko = Join-Path $script:CurrentDir 'ko.doc'
    Write-CsvFile $c1 "x`r`n1"
    Write-CsvFile $c2 "y`r`n2"
    [System.IO.File]::WriteAllText($ko, 'x')
    $r = Invoke-Converter @($c1, $c2, $ko)
    Assert-Equal 1 $r.ExitCode 'un echec -> exit 1'
    Assert-True (Test-Path -LiteralPath (Join-Path $script:CurrentDir 'a.xlsx')) 'a converti malgre echec de b... ko'
    Assert-True (Test-Path -LiteralPath (Join-Path $script:CurrentDir 'b.xlsx')) 'b converti'
}

# ############################################################################
#  BILAN
# ############################################################################

$pass  = @($script:Results | Where-Object Result -eq 'PASS').Count
$fail  = @($script:Results | Where-Object Result -eq 'FAIL').Count
$xfail = @($script:Results | Where-Object Result -eq 'XFAIL').Count
$xpass = @($script:Results | Where-Object Result -eq 'XPASS').Count
Write-Host ''
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "Moteur : $Engine   Culture : $($Cul.Name)" -ForegroundColor Cyan
Write-Host "Total : $($script:Results.Count)   PASS : $pass   FAIL : $fail   XFAIL (bogues connus confirmes) : $xfail   XPASS : $xpass" -ForegroundColor $(if ($fail) { 'Red' } else { 'Green' })
foreach ($f in ($script:Results | Where-Object Result -eq 'FAIL')) {
    Write-Host "  FAIL  $($f.Test) : $($f.Detail)" -ForegroundColor Red
}
foreach ($f in ($script:Results | Where-Object Result -eq 'XFAIL')) {
    Write-Host "  XFAIL $($f.Test) : $($f.Detail)" -ForegroundColor Yellow
}
foreach ($f in ($script:Results | Where-Object Result -eq 'XPASS')) {
    Write-Host "  XPASS $($f.Test) : $($f.Detail)" -ForegroundColor Magenta
}
Write-Host "Dossier de travail : $Work"

if (-not $Keep) {
    Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue
}

# Le code de sortie ne compte que les vraies regressions (FAIL).
exit [Math]::Min($fail, 1)
