# Static validation for the CaretDownAligned repository.
# This does not compile VBA. It only checks source presence, encoding,
# headers, and regression patterns.

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$failures = New-Object System.Collections.Generic.List[string]

function Add-Check {
    param([string]$Name, [bool]$Condition)
    if ($Condition) {
        Write-Host "PASS: $Name"
    } else {
        Write-Host "FAIL: $Name"
        $script:failures.Add($Name)
    }
}

function Read-Utf8 {
    param([string]$Path)
    return [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
}

$basPath = Join-Path $root 'src/CaretDownAligned.bas'
$zhPath  = Join-Path $root 'src/CaretDownAligned.zh-CN.vba.txt'

$required = @(
    'README.md',
    'README.en.md',
    'LICENSE',
    'NOTICE.md',
    'PRIVACY.md',
    'DISCLAIMER.md',
    'SECURITY.md',
    'CONTRIBUTING.md',
    'CODE_OF_CONDUCT.md',
    'CHANGELOG.md',
    'VERSION',
    'src/CaretDownAligned.bas',
    'src/CaretDownAligned.zh-CN.vba.txt',
    'docs/INSTALL.md',
    'docs/FAQ.md',
    'docs/DESIGN.md',
    'docs/PUBLISH.md',
    'docs/LEGAL.zh-CN.md',
    'assets/banner.svg',
    '.github/workflows/validate.yml'
)

foreach ($rel in $required) {
    $path = Join-Path $root $rel
    Add-Check "exists: $rel" (Test-Path -LiteralPath $path)
}

if (-not (Test-Path -LiteralPath $basPath) -or -not (Test-Path -LiteralPath $zhPath)) {
    Write-Host ''
    Write-Host "Missing source files; cannot continue."
    exit 1
}

$bas = Read-Utf8 $basPath
$zh  = Read-Utf8 $zhPath
$basBytes = [IO.File]::ReadAllBytes($basPath)
$zhBytes  = [IO.File]::ReadAllBytes($zhPath)
$basFirstLine = (($bas -split "`r?`n") | Select-Object -First 1).Trim()

Add-Check "English .bas header" ($basFirstLine -eq 'Attribute VB_Name = "CaretDownAlignedModule"')
Add-Check "English .bas contains macro" ($bas.Contains('Sub CaretDownAlignedFinal'))
Add-Check "Chinese source contains macro" ($zh.Contains('Sub CaretDownAlignedFinal'))
Add-Check "Chinese source has grid helper" ($zh.Contains('Private Sub MakeGridLine'))
Add-Check "Chinese source has overtype helper" ($zh.Contains('Private Sub SetOvertypeForGrid'))
Add-Check "Chinese source has no Attribute line" (-not $zh.Contains('Attribute VB_Name'))
Add-Check "English .bas is ASCII" (-not ($basBytes | Where-Object { $_ -gt 127 }))
Add-Check "Chinese source is UTF-8 without BOM" (-not ($zhBytes.Length -ge 3 -and $zhBytes[0] -eq 0xEF -and $zhBytes[1] -eq 0xBB -and $zhBytes[2] -eq 0xBF))

$forbidden = @(
    @{ Name = 'no Paragraph.Index usage'; Pattern = 'Paragraph\.Index' },
    @{ Name = 'no curP.Index usage';      Pattern = 'curP\.Index' },
    @{ Name = 'no document-end InsertAfter'; Pattern = 'InsertAfter\s+Chr\(13\)' },
    @{ Name = 'no Range(p.Range.End) insert'; Pattern = 'Range\(p\.Range\.End' }
)
foreach ($item in $forbidden) {
    Add-Check ("English .bas: " + $item.Name) (-not ($bas -match $item.Pattern))
    Add-Check ("Chinese source: " + $item.Name) (-not ($zh -match $item.Pattern))
}

$license = Read-Utf8 (Join-Path $root 'LICENSE')
$readme  = Read-Utf8 (Join-Path $root 'README.md')
$readmeEn = Read-Utf8 (Join-Path $root 'README.en.md')
$version = (Read-Utf8 (Join-Path $root 'VERSION')).Trim()
$changelog = Read-Utf8 (Join-Path $root 'CHANGELOG.md')

Add-Check "LICENSE is MIT" ($license -match 'MIT License')
Add-Check "README mentions macro name" ($readme.Contains('CaretDownAlignedFinal'))
Add-Check "README links English .bas" ($readme.Contains('src/CaretDownAligned.bas'))
Add-Check "English README mentions macro name" ($readmeEn.Contains('CaretDownAlignedFinal'))
Add-Check "VERSION is 1.0.0" ($version -eq '1.0.0')
Add-Check "CHANGELOG contains 1.0.0" ($changelog.Contains('[1.0.0]'))

Write-Host ''
if ($failures.Count -gt 0) {
    Write-Host ("Validation failed: " + $failures.Count + " check(s).")
    exit 1
}

Write-Host 'All static checks passed.'
exit 0
