$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$scriptPath = Join-Path $scriptDir "CleanPaste.ps1"
$content = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$marker = "# " + [char]0x2500 + [char]0x2500 + " Tray icon setup"
$funcEnd = $content.IndexOf($marker)
. ([scriptblock]::Create($content.Substring(0, $funcEnd)))

# Reproduce: single-column table misidentified as KQL
$t = "| Header |`n|--------|`n| Value  |"
Write-Host "Input:"
Write-Host $t
Write-Host ""

Write-Host "Test-IsCode result: $(Test-IsCode $t)"
Write-Host "Test-HasStructuralContent result: $(Test-HasStructuralContent $t)"
Write-Host "Test-IsTable result: $(Test-IsTable $t)"
Write-Host ""

# Show which lines are being counted as KQL pipe lines
$lines = $t -split "`n"
foreach ($line in $lines) {
    $isKql = $line -cmatch '^\s*\|' -and
             $line -notmatch '^\s*\|.*\|.*\|' -and
             $line -notmatch '^\s*\|[-:\s]+\|'
    Write-Host "  '$line' -> KQL pipe line: $isKql"
}

Write-Host ""
Write-Host "Invoke-CleanText kind: $((Invoke-CleanText $t).Kind)"
