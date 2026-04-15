<#
.SYNOPSIS
    Test suite for CleanPaste block parser and cleaning logic.
#>

# ── Load functions from CleanPaste.ps1 ──────────────────────────────────────
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $scriptDir) { $scriptDir = $PWD.Path }
$scriptPath = Join-Path $scriptDir "CleanPaste.ps1"
$content = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$funcEnd = $content.IndexOf("# ── Tray icon setup")
$sb = [scriptblock]::Create($content.Substring(0, $funcEnd))
. $sb

# ── Test runner ──────────────────────────────────────────────────────────────

$script:pass = 0
$script:fail = 0

function Assert-Test([string]$name, [scriptblock]$test) {
    try {
        $result = & $test
        if ($result) {
            $script:pass++
            Write-Host "  [PASS] $name" -ForegroundColor Green
        } else {
            $script:fail++
            Write-Host "  [FAIL] $name" -ForegroundColor Red
        }
    } catch {
        $script:fail++
        Write-Host "  [FAIL] $name — $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ── Tests ────────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  CleanPaste Test Suite" -ForegroundColor Cyan
Write-Host "  ─────────────────────" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "  Block Parser" -ForegroundColor Yellow

Assert-Test "Blank lines create Blank blocks" {
    $blocks = Parse-Blocks "hello`n`nworld"
    ($blocks | Where-Object { $_.Type -eq 'Blank' }).Count -ge 1
}

Assert-Test "Table lines create Table blocks" {
    $blocks = Parse-Blocks "| A | B |`n|---|---|`n| 1 | 2 |"
    @($blocks | Where-Object { $_.Type -eq 'Table' }).Count -ge 1
}

Assert-Test "List lines create List blocks" {
    $blocks = Parse-Blocks "- item one`n- item two"
    @($blocks | Where-Object { $_.Type -eq 'List' }).Count -ge 1
}

Assert-Test "Mixed content creates multiple block types" {
    $mixed = "Some text`n`n| A | B |`n|---|---|`n| 1 | 2 |`n`n- bullet"
    $blocks = Parse-Blocks $mixed
    $types = ($blocks | ForEach-Object { $_.Type }) | Sort-Object -Unique
    ($types -contains 'Paragraph') -and ($types -contains 'Table') -and ($types -contains 'List')
}

Assert-Test "List continuation joins to parent bullet" {
    $blocks = Parse-Blocks "- item one that wraps`ncontinuation text`n- item two"
    $listBlock = $blocks | Where-Object { $_.Type -eq 'List' } | Select-Object -First 1
    $listBlock.Lines.Count -eq 3
}

Write-Host ""
Write-Host "  Paragraph Rendering" -ForegroundColor Yellow

Assert-Test "Joins broken lines into single paragraph" {
    $result = Render-ParagraphBlock @("The root cause was a missing", "index on the elastic table.")
    $result -eq "The root cause was a missing index on the elastic table."
}

Assert-Test "Collapses multiple spaces" {
    $result = Render-ParagraphBlock @("Hello    world")
    $result -eq "Hello world"
}

Write-Host ""
Write-Host "  List Rendering" -ForegroundColor Yellow

Assert-Test "Preserves separate bullet items" {
    $items = Render-ListBlock @("- first item", "- second item")
    $items.Count -eq 2
}

Assert-Test "Joins wrapped continuation to bullet" {
    $items = Render-ListBlock @("- long item that", "wraps to next line", "- second item")
    $items.Count -eq 2 -and $items[0] -like "*wraps to next line*"
}

Write-Host ""
Write-Host "  Table Conversion" -ForegroundColor Yellow

Assert-Test "Markdown table converts to HTML+TSV" {
    $result = Convert-TableToHtml "| A | B |`n|---|---|`n| 1 | 2 |"
    $result -is [hashtable] -and $result.Html -like "*<table*" -and $result.PlainText -like "*1*2*"
}

Assert-Test "Box-drawing table converts to HTML" {
    $tableText = "┌───┬───┐`n│ A │ B │`n├───┼───┤`n│ 1 │ 2 │`n└───┴───┘"
    $result = Convert-TableToHtml $tableText
    $result -is [hashtable] -and $result.Html -like "*<th>A</th>*"
}

Assert-Test "Separator rows are excluded from output" {
    $result = Convert-TableToHtml "| A | B |`n|---|---|`n| 1 | 2 |"
    -not ($result.PlainText -match '---')
}

Write-Host ""
Write-Host "  Content Detection" -ForegroundColor Yellow

Assert-Test "Markdown table detected as table" {
    Test-IsTable "| A | B |`n|---|---|`n| 1 | 2 |"
}

Assert-Test "Box table detected as table" {
    Test-IsTable "┌───┬───┐`n│ A │ B │`n└───┴───┘"
}

Assert-Test "Tree diagram detected as excluded" {
    Test-IsExcluded "src/`n├── file1`n│   ├── file2`n│   └── file3`n└── file4"
}

Assert-Test "Plain text not detected as table" {
    -not (Test-IsTable "Hello world. This is plain text.")
}

Assert-Test "KQL query detected as code" {
    $kql = "AppServiceHTTPLogs`n| where Status >= 500`n| summarize count() by bin(Time, 1h)`n| order by count_ desc"
    Test-IsCode $kql
}

Assert-Test "Plain paragraph not detected as code" {
    -not (Test-IsCode "The deployment was completed successfully. All services are running normally.")
}

Write-Host ""
Write-Host "  Full Pipeline (Invoke-CleanText)" -ForegroundColor Yellow

Assert-Test "Paragraph lines are joined" {
    $result = Invoke-CleanText "The root cause was a missing`r`nindex on the elastic table."
    $result.PlainText -like "*missing index*"
}

Assert-Test "Table returns hashtable with Html" {
    $result = Invoke-CleanText "| A | B |`n|---|---|`n| 1 | 2 |"
    $result -is [hashtable] -and $result.Html -like "*<table*"
}

Assert-Test "Mixed content preserves both paragraphs and tables" {
    $mixed = "First paragraph.`n`n| A | B |`n|---|---|`n| 1 | 2 |`n`nSecond paragraph."
    $result = Invoke-CleanText $mixed
    $result -is [hashtable] -and
    $result.Html -like "*First paragraph*" -and
    $result.Html -like "*<table*" -and
    $result.Html -like "*Second paragraph*"
}

Assert-Test "ANSI codes are stripped" {
    $result = Invoke-CleanText "$([char]0x1B)[32mSuccess$([char]0x1B)[0m"
    $result.PlainText -eq "Success"
}

Assert-Test "Code passes through untouched" {
    $code = "function test() {`n    const x = 1;`n    return x + 2;`n}"
    $result = Invoke-CleanText $code
    $result.ShouldRewrite -eq $false -and $result.PlainText -like "*function test()*"
}

Assert-Test "Shell prompts are stripped" {
    $result = Invoke-CleanText "PS C:\Users\me> Get-Process`nPS C:\Users\me> npm install"
    $result.PlainText -like "*Get-Process*" -and $result.PlainText -notlike "*PS C:*"
}

Assert-Test "Blank blocks emit &nbsp; paragraph in HTML" {
    $result = Invoke-CleanText "First para.`n`nSecond para."
    $result.Html -like "*&nbsp;*"
}

Assert-Test "HTML content is escaped in table cells" {
    $result = Convert-TableToHtml "| <script> | &test |`n|---|---|`n| val | val2 |"
    $result.Html -like "*&lt;script&gt;*" -and $result.Html -like "*&amp;test*"
}

Assert-Test "Bare > not stripped from text" {
    $result = Invoke-CleanText "> This is a quote`n> from an email"
    $result.PlainText -like "*> This is a quote*"
}

Assert-Test "Structured payload always returned" {
    $result = Invoke-CleanText "Simple text."
    $result -is [hashtable] -and $result.ContainsKey('Kind') -and $result.ContainsKey('ShouldRewrite')
}

# ── Results ──────────────────────────────────────────────────────────────────

$total = $script:pass + $script:fail
Write-Host ""
$color = if ($script:fail -eq 0) { "Green" } else { "Red" }
Write-Host "  ══════════════════════════════════" -ForegroundColor $color
Write-Host "  Results: $($script:pass) / $total PASSED" -ForegroundColor $color
if ($script:fail -gt 0) {
    Write-Host "  $($script:fail) test(s) FAILED" -ForegroundColor Red
}
Write-Host "  ══════════════════════════════════" -ForegroundColor $color
Write-Host ""

exit $script:fail
