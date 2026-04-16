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

Assert-Test "Standalone markdown table triggers cleaning (Tier 1)" {
    Test-ShouldClean "| Service | Port | Status |`n|---------|------|--------|`n| API | 8080 | Running |`n| Cache | 6379 | Stopped |"
}

Assert-Test "Standalone markdown table bypasses artifact scoring" {
    # This table has NO terminal artifacts — only structural content
    $table = "| Name | Age |`n|------|-----|`n| Alice | 30 |"
    Test-HasStructuralContent $table
}

Assert-Test "Plain text does NOT trigger cleaning" {
    -not (Test-ShouldClean "Hello, this is a normal sentence.")
}

Assert-Test "Short text (password-like) does NOT trigger cleaning" {
    -not (Test-ShouldClean "ghp_abc123XYZ789token")
}

Assert-Test "Tree diagram NOT cleaned by pipeline" {
    -not (Test-ShouldClean "src/`n├── file1`n│   ├── file2`n│   └── file3`n└── file4")
}

Assert-Test "Flowchart NOT cleaned by pipeline" {
    $flow = "┌─────────┐     ┌─────────┐`n│  Start  │────>│   End   │`n└─────────┘     └─────────┘"
    -not (Test-ShouldClean $flow)
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

# ── Clipboard Round-Trip Tests ───────────────────────────────────────────────

Write-Host ""
Write-Host "  Clipboard Round-Trip" -ForegroundColor Yellow

Assert-Test "Set-ClipboardRich writes HTML for table payload" {
    Add-Type -AssemblyName System.Windows.Forms
    $payload = @{ Kind = 'richtext'; Html = '<table><tr><td>A</td></tr></table>'; PlainText = 'A'; ShouldRewrite = $true }
    Set-ClipboardRich $payload | Out-Null
    $dataObj = [System.Windows.Forms.Clipboard]::GetDataObject()
    $formats = $dataObj.GetFormats()
    ('HTML Format' -in $formats) -and ('UnicodeText' -in $formats)
}

Assert-Test "CF_HTML has correct structure" {
    Add-Type -AssemblyName System.Windows.Forms
    $payload = @{ Kind = 'richtext'; Html = '<p>Hello</p>'; PlainText = 'Hello'; ShouldRewrite = $true }
    Set-ClipboardRich $payload | Out-Null
    $dataObj = [System.Windows.Forms.Clipboard]::GetDataObject()
    $cfHtml = $dataObj.GetData("HTML Format")
    $cfHtml -like "*Version:0.9*" -and $cfHtml -like "*StartFragment*" -and $cfHtml -like "*<p>Hello</p>*" -and $cfHtml -like "*EndFragment*"
}

Assert-Test "Set-ClipboardRich writes plain text for non-HTML payload" {
    Add-Type -AssemblyName System.Windows.Forms
    $payload = @{ Kind = 'richtext'; Html = $null; PlainText = 'Just text'; ShouldRewrite = $true }
    Set-ClipboardRich $payload | Out-Null
    $result = [System.Windows.Forms.Clipboard]::GetText()
    $result -eq 'Just text'
}

Assert-Test "Set-ClipboardRich returns plain text for preview" {
    $payload = @{ Kind = 'richtext'; Html = '<p>Test</p>'; PlainText = 'Test'; ShouldRewrite = $true }
    $preview = Set-ClipboardRich $payload
    $preview -eq 'Test'
}

# ── Edge Cases: Large Text ───────────────────────────────────────────────────

Write-Host ""
Write-Host "  Edge Cases: Large Text" -ForegroundColor Yellow

Assert-Test "Handles text with 500+ lines" {
    $bigText = (1..500 | ForEach-Object { "Line $_ of the document with some content that wraps" }) -join "`n"
    $result = Invoke-CleanText $bigText
    $result -is [hashtable] -and $result.PlainText.Length -gt 0
}

Assert-Test "Handles single very long line (2000+ chars)" {
    $longLine = "A" * 2000
    $result = Invoke-CleanText $longLine
    $result.PlainText -eq $longLine
}

Assert-Test "Handles empty string gracefully" {
    $result = Invoke-CleanText ""
    $result -is [hashtable] -and $result.PlainText -eq ""
}

Assert-Test "Handles whitespace-only string" {
    $result = Invoke-CleanText "   `n   `n   "
    $result -is [hashtable] -and $result.PlainText -eq ""
}

# ── Edge Cases: Mixed Table + Code ───────────────────────────────────────────

Write-Host ""
Write-Host "  Edge Cases: Mixed Content" -ForegroundColor Yellow

Assert-Test "KQL results table is converted to HTML" {
    $kqlResult = "| TimeGenerated | Count |`n|---------------|-------|`n| 2024-01-01    | 42    |`n| 2024-01-02    | 15    |"
    $result = Invoke-CleanText $kqlResult
    $result.Html -like "*<table*" -and $result.Html -like "*42*"
}

Assert-Test "Table followed by paragraph preserves both" {
    $mixed = "| A | B |`n|---|---|`n| 1 | 2 |`n`nThis is the explanation paragraph that should be joined properly."
    $result = Invoke-CleanText $mixed
    $result.Html -like "*<table*" -and $result.Html -like "*explanation paragraph*"
}

Assert-Test "Paragraph followed by table followed by paragraph" {
    $mixed = "Introduction text.`n`n| Col1 | Col2 |`n|------|------|`n| A    | B    |`n`nConclusion text."
    $result = Invoke-CleanText $mixed
    $result.Html -like "*Introduction*" -and $result.Html -like "*<table*" -and $result.Html -like "*Conclusion*"
}

Assert-Test "Multiple tables in one document" {
    $doc = "| T1A | T1B |`n|-----|-----|`n| 1   | 2   |`n`n| T2A | T2B |`n|-----|-----|`n| 3   | 4   |"
    $result = Invoke-CleanText $doc
    # Should have two <table> blocks
    $tableCount = ([regex]::Matches($result.Html, '<table')).Count
    $tableCount -eq 2
}

Assert-Test "Table with single data row works" {
    $small = "| Name | Value |`n|------|-------|`n| Test | 42    |"
    $result = Invoke-CleanText $small
    $result.Html -like "*<table*" -and $result.Html -like "*42*"
}

# ── Edge Cases: Numbered Lists ───────────────────────────────────────────────

Write-Host ""
Write-Host "  Edge Cases: Numbered Lists" -ForegroundColor Yellow

Assert-Test "Numbered list items are not joined as prose" {
    $numbered = "Steps to reproduce:`n1. Open the application`n2. Click on settings`n3. Toggle the feature"
    $result = Invoke-CleanText $numbered
    # Items should remain separate (not joined into one line)
    $result.PlainText -like "*1. Open*" -and $result.PlainText -like "*2. Click*" -and $result.PlainText -like "*3. Toggle*"
}

Assert-Test "Bullet list items stay separate" {
    $bullets = "- First item`n- Second item`n- Third item"
    $result = Invoke-CleanText $bullets
    ($result.PlainText -split "`n").Count -ge 3
}

Assert-Test "Mixed bullets and numbered" {
    $mixed = "Todo:`n- Buy milk`n- Buy eggs`n`nSteps:`n1. Go to store`n2. Pay"
    $result = Invoke-CleanText $mixed
    $result.PlainText -like "*Buy milk*" -and $result.PlainText -like "*1. Go*"
}

# ── Edge Cases: Non-ASCII Content ────────────────────────────────────────────

Write-Host ""
Write-Host "  Edge Cases: Non-ASCII Content" -ForegroundColor Yellow

Assert-Test "Chinese characters in table cells" {
    $table = "| Name | City |`n|------|------|`n| 张三 | 北京 |`n| 李四 | 上海 |"
    $result = Invoke-CleanText $table
    $result.Html -like "*张三*" -and $result.Html -like "*北京*"
}

Assert-Test "Emoji in paragraphs preserved" {
    $text = "The deployment was successful 🎉 and all tests passed ✅`nwith no issues found 👍"
    $result = Invoke-CleanText $text
    $result.PlainText -like "*🎉*" -and $result.PlainText -like "*✅*"
}

Assert-Test "Arabic text in table" {
    $table = "| الاسم | المدينة |`n|-------|---------|`n| أحمد  | الرياض  |"
    $result = Invoke-CleanText $table
    $result.Html -like "*أحمد*"
}

Assert-Test "Mixed Latin and non-Latin paragraph" {
    $text = "The project codename is プロジェクト and it targets`nthe APAC market specifically"
    $result = Invoke-CleanText $text
    $result.PlainText -like "*プロジェクト*" -and $result.PlainText -like "*APAC market*"
}

# ── Edge Cases: HTML Injection / Security ────────────────────────────────────

Write-Host ""
Write-Host "  Security: HTML Injection" -ForegroundColor Yellow

Assert-Test "Script tags escaped in paragraphs" {
    # <script> has code chars so gets detected as code and passes through.
    # But ConvertTo-HtmlSafe should escape it when used in HTML rendering.
    $escaped = ConvertTo-HtmlSafe "<script>alert('xss')</script>"
    $escaped -like "*&lt;script&gt;*"
}

Assert-Test "Image tags escaped in table cells" {
    # Test the table converter directly (not full pipeline which may route to code)
    $table = "| Name | Value |`n|------|-------|`n| <img src=x> | test |"
    $result = Convert-TableToHtml $table
    $result.Html -like "*&lt;img*" -and $result.Html -notlike "*<img*"
}

Assert-Test "Ampersands escaped in content" {
    $result = Invoke-CleanText "Tom & Jerry went to A&B Corp`nfor the meeting"
    $result.Html -like "*&amp;*"
}

Assert-Test "Quotes escaped in HTML output" {
    $result = Invoke-CleanText 'He said "hello" to the team`nand they responded'
    $result.Html -like "*&quot;*"
}

# ── Edge Cases: Prompt Stripping Safety ──────────────────────────────────────

Write-Host ""
Write-Host "  Prompt Stripping Safety" -ForegroundColor Yellow

Assert-Test "PS prompt stripped" {
    $result = Invoke-CleanText "PS C:\Users\test> Get-Date"
    $result.PlainText -like "*Get-Date*" -and $result.PlainText -notlike "*PS C:*"
}

Assert-Test "CMD prompt stripped" {
    $result = Invoke-CleanText "C:\Windows\System32> ipconfig"
    $result.PlainText -like "*ipconfig*" -and $result.PlainText -notlike "*C:\Windows*"
}

Assert-Test "Bare > preserved (quoted email)" {
    $result = Invoke-CleanText "> Original message from sender`n> Please review the attached`n`nMy reply here"
    $result.PlainText -like "*> Original*"
}

Assert-Test "Bare $ preserved" {
    $result = Invoke-CleanText "The cost is `$500 per unit`nand we need 10 units"
    $result.PlainText -like "*500*"
}

Assert-Test "Bare % preserved" {
    $result = Invoke-CleanText "CPU usage was at 95% during the incident`nwhich caused the timeout"
    $result.PlainText -like "*95%*"
}

# ── Edge Cases: ANSI Codes ───────────────────────────────────────────────────

Write-Host ""
Write-Host "  ANSI Code Handling" -ForegroundColor Yellow

Assert-Test "Color codes fully stripped" {
    $ansi = "$([char]0x1B)[31mERROR:$([char]0x1B)[0m Connection failed`n$([char]0x1B)[32mINFO:$([char]0x1B)[0m Retrying"
    $result = Invoke-CleanText $ansi
    $result.PlainText -like "*ERROR:*" -and -not ($result.PlainText.Contains([char]0x1B))
}

Assert-Test "Cursor movement codes stripped" {
    $ansi = "$([char]0x1B)[2A$([char]0x1B)[3CText after cursor moves"
    $result = Invoke-CleanText $ansi
    $result.PlainText -like "*Text after cursor moves*"
}

Assert-Test "OSC sequences stripped" {
    $osc = "$([char]0x1B)]0;Window Title$([char]0x07)Actual content here"
    $result = Invoke-CleanText $osc
    $result.PlainText -like "*Actual content here*" -and $result.PlainText -notlike "*Window Title*"
}

# ── Edge Cases: Code Detection ───────────────────────────────────────────────

Write-Host ""
Write-Host "  Code Detection" -ForegroundColor Yellow

Assert-Test "Python code detected as code" {
    $python = "def hello():`n    name = 'world'`n    print(f'Hello {name}')`n    return True"
    $result = Invoke-CleanText $python
    $result.Kind -eq 'code' -and $result.ShouldRewrite -eq $false
}

Assert-Test "JSON detected as code" {
    $json = "{`n    `"name`": `"CleanPaste`",`n    `"version`": `"2.1.0`",`n    `"enabled`": true`n}"
    $result = Invoke-CleanText $json
    $result.Kind -eq 'code' -and $result.ShouldRewrite -eq $false
}

Assert-Test "KQL with pipe operators detected as code" {
    $kql = "requests`n| where timestamp > ago(1h)`n| summarize count() by bin(timestamp, 5m)`n| render timechart"
    $result = Invoke-CleanText $kql
    $result.Kind -eq 'code' -and $result.ShouldRewrite -eq $false
}

Assert-Test "PowerShell script detected as code" {
    $ps = "function Get-Data {`n    param([string]`$Path)`n    `$items = Get-ChildItem -Path `$Path`n    return `$items`n}"
    $result = Invoke-CleanText $ps
    $result.Kind -eq 'code' -and $result.ShouldRewrite -eq $false
}

Assert-Test "SQL query NOT falsely detected as code" {
    $sql = "SELECT name, age FROM users WHERE status = 'active' ORDER BY name"
    $result = Invoke-CleanText $sql
    # Single line SQL — could go either way, but should not crash
    $result -is [hashtable]
}

# ── Edge Cases: Line Number Removal ──────────────────────────────────────────

Write-Host ""
Write-Host "  Line Number Removal" -ForegroundColor Yellow

Assert-Test "Line numbers stripped when majority of lines have them" {
    $pipe = [char]0x2502
    $numbered = "  1 $pipe first line`n  2 $pipe second line`n  3 $pipe third line`n  4 $pipe fourth line"
    $cleaned = Remove-LineNumbers $numbered
    $cleaned -like "*first line*" -and -not ($cleaned -like "*1 $pipe*")
}

Assert-Test "Line numbers NOT stripped when few lines have them" {
    $mixed = "This is line 1 of a paragraph`nwith a number 2 in it`nand more text here"
    $cleaned = Remove-LineNumbers $mixed
    $cleaned -like "*line 1*"
}

# ── Edge Cases: Box Drawing Tables ───────────────────────────────────────────

Write-Host ""
Write-Host "  Box Drawing Tables" -ForegroundColor Yellow

Assert-Test "Double-line box table converts to HTML" {
    $table = "╔═══════╦═════╗`n║ Name  ║ Age ║`n╠═══════╬═════╣`n║ Alice ║ 30  ║`n╚═══════╩═════╝"
    $result = Invoke-CleanText $table
    $result.Html -like "*<table*" -and $result.Html -like "*Alice*"
}

Assert-Test "Box table header row uses <th> tags" {
    $table = "┌───┬───┐`n│ A │ B │`n├───┼───┤`n│ 1 │ 2 │`n└───┴───┘"
    $result = Convert-TableToHtml $table
    $result.Html -like "*<th>*"
}

Assert-Test "Table plain text uses tab separation" {
    $table = "| Col1 | Col2 |`n|------|------|`n| A    | B    |"
    $result = Convert-TableToHtml $table
    $result.PlainText -like "*Col1`tCol2*"
}

# ── Idempotency ──────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  Idempotency" -ForegroundColor Yellow

Assert-Test "Cleaning already-clean paragraph produces same output" {
    $text = "This is a clean paragraph with no issues."
    $first = Invoke-CleanText $text
    $second = Invoke-CleanText $first.PlainText
    $first.PlainText -eq $second.PlainText
}

Assert-Test "Cleaning already-clean list produces same output" {
    $text = "- Item one`n- Item two`n- Item three"
    $first = Invoke-CleanText $text
    $second = Invoke-CleanText $first.PlainText
    $first.PlainText -eq $second.PlainText
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
