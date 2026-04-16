<#
.SYNOPSIS
    Test suite for CleanPaste v3 — spec-driven end-to-end tests.
    All tests target Get-CleanupPlan as the single entry point.
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

# ── Spec Tests (T01-T33): End-to-end via Get-CleanupPlan ─────────────────────

$esc = [char]0x1B

Write-Host ""
Write-Host "  CleanPaste v3 Test Suite" -ForegroundColor Cyan
Write-Host "  ════════════════════════" -ForegroundColor DarkCyan

# ── 6.1 Tables — MUST clean ─────────────────────────────────────────────────

Write-Host ""
Write-Host "  T01-T05: Tables" -ForegroundColor Yellow

Assert-Test "T01: Markdown table" {
    $r = Get-CleanupPlan "| Service | Port |`n|---------|------|`n| API | 8080 |"
    $r.ShouldRewrite -and $r.Html -like "*<table*" -and $r.PlainText -like "*API*8080*"
}

Assert-Test "T02: Box-drawing table (simple)" {
    $r = Get-CleanupPlan "┌───┬───┐`n│ A │ B │`n├───┼───┤`n│ 1 │ 2 │`n└───┴───┘"
    $r.ShouldRewrite -and $r.Html -like "*<table*" -and $r.Html -like "*<th>A</th>*"
}

Assert-Test "T03: Box-drawing table (large, real-world)" {
    $t = "┌────┬──────────┬──────────────────────────────────────────────────────────┬──────────┬────────────────┐`n"
    $t += "│ #  │ ADO ID   │ Title                                                    │ KPI      │ Assigned To    │`n"
    $t += "├────┼──────────┼──────────────────────────────────────────────────────────┼──────────┼────────────────┤`n"
    $t += "│ 13 │ 36963814 │ Fix ConvertTo-SecureString                                │ PS3.1    │ Jacob Allen    │`n"
    $t += "├────┼──────────┼──────────────────────────────────────────────────────────┼──────────┼────────────────┤`n"
    $t += "│ 15 │ 37439830 │ Upgrade instrumentationframework 3.5.1.1                  │ ES5.2    │ Unassigned     │`n"
    $t += "├────┼──────────┼──────────────────────────────────────────────────────────┼──────────┼────────────────┤`n"
    $t += "│ 16 │ 37439831 │ Upgrade .NET SDK 10.0                                     │ ES5.2    │ Unassigned     │`n"
    $t += "└────┴──────────┴──────────────────────────────────────────────────────────┴──────────┴────────────────┘"
    $r = Get-CleanupPlan $t
    $r.ShouldRewrite -and $r.Html -like "*<table*" -and $r.Html -like "*36963814*" -and $r.Html -like "*37439831*"
}

Assert-Test "T04: Single-column markdown table" {
    $r = Get-CleanupPlan "| Header |`n|--------|`n| Value  |"
    $r.ShouldRewrite -and $r.Html -like "*<table*"
}

Assert-Test "T05: Markdown table with extra whitespace" {
    $r = Get-CleanupPlan "| Name  | Age  |`n|-------|------|`n| Alice |  30  |`n| Bob   |  25  |"
    $r.ShouldRewrite -and $r.Html -like "*<table*" -and $r.Html -like "*Alice*"
}

# ── 6.2 Lists — MUST clean ──────────────────────────────────────────────────

Write-Host ""
Write-Host "  T06-T08: Lists" -ForegroundColor Yellow

Assert-Test "T06: Bullet list with wrapped continuations" {
    $list = "- First, we identified the root cause in the auth`n  service configuration`n- Second, we applied the hotfix to all production`n  clusters in East US and West EU"
    $r = Get-CleanupPlan $list
    $r.ShouldRewrite -and $r.Html -like "*<ul*" -and $r.PlainText -like "*root cause*service configuration*"
}

Assert-Test "T07: Numbered list with continuations" {
    $list = "1. Deploy the new configuration to staging`n   environment first`n2. Run smoke tests against the`n   staging endpoint"
    $r = Get-CleanupPlan $list
    $r.ShouldRewrite -and $r.Html -like "*<ol*"
}

Assert-Test "T08: Simple bullet list (HTML enrichment)" {
    $list = "- item one`n- item two`n- item three"
    $r = Get-CleanupPlan $list
    $r.ShouldRewrite -and $r.Html -like "*<ul*" -and $r.Html -like "*<li>*"
}

# ── 6.3 Terminal output — MUST clean ────────────────────────────────────────

Write-Host ""
Write-Host "  T09-T12: Terminal Output" -ForegroundColor Yellow

Assert-Test "T09: PowerShell prompt + output" {
    $t = "PS C:\Users\me> Get-Service | Select -First 3`n`nStatus   Name         DisplayName`n------   ----         -----------`nRunning  Appinfo      Application Information"
    $r = Get-CleanupPlan $t
    $r.ShouldRewrite -and $r.PlainText -notlike "*PS C:*" -and $r.PlainText -like "*Appinfo*"
}

Assert-Test "T10: Bash prompt + output" {
    $t = "user@host:~/project`$ ls -la`ntotal 32`ndrwxr-xr-x  5 user group 4096 Jan 15 10:30 .`ndrwxr-xr-x  3 user group 4096 Jan 15 10:30 .."
    $r = Get-CleanupPlan $t
    $r.ShouldRewrite -and $r.PlainText -notlike "*user@host*"
}

Assert-Test "T11: ANSI colored output" {
    $t = "${esc}[32m[SUCCESS]${esc}[0m Deployed auth-service v2.4.1`n${esc}[31m[ERROR]${esc}[0m Connection timeout to redis-cache-01"
    $r = Get-CleanupPlan $t
    $r.ShouldRewrite -and -not ($r.PlainText.Contains($esc)) -and $r.PlainText -like "*SUCCESS*" -and $r.PlainText -like "*ERROR*"
}

Assert-Test "T12: Terminal-wrapped paragraph (wide lines)" {
    # Simulate 120-char terminal wrapping
    $long = "The root cause of the outage was identified as a misconfigured load balancer health check that was incorrectly marking healthy instances as unhealthy. This caused the auto-scaler to terminate instances faster than they could be replaced, leading to a cascading failure across the East US region. The engineering team applied an emergency configuration change to restore the correct health check parameters and service was restored."
    # Wrap at 120 chars to simulate terminal
    $wrapped = ""
    for ($i = 0; $i -lt $long.Length; $i += 120) {
        $chunk = $long.Substring($i, [Math]::Min(120, $long.Length - $i))
        $wrapped += $chunk + "`n"
    }
    $r = Get-CleanupPlan $wrapped.TrimEnd()
    $r.ShouldRewrite -and $r.PlainText -like "*root cause*cascading failure*"
}

# ── 6.4 Pass-through — MUST NOT clean ──────────────────────────────────────

Write-Host ""
Write-Host "  T13-T24: Pass-through" -ForegroundColor Yellow

Assert-Test "T13: Normal short text" {
    $r = Get-CleanupPlan "Please review the attached document and let me know your thoughts."
    -not $r.ShouldRewrite
}

Assert-Test "T14: GitHub PAT" {
    $r = Get-CleanupPlan "ghp_xK9mP2qRvTnL8wJsYdBcUf3eAh6iOlN1"
    -not $r.ShouldRewrite
}

Assert-Test "T15: JWT token" {
    $r = Get-CleanupPlan "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9"
    -not $r.ShouldRewrite
}

Assert-Test "T16: Password" {
    $r = Get-CleanupPlan "P@ssw0rd!2024#Secure"
    -not $r.ShouldRewrite
}

Assert-Test "T17: API key" {
    $r = Get-CleanupPlan "sk-ant-api03-abc123def456ghi789"
    -not $r.ShouldRewrite
}

Assert-Test "T18: File tree diagram" {
    $r = Get-CleanupPlan "src/`n├── components`n│   ├── App.tsx`n│   └── Header.tsx`n└── index.ts"
    -not $r.ShouldRewrite
}

Assert-Test "T19: Flowchart" {
    $flow = "┌─────────┐     ┌─────────┐`n│  Start  │────>│   End   │`n└─────────┘     └─────────┘"
    $r = Get-CleanupPlan $flow
    -not $r.ShouldRewrite
}

Assert-Test "T20: KQL query" {
    $kql = "StormEvents`n| where StartTime between (datetime(2007-01-01) .. datetime(2007-12-31))`n| where State == `"FLORIDA`"`n| summarize count() by EventType"
    $r = Get-CleanupPlan $kql
    -not $r.ShouldRewrite
}

Assert-Test "T21: Python code" {
    $py = "def calculate(x, y):`n    result = x + y`n    return result`n`nprint(calculate(1, 2))"
    $r = Get-CleanupPlan $py
    -not $r.ShouldRewrite
}

Assert-Test "T22: JSON" {
    $json = "{`n  `"name`": `"test`",`n  `"version`": `"1.0`",`n  `"dependencies`": {`n    `"lodash`": `"^4.17`"`n  }`n}"
    $r = Get-CleanupPlan $json
    -not $r.ShouldRewrite
}

Assert-Test "T23: Short lines (intentionally formatted)" {
    $t = "Meeting notes:`nDiscussed roadmap`nAgreed on timeline`nNext steps pending`nAdjourned at 3pm"
    $r = Get-CleanupPlan $t
    -not $r.ShouldRewrite
}

Assert-Test "T24: Empty/whitespace" {
    $r = Get-CleanupPlan "   `n  `n   "
    -not $r.ShouldRewrite
}

# ── 6.5 Security ────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  T25-T26: Security" -ForegroundColor Yellow

Assert-Test "T25: HTML injection in table cells" {
    $r = Get-CleanupPlan "| <script>alert(1)</script> | x |`n|---|---|`n| val | y |"
    $r.ShouldRewrite -and $r.Html -like "*&lt;script&gt;*" -and $r.Html -notlike "*<script>*"
}

Assert-Test "T26: Prompt + token preserves token value" {
    $r = Get-CleanupPlan "PS C:\> echo ghp_xK9mP2qRvTnL8wJsYdBcUf3eAh6iOlN1"
    # The prompt gets stripped but the token value must be preserved
    $r.PlainText -like "*ghp_*"
}

# ── 6.6 Unicode ─────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  T27-T28: Unicode" -ForegroundColor Yellow

Assert-Test "T27: Em-dashes and accents" {
    $t = "The deployment $([char]0x2014) which took 3 hours $([char]0x2014) succeeded.`nR$([char]0xE9)sum$([char]0xE9): na$([char]0xEF)ve $([char]0x2192) optimized."
    $r = Get-CleanupPlan $t
    # If it gets cleaned (terminal wrapping), characters must survive
    $r.PlainText -like "*$([char]0x2014)*" -and $r.PlainText -like "*$([char]0xE9)*"
}

Assert-Test "T28: CJK in table" {
    $r = Get-CleanupPlan "| 名前 | 都市 |`n|------|------|`n| 太郎 | 東京 |"
    $r.ShouldRewrite -and $r.Html -like "*太郎*" -and $r.Html -like "*東京*"
}

# ── 6.7 Idempotency ─────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  T29-T30: Idempotency" -ForegroundColor Yellow

Assert-Test "T29: Already-clean paragraph" {
    $text = "The deployment was successful across all regions."
    $first = Get-CleanupPlan $text
    $second = Get-CleanupPlan $first.PlainText
    $first.PlainText -eq $second.PlainText
}

Assert-Test "T30: Cleaned table is stable" {
    $table = "| A | B |`n|---|---|`n| 1 | 2 |"
    $first = Get-CleanupPlan $table
    $second = Get-CleanupPlan $first.PlainText
    $first.PlainText -eq $second.PlainText
}

# ── 6.8 Edge Cases ──────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  T31-T33: Edge Cases" -ForegroundColor Yellow

Assert-Test "T31: Mixed content (paragraph + table + list)" {
    $mixed = "Summary paragraph.`n`n| A | B |`n|---|---|`n| 1 | 2 |`n`n- bullet one`n- bullet two"
    $r = Get-CleanupPlan $mixed
    $r.ShouldRewrite -and $r.Html -like "*<p*" -and $r.Html -like "*<table*" -and $r.Html -like "*<ul*"
}

Assert-Test "T32: Large table (200 rows)" {
    $header = "| ID | Name | Value |`n|----|------|-------|`n"
    $rows = (1..200 | ForEach-Object { "| $_ | item$_ | $($_ * 10) |" }) -join "`n"
    $r = Get-CleanupPlan "$header$rows"
    $r.ShouldRewrite -and ([regex]::Matches($r.Html, '<tr>')).Count -eq 201
}

Assert-Test "T33: Structural payload shape" {
    $r = Get-CleanupPlan "Simple text."
    $r -is [hashtable] -and $r.ContainsKey('Kind') -and $r.ContainsKey('ShouldRewrite') -and $r.ContainsKey('Html') -and $r.ContainsKey('PlainText')
}

Assert-Test "T34: Single box diagram not corrupted" {
    $box = "┌─────────┐`n│  Hello  │`n└─────────┘"
    $r = Get-CleanupPlan $box
    # Should either pass through or not produce empty table
    (-not $r.ShouldRewrite) -or ($r.Html -notlike "*<table*<tr></tr>*")
}

Assert-Test "T35: List does not swallow trailing paragraph" {
    $text = "- item one`n- item two`nConclusion paragraph here."
    $r = Get-CleanupPlan $text
    $r.PlainText -notlike "*two Conclusion*"
}

Assert-Test "T36: Bare dollar not stripped from docs" {
    $text = "`$ npm install`n`$ npm test`n`$ npm run build"
    $r = Get-CleanupPlan $text
    # Should not strip the $ or should pass through
    $r.PlainText -like "*npm install*" -and ($r.PlainText -like "*`$*" -or -not $r.ShouldRewrite)
}

Assert-Test "T37: PEM key passes through" {
    $pem = "-----BEGIN RSA PRIVATE KEY-----`nMIIEpAIBAAKCAQEA0Z3VS5JJcds3xfn/ygWyF8PbnGy`nZWFQS0NBUUVBMFozVlM1SkpjZHMzeGZu`n-----END RSA PRIVATE KEY-----"
    $r = Get-CleanupPlan $pem
    -not $r.ShouldRewrite
}

# ── Helper function tests ───────────────────────────────────────────────────

Write-Host ""
Write-Host "  Helper Functions" -ForegroundColor Yellow

Assert-Test "Test-IsSecret: GitHub PAT" { Test-IsSecret "ghp_xK9mP2qRvTnL8wJsYdBcUf3eAh6iOlN1" }
Assert-Test "Test-IsSecret: JWT" { Test-IsSecret "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9" }
Assert-Test "Test-IsSecret: API key" { Test-IsSecret "sk-ant-api03-abc123def456ghi789jkl" }
Assert-Test "Test-IsSecret: PEM private key" {
    Test-IsSecret "-----BEGIN RSA PRIVATE KEY-----`nMIIEpAIBAAKCAQEA0Z3VS5JJcds3xfn/ygWyF8PbnGy`n-----END RSA PRIVATE KEY-----"
}
Assert-Test "Test-IsSecret: SSH key" {
    Test-IsSecret "-----BEGIN OPENSSH PRIVATE KEY-----`nb3BlbnNzaC1rZXktdjEAAAAABG5vbmU`n-----END OPENSSH PRIVATE KEY-----"
}
Assert-Test "Test-IsSecret: connection string" {
    Test-IsSecret "Server=myserver.database.windows.net;Database=mydb;Password=s3cret123"
}
Assert-Test "Test-IsSecret: base64 blob" {
    Test-IsSecret "TUlJRXBBSUJBQUtDQVFFQTBaM1ZTNU`nKSmNkczN4Zm4veWdXeUY4UGJuR3kwWg`nWkZQS0NBUUVBMFozVlM1SkpjZHMzeGZu"
}
Assert-Test "Test-IsSecret: not a sentence" { -not (Test-IsSecret "This is a normal sentence") }
Assert-Test "Test-IsSecret: not multi-paragraph" { -not (Test-IsSecret "token123`nsome other text`nmore text`neven more") }

Assert-Test "Test-IsList: bullet list" { Test-IsList "- item one`n- item two`n- item three" }
Assert-Test "Test-IsList: numbered list" { Test-IsList "1. step one`n2. step two" }
Assert-Test "Test-IsList: not a list" { -not (Test-IsList "Just a paragraph of normal text.") }

Assert-Test "Test-IsExcluded: tree" { Test-IsExcluded "src/`n├── file1`n│   ├── file2`n│   └── file3`n└── file4" }
Assert-Test "Test-IsExcluded: flowchart" { Test-IsExcluded "┌─────┐     ┌─────┐`n│ A   │────>│ B   │`n└─────┘     └─────┘" }
Assert-Test "Test-IsExcluded: NOT box table" {
    -not (Test-IsExcluded "┌───┬───┐`n│ A │ B │`n├───┼───┤`n│ 1 │ 2 │`n└───┴───┘")
}

Assert-Test "Test-HasPrompts: PS prompt" { Test-HasPrompts "PS C:\Users\me> Get-Date" }
Assert-Test "Test-HasPrompts: bash prompt" { Test-HasPrompts "user@host:~/project`$ ls -la" }
Assert-Test "Test-HasPrompts: no prompt" { -not (Test-HasPrompts "Hello, this is normal text") }

Assert-Test "Remove-ShellPrompts: bash" {
    $r = Remove-ShellPrompts "user@host:~/project`$ ls -la"
    $r -like "*ls -la*" -and $r -notlike "*user@host*"
}

# ── Clipboard Round-Trip ────────────────────────────────────────────────────

Write-Host ""
Write-Host "  Clipboard Round-Trip" -ForegroundColor Yellow

Assert-Test "Set-ClipboardRich writes HTML+text for table" {
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
    $rawData = $dataObj.GetData("HTML Format")
    # May be a MemoryStream (UTF-8 bytes) or string depending on how it was set
    if ($rawData -is [System.IO.MemoryStream]) {
        $cfHtml = [System.Text.Encoding]::UTF8.GetString($rawData.ToArray())
    } else {
        $cfHtml = [string]$rawData
    }
    $cfHtml -like "*Version:0.9*" -and $cfHtml -like "*StartFragment*" -and $cfHtml -like "*<p>Hello</p>*"
}

Assert-Test "Set-ClipboardRich plain text fallback" {
    Add-Type -AssemblyName System.Windows.Forms
    $payload = @{ Kind = 'richtext'; Html = $null; PlainText = 'Just text'; ShouldRewrite = $true }
    Set-ClipboardRich $payload | Out-Null
    $result = [System.Windows.Forms.Clipboard]::GetText()
    $result -eq 'Just text'
}

# ── Block Parser ────────────────────────────────────────────────────────────

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

Assert-Test "List continuation joins to parent" {
    $blocks = Parse-Blocks "- item one that wraps`n  continuation text`n- item two"
    $listBlock = $blocks | Where-Object { $_.Type -eq 'List' } | Select-Object -First 1
    $listBlock.Lines.Count -eq 3
}

Assert-Test "Render-ParagraphBlock joins lines" {
    $result = Render-ParagraphBlock @("The root cause was a missing", "index on the elastic table.")
    $result -eq "The root cause was a missing index on the elastic table."
}

Assert-Test "Render-ListBlock preserves items" {
    $items = Render-ListBlock @("- first item", "- second item")
    $items.Count -eq 2
}

Assert-Test "Convert-TableToHtml escapes content" {
    $r = Convert-TableToHtml "| <b>Bold</b> | x |`n|---|---|`n| a | b |"
    $r.Html -like "*&lt;b&gt;*"
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
