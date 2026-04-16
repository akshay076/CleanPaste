<#
.SYNOPSIS
    Stress test for CleanPaste - randomized and adversarial inputs.
#>

# Load CleanPaste functions
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$scriptPath = Join-Path $scriptDir "CleanPaste.ps1"
$content = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$marker = "# " + [char]0x2500 + [char]0x2500 + " Tray icon setup"
$funcEnd = $content.IndexOf($marker)
$sb = [scriptblock]::Create($content.Substring(0, $funcEnd))
. $sb

$pass = 0; $fail = 0; $errors = [System.Collections.ArrayList]::new()
$rng = [System.Random]::new(42)

function Test-Stress([string]$name, [scriptblock]$test) {
    try {
        $result = & $test
        if ($result) { $script:pass++ }
        else {
            $script:fail++
            [void]$script:errors.Add("FAIL: $name")
        }
    } catch {
        $script:fail++
        [void]$script:errors.Add("ERROR: $name -- $($_.Exception.Message)")
    }
}

function Get-RandomWord {
    $words = @("the","quick","brown","fox","jumps","over","lazy","dog","service","error",
               "deployment","failed","warning","timeout","exception","retry","success",
               "cluster","network","latency","threshold","metric","query","result","table")
    $words[$script:rng.Next($words.Count)]
}

function Get-RandomSentence([int]$wordCount = 10) {
    $w = 1..$wordCount | ForEach-Object { Get-RandomWord }
    $s = ($w -join " ")
    return $s.Substring(0,1).ToUpper() + $s.Substring(1) + "."
}

function Get-RandomParagraph([int]$sentences = 3) {
    (1..$sentences | ForEach-Object { Get-RandomSentence (5 + $rng.Next(8)) }) -join " "
}

Write-Host ""
Write-Host "  CleanPaste Stress Test" -ForegroundColor Cyan
Write-Host "  ──────────────────────" -ForegroundColor DarkCyan
Write-Host ""

# ── 1. Pipeline never crashes ─────────────────────────────────────────────────
Write-Host "  1. Pipeline Stability (100 random inputs)" -ForegroundColor Yellow

$chars = [char[]]([char[]]([char]0x20..[char]0x7E)) + [char[]]@(0x2502,0x251C,0x2514,0x2500,0x252C)

for ($i = 0; $i -lt 100; $i++) {
    $len = $rng.Next(1, 500)
    $str = -join (1..$len | ForEach-Object { $chars[$rng.Next($chars.Count)] })
    $captured = $str
    Test-Stress "Random printable string #$i" {
        $r = Invoke-CleanText $captured
        $r -is [hashtable] -and $r.ContainsKey('Kind') -and $r.ContainsKey('PlainText')
    }
}

# ── 2. Idempotency on random inputs ──────────────────────────────────────────
Write-Host "  2. Idempotency (50 random paragraphs)" -ForegroundColor Yellow

for ($i = 0; $i -lt 50; $i++) {
    $lines = 2..($rng.Next(3,8)) | ForEach-Object { Get-RandomSentence (6 + $rng.Next(6)) }
    $text = $lines -join "`n"
    $captured = $text
    Test-Stress "Idempotent paragraph #$i" {
        $first  = Invoke-CleanText $captured
        $second = Invoke-CleanText $first.PlainText
        $first.PlainText -eq $second.PlainText
    }
}

# ── 3. No data loss on clean prose ───────────────────────────────────────────
Write-Host "  3. No Data Loss (clean prose passthrough)" -ForegroundColor Yellow

for ($i = 0; $i -lt 30; $i++) {
    $text = Get-RandomSentence (8 + $rng.Next(5))
    $captured = $text
    Test-Stress "Single clean sentence #$i" {
        $r = Invoke-CleanText $captured
        $r.PlainText.Length -gt 0
    }
}

# ── 4. Control/special characters don't crash ────────────────────────────────
Write-Host "  4. Control and Special Characters" -ForegroundColor Yellow

$esc = [char]0x1B
$specialInputs = @(
    [string][char]0x08,
    ($esc + "[31m" + $esc + "[0m"),
    ("`t`t`t"),
    ("`r`n`r`n`r`n"),
    ([string]([char]0xFFFD)),
    ("=" * 200),
    ("|" * 50),
    ([string][char]0x2500 * 100),
    ($esc + "]0;Title" + [char]0x07 + "text"),
    ("PS C:\> " * 20),
    ("`n" * 50 + "hello" + "`n" * 50),
    ($esc + "[" + "0;" * 50 + "32mtext"),
    ("a" * 100000)
)

foreach ($inp in $specialInputs) {
    $captured = $inp
    $label = if ($inp.Length -gt 40) { $inp.Substring(0,40) + "..." } else { $inp -replace "`n","\\n" -replace "`t","\\t" }
    Test-Stress "Special chars: '$label'" {
        $r = Invoke-CleanText $captured
        $r -is [hashtable]
    }
}

# ── 5. False positive resistance ─────────────────────────────────────────────
Write-Host "  5. False Positive Resistance" -ForegroundColor Yellow

$shouldNotClean = @(
    "Please review the attached document",
    "Hi team, just a quick note about the meeting",
    "ghp_xK9mP2qRvTnL8wJsYdBcUf3eAh6iOlN1",
    "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCJ9",
    "correct horse battery staple",
    "127.0.0.1",
    "P@ssw0rd!2024#Secure",
    "sk-ant-api03-abc123",
    "user@example.com",
    "2024-01-15T10:30:00Z"
)

foreach ($t in $shouldNotClean) {
    $captured = $t
    Test-Stress "Not cleaned: '$t'" {
        -not (Test-ShouldClean $captured)
    }
}

# ── 6. Adversarial table content ─────────────────────────────────────────────
Write-Host "  6. Adversarial Table Content" -ForegroundColor Yellow

$adversarialTables = @(
    "| | |`n|---|---|`n| | |",
    "| Header |`n|--------|`n| Value  |",
    "| Key | Value |`n|-----|-------|`n| x | " + ("a" * 200) + " |",
    "| A | B |`n|---|---|`n| 1 | 2 |`n|---|---|`n| 3 | 4 |",
    "| Name | City |`n|------|------|`n| Jose | Madrid |",
    "| Col1 | Col2 | Col3 | Col4 | Col5 | Col6 | Col7 | Col8 | Col9 | Col10 |`n|------|------|------|------|------|------|------|------|------|-------|`n| a | b | c | d | e | f | g | h | i | j |"
)

foreach ($t in $adversarialTables) {
    $captured = $t
    Test-Stress "Adversarial table" {
        $r = Invoke-CleanText $captured
        $r -is [hashtable] -and $r.Html -like "*<table*"
    }
}

# ── 7. ANSI combinations ─────────────────────────────────────────────────────
Write-Host "  7. ANSI Combinations" -ForegroundColor Yellow

$ansiCombos = @(
    "$esc[1;31;40m Bold Red on Black $esc[0m normal",
    "$esc[38;2;255;128;0m Truecolor $esc[0m text",
    "$esc[?25l$esc[2J$esc[H visible text $esc[?25h",
    "$esc[A$esc[B$esc[C$esc[D cursor moves then text",
    (("$esc[32m") * 20 + "deep nesting" + ("$esc[0m") * 20),
    "line1`n$esc[31mline2$esc[0m`nline3`n$esc[32mline4$esc[0m"
)

foreach ($a in $ansiCombos) {
    $captured = $a
    Test-Stress "ANSI stripped cleanly" {
        $r = Invoke-CleanText $captured
        -not ($r.PlainText -match [char]0x1B)
    }
}

# ── 8. Large structured inputs ───────────────────────────────────────────────
Write-Host "  8. Large Structured Inputs" -ForegroundColor Yellow

$bigTable = "| ID | Name | Value | Status |`n|----|------|-------|--------|`n"
$bigTable += (1..200 | ForEach-Object { "| $_ | item$_ | $($rng.Next(1000)) | $(if($rng.Next(2)){'OK'}else{'FAIL'}) |" }) -join "`n"
Test-Stress "200-row table converts to HTML" {
    $r = Invoke-CleanText $bigTable
    $r.Html -like "*<table*" -and ([regex]::Matches($r.Html, '<tr>')).Count -eq 201
}

$bigList = (1..100 | ForEach-Object { "- $(Get-RandomSentence 5)" }) -join "`n"
Test-Stress "100-item bullet list renders" {
    $r = Invoke-CleanText $bigList
    $r -is [hashtable] -and $r.Html -like "*<ul*"
}

$mixed  = (Get-RandomParagraph 2) + "`n`n"
$mixed += "| Metric | Value |`n|--------|-------|`n| CPU | 92% |`n| Memory | 4.2GB |`n`n"
$mixed += (1..5 | ForEach-Object { "- $(Get-RandomSentence 6)" }) -join "`n"
$mixed += "`n`n" + (Get-RandomParagraph 2)
Test-Stress "Realistic mixed document (para + table + list + para)" {
    $r = Invoke-CleanText $mixed
    $r.Html -like "*<table*" -and $r.Html -like "*<ul*" -and $r.Html -like "*<p*"
}

# ── 9. Line number removal edge cases ────────────────────────────────────────
Write-Host "  9. Line Number Removal Edge Cases" -ForegroundColor Yellow

$pipe = [char]0x2502
$atThreshold  = (1..5 | ForEach-Object { "  $_ $pipe line content here" }) -join "`n"
$atThreshold += "`nextra line without number"
$belowThreshold = (1..3 | ForEach-Object { "  $_ $pipe line $_ content" }) -join "`n"
$belowThreshold += "`nnormal`nnormal`nnormal"

Test-Stress "Line numbers stripped at 83% coverage" {
    $r = Remove-LineNumbers $atThreshold
    $r -notlike "*1 $pipe*"
}
Test-Stress "Line numbers preserved at 50% coverage" {
    $r = Remove-LineNumbers $belowThreshold
    $r -like "*1 $pipe*"
}

# ── 10. HTML injection via random Unicode ────────────────────────────────────
Write-Host "  10. HTML Injection via Varied Input" -ForegroundColor Yellow

$injectionAttempts = @(
    "| <script>alert(1)</script> | x |`n|---|---|`n| val | y |",
    "| <img src=x onerror=alert(1)> | test |`n|---|---|`n| a | b |",
    "| javascript:void(0) | link |`n|---|---|`n| c | d |",
    "| `"onmouseover=`"alert(1) | x |`n|---|---|`n| e | f |",
    "Some text with <b>bold</b> and & ampersand`nstuff here wrapped"
)

foreach ($inj in $injectionAttempts) {
    $captured = $inj
    Test-Stress "No raw HTML in output" {
        $r = Invoke-CleanText $captured
        $html = $r.Html
        if (-not $html) { return $true }
        $html -notlike "*<script*" -and
        $html -notlike "*onerror=*" -and
        $html -notlike "*javascript:*" -and
        (-not ($html -match '<(?!/?(?:table|tr|th|td|ul|ol|li|p|br)\b)'))
    }
}

# ── Results ──────────────────────────────────────────────────────────────────
$total = $pass + $fail
Write-Host ""
$color = if ($fail -eq 0) { "Green" } else { "Red" }
Write-Host ("  " + "=" * 42) -ForegroundColor $color
Write-Host "  Stress Results: $pass / $total PASSED" -ForegroundColor $color
if ($fail -gt 0) {
    Write-Host "  $fail FAILED:" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "    $_" -ForegroundColor Red }
}
Write-Host ("  " + "=" * 42) -ForegroundColor $color
Write-Host ""

exit $fail
