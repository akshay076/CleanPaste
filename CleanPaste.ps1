<#
.SYNOPSIS
    CleanPaste - Always-on clipboard monitor that auto-cleans terminal output.

.DESCRIPTION
    Monitors the Windows clipboard and automatically cleans terminal output
    for pasting into Word, Outlook, Teams, and other Office apps. Converts
    tables to HTML with borders, re-joins broken paragraph wraps, and
    preserves bullet lists and code.

    Uses two-tier detection:
      Tier 1 (Structural): Tables, bullet lists → always clean
      Tier 2 (Artifacts):  ANSI codes, shell prompts, terminal wrapping → score-based

.NOTES
    Author : Akshay Gautam
    Version: 3.0.0
    License: MIT
#>

param(
    [switch]$NoBalloon,      # Suppress tray notifications
    [int]$PollMs = 300       # Clipboard check interval in milliseconds
)

# ── Version check ────────────────────────────────────────────────────────────
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "ERROR: CleanPaste requires PowerShell 5.1 or later." -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Red
    Write-Host "Install PowerShell 7: https://aka.ms/powershell" -ForegroundColor Yellow
    exit 1
}

# ── Imports ──────────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Win32 API for efficient clipboard change detection
try {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class ClipboardApi {
    [DllImport("user32.dll")]
    public static extern int GetClipboardSequenceNumber();
}
"@ -ErrorAction Stop
} catch {
    # Type may already be loaded in this session — that's fine
    if (-not ([System.Management.Automation.PSTypeName]'ClipboardApi').Type) {
        Write-Host "ERROR: Failed to load Win32 clipboard API: $($_.Exception.Message)" -ForegroundColor Red
        exit 1
    }
}

# ── State ────────────────────────────────────────────────────────────────────
$script:lastSeqNum     = 0
$script:cleanedCount   = 0
$script:running        = $true
$script:writtenSeqNum  = -1

# ── HTML escaping ────────────────────────────────────────────────────────────

function ConvertTo-HtmlSafe([string]$text) {
    return $text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
}

# ── Content detection ────────────────────────────────────────────────────────

function Remove-AnsiEscapes([string]$text) {
    return $text -replace '\x1B\[[0-9;]*[A-Za-z]', '' `
                 -replace '\x1B\].*?\x07', '' `
                 -replace '\x1B[()][AB012]', '' `
                 -replace '\x1B', ''
}

function Test-IsTable([string]$text) {
    $lines = ($text -split "`n") | Where-Object { $_.Trim() -ne '' }
    if ($lines.Count -ge 2) {
        # Markdown table: pipe-delimited data + separator row (supports 1+ columns)
        $hasSeparator = ($lines | Where-Object {
            $_ -match '^\s*\|?\s*[-:]+(\s*\|\s*[-:]+)*\s*\|?\s*$'
        }).Count -gt 0
        $pipeDataLines = ($lines | Where-Object { $_ -match '^\s*\|.+\|' }).Count
        if ($hasSeparator -and ($pipeDataLines -ge ($lines.Count * 0.5))) { return $true }
    }
    # Box-drawing table
    $hasCorners = $text -match '[┌┐└┘╔╗╚╝]'
    $hasHorizLine = $text -match '[─━═]'
    $allLines = $text -split "`n"
    $dataRows = ($allLines | Where-Object { $_ -match '^\s*[│║┃].+[│║┃]\s*$' }).Count
    if ($hasCorners -and $hasHorizLine -and $dataRows -ge 1) { return $true }
    return $false
}

function Test-IsExcluded([string]$text) {
    # Trees and flowcharts — should not be cleaned.
    # Box-drawing TABLES must NOT be excluded.
    $allLines = $text -split "`n"

    # File tree diagrams: lines with ├── or └── that are NOT table separators
    $treeLines = @($allLines | Where-Object {
        $_ -match '[├└]──' -and $_ -notmatch '[┼┤┴╪╡╧┬]'
    }).Count
    if ($allLines.Count -gt 2 -and $treeLines -ge 2) { return $true }

    # Flowcharts: arrows are a definitive signal
    $hasCorners = $text -match '[┌┐└┘╔╗╚╝]'
    if ($hasCorners) {
        $hasArrows = $text -match '──>' -or $text -match '[←→↑↓▶◀]'
        if ($hasArrows) { return $true }

        # Multiple top-left corners without table data rows = separate boxes (flowchart)
        $cornerCount = ([regex]::Matches($text, '[┌╔]')).Count
        $dataRows = @($allLines | Where-Object { $_ -match '^\s*[│║┃].+[│║┃]\s*$' }).Count
        if ($dataRows -ge 1) { return $false }
        if ($cornerCount -ge 2) { return $true }
    }

    return $false
}

function Test-IsCode([string]$text) {
    $lines = ($text -split "`n") | Where-Object { $_.Trim() -ne '' }
    $lineCount = $lines.Count
    if ($lineCount -eq 0) { return $false }
    # KQL/SQL pipe queries (exclude table rows)
    $pipeQueryLines = ($lines | Where-Object {
        $_ -cmatch '^\s*\|' -and $_ -notmatch '^\s*\|.*\|.*\|' -and $_ -notmatch '^\s*\|[-:\s]+\|'
    }).Count
    if ($lineCount -gt 2 -and $pipeQueryLines -ge ($lineCount * 0.3)) { return $true }
    # Code syntax density
    $codeChars = ([regex]::Matches($text, '[{}();=]|=>|->|::')).Count
    if ($codeChars / $lineCount -gt 0.5) { return $true }
    # Indented blocks
    $indentedLines = ($lines | Where-Object { $_ -match '^(\s{4,}|\t)' }).Count
    if ($indentedLines -ge ($lineCount * 0.6)) { return $true }
    return $false
}

function Test-IsSecret([string]$text) {
    # Detect tokens/passwords/keys/certs — these must never be modified.
    $trimmed = $text.Trim()

    # Multi-line secrets: PEM certificates, SSH keys, PGP blocks
    if ($trimmed -match '-----BEGIN\s+(RSA\s+)?(PRIVATE|PUBLIC|CERTIFICATE|OPENSSH|PGP|EC)') { return $true }

    # Bearer/auth tokens on their own
    if ($trimmed -match '(?i)^Bearer\s+[A-Za-z0-9\-_\.]+$') { return $true }

    # Connection strings
    if ($trimmed -match '(?i)(Server|Data Source|Host)=.*?(Password|Pwd)=') { return $true }

    [string[]]$lines = @(($trimmed -split "`n") | Where-Object { $_.Trim() -ne '' })

    # Multi-line base64 blob (e.g., cert body, key body)
    if ($lines.Count -ge 2 -and $lines.Count -le 50) {
        $base64Lines = @($lines | Where-Object { $_ -match '^[A-Za-z0-9+/=]{20,}$' }).Count
        if ($base64Lines -ge ($lines.Count * 0.8)) { return $true }
    }

    if ($lines.Count -gt 3) { return $false }

    $singleLine = if ($lines.Count -eq 1) { $lines[0].Trim() } else { $null }
    if ($singleLine) {
        # GitHub PAT
        if ($singleLine -match '^gh[ps]_[A-Za-z0-9]{20,}$') { return $true }
        # JWT
        if ($singleLine -match '^eyJ[A-Za-z0-9_\-\.]{20,}$') { return $true }
        # API key prefixes
        if ($singleLine -match '^(sk-|api-|key-|token-)[A-Za-z0-9\-_]{15,}$') { return $true }
        # Azure/AWS style keys
        if ($singleLine -match '^[A-Za-z0-9+/]{40,}={0,2}$') { return $true }
        # Generic: no spaces, 20+ chars, mix of upper/lower/digits
        if ($singleLine.Length -ge 20 -and $singleLine -notmatch '\s' -and
            $singleLine -match '[A-Z]' -and $singleLine -match '[a-z]' -and $singleLine -match '[0-9]') {
            return $true
        }
    }
    return $false
}

function Test-IsList([string]$text) {
    # Detect bullet or numbered lists
    $lines = ($text -split "`n") | Where-Object { $_.Trim() -ne '' }
    if ($lines.Count -lt 2) { return $false }
    $bulletLines = @($lines | Where-Object {
        $_ -match '^\s*[-*•●▶▪◦]\s+' -or $_ -match '^\s*\d+\.\s+'
    }).Count
    return $bulletLines -ge 2
}

# ── Cleaning functions ───────────────────────────────────────────────────────

function Remove-ShellPrompts([string]$text) {
    # Strip prompts with clear path/user prefix.
    # Do NOT strip bare > or $ to avoid damaging quoted email or code.
    $lines = $text -split "`n"
    $cleaned = foreach ($line in $lines) {
        $line -replace '^\s*PS\s+[A-Za-z]:\\[^>]*>\s*', '' `
              -replace '^\s*PS\s+/[^>]*>\s*', '' `
              -replace '^\s*[A-Z]:\\[^>]*>\s*', '' `
              -replace '^\s*\S+@\S+:[^\$#]*[\$#]\s*', ''
    }
    return $cleaned -join "`n"
}

function Test-HasPrompts([string]$text) {
    $lines = $text -split "`n"
    $promptLines = ($lines | Where-Object {
        $_ -match '^\s*PS\s+[A-Za-z]:[/\\]' -or
        $_ -match '^\s*[A-Z]:\\[^>]*>' -or
        $_ -match '^\s*\S+@\S+:[^\$#]*[\$#]\s'
    }).Count
    return $promptLines -ge 1
}

function Remove-LineNumbers([string]$text) {
    $lines = $text -split "`n"
    $pipe = [char]0x2502  # │
    $doublePipe = [char]0x2503  # ┃
    # Count lines that start with digits followed by a separator
    $numberedCount = ($lines | Where-Object {
        $trimmed = $_.TrimStart()
        $trimmed -match '^\d+\s*[\.\|]' -or $trimmed.Length -gt 2 -and [char]::IsDigit($trimmed[0]) -and ($trimmed.Contains($pipe) -or $trimmed.Contains($doublePipe))
    }).Count
    if ($lines.Count -gt 2 -and $numberedCount -ge ($lines.Count * 0.6)) {
        $cleaned = foreach ($line in $lines) {
            $line -replace ('^\s*\d+\s*[' + $pipe + $doublePipe + '\.\|]\s?'), ''
        }
        return $cleaned -join "`n"
    }
    return $text
}

function Convert-TableToHtml([string]$text) {
    $lines = $text -split "`n"
    $rows = [System.Collections.ArrayList]::new()

    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ($trimmed -eq '') { continue }
        if ($trimmed -match '^[┌┐└┘├┤┬┴┼╔╗╚╝╠╣╦╩╬─━═┈┉\s]+$') { continue }
        if ($trimmed -match '^\|?\s*[-:]+[-|:\s]+\s*\|?\s*$') { continue }

        if ($trimmed -match '[│║┃|]') {
            $cells = $trimmed -split '[│║┃|]' |
                     ForEach-Object { $_.Trim() } |
                     Where-Object { $_ -ne '' }
            if ($cells.Count -ge 1) {
                [void]$rows.Add($cells)
            }
        }
    }

    # Build HTML table with escaped content — skip if no data rows parsed
    if ($rows.Count -eq 0) {
        $tsv = ($lines | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) -join "`n"
        return @{ Html = $null; PlainText = $tsv }
    }

    $htmlParts = [System.Collections.ArrayList]::new()
    [void]$htmlParts.Add('<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">')
    $isFirst = $true
    foreach ($row in $rows) {
        $tag = if ($isFirst) { 'th' } else { 'td' }
        $cells = ($row | ForEach-Object { "<$tag>$(ConvertTo-HtmlSafe $_)</$tag>" }) -join ''
        [void]$htmlParts.Add("<tr>$cells</tr>")
        $isFirst = $false
    }
    [void]$htmlParts.Add('</table>')
    $html = $htmlParts -join ''

    # Plain text fallback (tab-separated)
    $tsv = ($rows | ForEach-Object { $_ -join "`t" }) -join "`n"

    return @{ Html = $html; PlainText = $tsv }
}

# ── Block parser ─────────────────────────────────────────────────────────────

function Parse-Blocks([string]$text) {
    $lines = $text -split "`n"
    $blocks = [System.Collections.ArrayList]::new()
    $currentLines = [System.Collections.ArrayList]::new()
    $currentType = $null

    foreach ($line in $lines) {
        $trimmed = $line.TrimEnd()
        $stripped = $trimmed.TrimStart()

        if ($stripped -eq '') {
            $lineType = 'Blank'
        } elseif (
            ($stripped -match '[│║┃].*[│║┃]') -or
            ($stripped -match '^\s*\|.+\|') -or
            ($stripped -match '^\|?\s*[-:]+[-|:\s]+[-:]+\s*\|?\s*$') -or
            ($stripped -match '^[┌┐└┘├┤┬┴┼╔╗╚╝╠╣╦╩╬─━═┈┉\s]+$')
        ) {
            $lineType = 'Table'
        } elseif ($stripped -match '^\s*[-*•●▶▪◦]\s+' -or $stripped -match '^\s*\d+\.\s+') {
            $lineType = 'List'
        } else {
            $lineType = 'Paragraph'
        }

        if ($lineType -eq 'Blank') {
            if ($currentLines.Count -gt 0) {
                [void]$blocks.Add(@{ Type = $currentType; Lines = [string[]]$currentLines.ToArray() })
                $currentLines = [System.Collections.ArrayList]::new()
            }
            [void]$blocks.Add(@{ Type = 'Blank'; Lines = @('') })
            $currentType = $null
            continue
        }

        if ($currentType -and $lineType -ne $currentType) {
            # Only join paragraph lines to list if they're indented (continuation)
            if ($currentType -eq 'List' -and $lineType -eq 'Paragraph' -and $trimmed -match '^\s{2,}\S') {
                [void]$currentLines.Add($trimmed)
                continue
            }
            [void]$blocks.Add(@{ Type = $currentType; Lines = [string[]]$currentLines.ToArray() })
            $currentLines = [System.Collections.ArrayList]::new()
        }

        $currentType = $lineType
        [void]$currentLines.Add($trimmed)
    }

    if ($currentLines.Count -gt 0) {
        [void]$blocks.Add(@{ Type = $currentType; Lines = [string[]]$currentLines.ToArray() })
    }

    return $blocks
}

function Render-ParagraphBlock([string[]]$lines) {
    $joined = ($lines | ForEach-Object { $_.Trim() }) -join ' '
    $joined = $joined -replace '[ \t]+', ' '
    return $joined.Trim()
}

function Render-ListBlock([string[]]$lines) {
    $items = [System.Collections.ArrayList]::new()
    $currentItem = ''

    foreach ($line in $lines) {
        $stripped = $line.Trim()
        if ($stripped -match '^\s*[-*•●▶▪◦]\s+' -or $stripped -match '^\s*\d+\.\s+') {
            if ($currentItem -ne '') { [void]$items.Add($currentItem) }
            $currentItem = $stripped
        } else {
            $currentItem = "$currentItem $stripped"
        }
    }
    if ($currentItem -ne '') { [void]$items.Add($currentItem) }
    return $items
}

# ── Main pipeline ────────────────────────────────────────────────────────────

function Get-CleanupPlan([string]$text) {
    # Single entry point: normalize → guard → classify → render → return.
    # Returns: @{ ShouldRewrite, Kind, Html, PlainText }

    $passthrough = @{ Kind = 'passthrough'; Html = $null; PlainText = $text; ShouldRewrite = $false }

    # ── 1. GUARD (pre-normalization) ─────────────────────────────────────
    if ([string]::IsNullOrWhiteSpace($text)) { return $passthrough }

    $lines = $text -split "`n"
    $nonEmpty = @($lines | Where-Object { $_.Trim() -ne '' })
    if ($nonEmpty.Count -le 1 -and $text.Trim().Length -lt 50) { return $passthrough }

    # ── 2. NORMALIZE ─────────────────────────────────────────────────────
    $hadAnsi = $text -match '\x1B\['
    $hadPrompt = Test-HasPrompts $text

    $normalized = Remove-AnsiEscapes $text
    $normalized = Remove-LineNumbers $normalized
    $normalized = Remove-ShellPrompts $normalized

    # ── 3. GUARD (post-normalization) ────────────────────────────────────
    if (Test-IsSecret $normalized) { return $passthrough }
    if (Test-IsExcluded $normalized) { return $passthrough }

    # ── 4. CLASSIFY ──────────────────────────────────────────────────────
    # Check structural content FIRST — tables/lists take priority over code detection
    # (a table cell containing code-like chars should still be treated as a table)
    $hasTable = Test-IsTable $normalized
    $hasList = Test-IsList $normalized
    $hasStructure = $hasTable -or $hasList

    if (-not $hasStructure -and (Test-IsCode $normalized)) {
        # Code passes through but we still strip ANSI/prompts if they were present
        if ($hadAnsi -or $hadPrompt) {
            return @{ Kind = 'code'; Html = $null; PlainText = $normalized.Trim(); ShouldRewrite = $true }
        }
        return $passthrough
    }

    # Terminal artifact score (on normalized text, plus flags from raw)
    $artifactScore = 0
    if ($hadAnsi) { $artifactScore += 4 }
    if ($hadPrompt) { $artifactScore += 3 }
    if ($normalized -match '[│┃╏╎▌┌┐└┘├┤┬┴┼╔╗╚╝╠╣╦╩╬═║─━]') { $artifactScore += 3 }

    # Terminal-width wrapping detection
    $normLines = @($normalized -split "`n" | Where-Object { $_.Trim() -ne '' })
    if ($normLines.Count -ge 2) {
        $lengths = $normLines | ForEach-Object { $_.TrimEnd().Length }
        $maxLen = ($lengths | Measure-Object -Maximum).Maximum
        $nearMax = @($lengths | Where-Object { $_ -ge ($maxLen - 10) }).Count
        $minWidth = if ($normLines.Count -eq 2) { 100 } else { 60 }
        if ($nearMax -ge 2 -and $maxLen -ge $minWidth) { $artifactScore += 3 }
    }

    $hasArtifacts = $artifactScore -ge 3

    if (-not $hasStructure -and -not $hasArtifacts) { return $passthrough }

    # ── 5. RENDER ────────────────────────────────────────────────────────
    $blocks = Parse-Blocks $normalized

    $htmlParts = [System.Collections.ArrayList]::new()
    $plainParts = [System.Collections.ArrayList]::new()
    $hasSemanticBlocks = $false

    foreach ($block in $blocks) {
        switch ($block.Type) {
            'Blank' {
                [void]$htmlParts.Add('<p style="margin:0">&nbsp;</p>')
                [void]$plainParts.Add('')
            }
            'Paragraph' {
                $para = Render-ParagraphBlock $block.Lines
                if ($para -ne '') {
                    $hasSemanticBlocks = $true
                    [void]$htmlParts.Add("<p style=`"margin:0`">$(ConvertTo-HtmlSafe $para)</p>")
                    [void]$plainParts.Add($para)
                }
            }
            'List' {
                $items = Render-ListBlock $block.Lines
                $hasSemanticBlocks = $true
                $isNumbered = $items[0] -match '^\s*\d+\.\s+'
                $listTag = if ($isNumbered) { 'ol' } else { 'ul' }
                $htmlItems = ($items | ForEach-Object {
                    $stripped = $_ -replace '^\s*[-*•●▶▪◦]\s+', '' -replace '^\s*\d+\.\s+', ''
                    "<li>$(ConvertTo-HtmlSafe $stripped)</li>"
                }) -join ''
                [void]$htmlParts.Add("<$listTag style=`"margin:0`">$htmlItems</$listTag>")
                [void]$plainParts.Add(($items -join "`n"))
            }
            'Table' {
                $tableResult = Convert-TableToHtml ($block.Lines -join "`n")
                if ($tableResult.Html) {
                    $hasSemanticBlocks = $true
                    [void]$htmlParts.Add($tableResult.Html)
                }
                [void]$plainParts.Add($tableResult.PlainText)
            }
        }
    }

    $plainText = ($plainParts -join "`n") -replace '(\r?\n){3,}', "`n`n"
    $plainText = $plainText.Trim()
    $htmlText = $htmlParts -join "`n"
    $finalHtml = if ($hasSemanticBlocks) { $htmlText } else { $null }

    # ── 6. FINAL CHECK ───────────────────────────────────────────────────
    # Rewrite if: text content changed OR HTML enrichment is available
    $textChanged = $plainText -ne $text.Trim()
    $hasHtmlEnrichment = $finalHtml -ne $null
    if (-not $textChanged -and -not $hasHtmlEnrichment) { return $passthrough }

    return @{
        Kind          = 'richtext'
        Html          = $finalHtml
        PlainText     = $plainText
        ShouldRewrite = $true
    }
}

# Backward-compatible wrapper (used by tests during transition)
function Invoke-CleanText([string]$text) {
    $plan = Get-CleanupPlan $text
    return $plan
}

# ── Clipboard operations ────────────────────────────────────────────────────

function Set-ClipboardRich([hashtable]$payload) {
    # Puts cleaned content on clipboard. Uses HTML format when available.
    if ($payload.Html) {
        $html = $payload.Html

        # Build CF_HTML format (Windows clipboard HTML standard)
        $preamble = "<!DOCTYPE html><html><head><meta charset=`"utf-8`"></head><body><!--StartFragment-->"
        $postamble = "<!--EndFragment--></body></html>"
        $body = "$preamble$html$postamble"

        $headerTemplate = "Version:0.9`r`nStartHTML:XXXXXXXXXX`r`nEndHTML:XXXXXXXXXX`r`nStartFragment:XXXXXXXXXX`r`nEndFragment:XXXXXXXXXX`r`n"
        $headerLen = [System.Text.Encoding]::UTF8.GetByteCount($headerTemplate)
        $sHtml = $headerLen
        $sFrag = $headerLen + [System.Text.Encoding]::UTF8.GetByteCount($preamble)
        $eFrag = $headerLen + [System.Text.Encoding]::UTF8.GetByteCount("$preamble$html")
        $eHtml = $headerLen + [System.Text.Encoding]::UTF8.GetByteCount($body)

        $header = "Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n" -f $sHtml, $eHtml, $sFrag, $eFrag
        $cfHtml = "$header$body"

        $dataObj = New-Object System.Windows.Forms.DataObject
        # CF_HTML must be UTF-8 bytes, not a .NET string (UTF-16), or multi-byte
        # characters like em-dashes get corrupted
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($cfHtml)
        $stream = New-Object System.IO.MemoryStream(,$bytes)
        $dataObj.SetData("HTML Format", $stream)
        $dataObj.SetData([System.Windows.Forms.DataFormats]::UnicodeText, $payload.PlainText)
        [System.Windows.Forms.Clipboard]::SetDataObject($dataObj, $true)
    } else {
        [System.Windows.Forms.Clipboard]::SetText($payload.PlainText)
    }
    return $payload.PlainText
}

# ── Logging ──────────────────────────────────────────────────────────────────

$script:logDir  = Join-Path $env:USERPROFILE ".cleanpaste"
$script:logFile = Join-Path $script:logDir "cleanpaste.log"

function Write-Log([string]$message, [string]$level = 'INFO') {
    try {
        if (-not (Test-Path $script:logDir)) {
            New-Item -ItemType Directory -Path $script:logDir -Force | Out-Null
        }
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $entry = "[$timestamp] [$level] $message"
        Add-Content -Path $script:logFile -Value $entry -ErrorAction SilentlyContinue

        # Rotate log if > 1MB
        if ((Test-Path $script:logFile) -and (Get-Item $script:logFile).Length -gt 1MB) {
            $backupLog = "$($script:logFile).bak"
            if (Test-Path $backupLog) { Remove-Item $backupLog -Force }
            Rename-Item $script:logFile $backupLog -Force
        }
    } catch {
        # Logging should never crash the monitor
    }
}

# ── Tray icon setup ─────────────────────────────────────────────────────────

$notifyIcon  = $null
$statusItem  = $null
if (-not $NoBalloon) {
    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
    $notifyIcon.Text = "CleanPaste - Clipboard Monitor"
    $notifyIcon.Visible = $true

    # Add right-click context menu
    $contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $exitItem = $contextMenu.Items.Add("Exit CleanPaste")
    $exitItem.Add_Click({ $script:running = $false })
    $statusItem = $contextMenu.Items.Add("Cleaned: 0 items")
    $statusItem.Enabled = $false
    $notifyIcon.ContextMenuStrip = $contextMenu
}

# ── Main loop ────────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "  ║         CleanPaste v3.0.0                ║" -ForegroundColor Cyan
Write-Host "  ║   Clipboard monitor is running...        ║" -ForegroundColor Cyan
Write-Host "  ║   Copy terminal text and it gets         ║" -ForegroundColor Cyan
Write-Host "  ║   cleaned automatically.                 ║" -ForegroundColor Cyan
Write-Host "  ║                                          ║" -ForegroundColor Cyan
Write-Host "  ║   Press Ctrl+C to stop.                  ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

Write-Log "CleanPaste v3.0.0 started (PollMs=$PollMs, NoBalloon=$NoBalloon)"

# Initialize clipboard sequence number
$script:lastSeqNum = [ClipboardApi]::GetClipboardSequenceNumber()

$script:clipFailCount = 0
$script:maxBackoffMs  = 5000

try {
    while ($script:running) {
        Start-Sleep -Milliseconds $PollMs

        try {
            # Efficient change detection via Win32 sequence number (no hashing)
            $currentSeqNum = [ClipboardApi]::GetClipboardSequenceNumber()
            if ($currentSeqNum -eq $script:lastSeqNum) { continue }
            $script:lastSeqNum = $currentSeqNum

            # Skip if this is the clipboard write we just made
            if ($currentSeqNum -eq $script:writtenSeqNum) { continue }

            # Check if clipboard has text
            if (-not [System.Windows.Forms.Clipboard]::ContainsText()) { continue }

            # Skip if clipboard already has rich HTML — content came from Word/browser/Outlook,
            # not a terminal. Only terminal output needs cleaning.
            $dataObj = [System.Windows.Forms.Clipboard]::GetDataObject()
            if ($dataObj -and $dataObj.GetDataPresent("HTML Format")) { continue }

            $currentText = [System.Windows.Forms.Clipboard]::GetText()
            if ([string]::IsNullOrWhiteSpace($currentText)) { continue }

            # Skip very large content to avoid blocking the loop
            if ($currentText.Length -gt 512000) { continue }

            # Reset backoff on successful read
            $script:clipFailCount = 0

            # Single-pass: normalize, classify, render
            $plan = Get-CleanupPlan $currentText
            if (-not $plan.ShouldRewrite) { continue }

            # Write cleaned content to clipboard
            $previewText = Set-ClipboardRich $plan
            # Track the sequence number AFTER successful write
            $script:writtenSeqNum = [ClipboardApi]::GetClipboardSequenceNumber()
            $script:lastSeqNum = $script:writtenSeqNum
            $script:cleanedCount++

            Write-Log "Cleaned #$($script:cleanedCount) ($($plan.Kind))"

            $preview = if ($previewText.Length -gt 80) { $previewText.Substring(0, 80) + "..." } else { $previewText }
            Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] Cleaned clipboard (#$($script:cleanedCount))" -ForegroundColor Green
            Write-Host "    Preview: $preview" -ForegroundColor DarkGray

            if ($notifyIcon) {
                $statusItem.Text = "Cleaned: $($script:cleanedCount) items"
                $notifyIcon.ShowBalloonTip(
                    2000,
                    "CleanPaste",
                    "Clipboard cleaned",
                    [System.Windows.Forms.ToolTipIcon]::Info
                )
            }
        }
        catch {
            $script:clipFailCount++
            $backoffMs = [Math]::Min($PollMs * [Math]::Pow(2, $script:clipFailCount), $script:maxBackoffMs)
            if ($script:clipFailCount -le 3 -or $script:clipFailCount % 10 -eq 0) {
                Write-Log "Clipboard error (attempt $($script:clipFailCount), backoff ${backoffMs}ms): $($_.Exception.Message)" 'WARN'
            }
            Start-Sleep -Milliseconds $backoffMs
        }
    }
}
finally {
    if ($notifyIcon) {
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
    }
    Write-Log "CleanPaste stopped. Cleaned $($script:cleanedCount) items total."
    Write-Host "`n  CleanPaste stopped. Cleaned $($script:cleanedCount) items total." -ForegroundColor Yellow
}
