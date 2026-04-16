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
    Version: 2.1.0
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
$script:skipNextChange = $false

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
        $hasSeparator = ($lines | Where-Object { $_ -match '^\s*\|?\s*[-:]+[-|:\s]+[-:]+\s*\|?\s*$' }).Count -gt 0
        $pipeDataLines = ($lines | Where-Object { $_ -match '^\s*\|.+\|' }).Count
        if ($hasSeparator -and ($pipeDataLines -ge ($lines.Count * 0.5))) { return $true }
    }
    $hasCorners = $text -match '[┌┐└┘╔╗╚╝]'
    $hasHorizLine = $text -match '[─━═]'
    $allLines = $text -split "`n"
    $dataRows = ($allLines | Where-Object { $_ -match '^\s*[│║┃].+[│║┃]\s*$' }).Count
    if ($hasCorners -and $hasHorizLine -and $dataRows -ge 1) { return $true }
    return $false
}

function Test-IsExcluded([string]$text) {
    # Trees and flowcharts — should not be cleaned
    $allLines = $text -split "`n"
    $junctionLines = ($allLines | Where-Object { $_ -match '[├└]──' }).Count
    if ($allLines.Count -gt 2 -and $junctionLines -ge 2) { return $true }
    $hasCorners = $text -match '[┌┐└┘╔╗╚╝]'
    if ($hasCorners -and ($text -match '[─━═│┃║]') -and -not (Test-IsTable $text)) { return $true }
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

function Test-HasStructuralContent([string]$text) {
    # Tier 1: Structural detection — we KNOW what this content is.
    # These should always be cleaned regardless of artifact score.
    $lines = ($text -split "`n") | Where-Object { $_.Trim() -ne '' }
    if ($lines.Count -lt 2) { return $false }

    # Markdown tables: | col | col | with a |---| separator
    $hasSeparator = @($lines | Where-Object { $_ -match '^\s*\|?\s*[-:]+[-|:\s]+[-:]+\s*\|?\s*$' }).Count -gt 0
    $pipeDataLines = @($lines | Where-Object { $_ -match '^\s*\|.+\|' }).Count
    if ($hasSeparator -and $pipeDataLines -ge 2) { return $true }

    # Box-drawing tables
    $hasCorners = $text -match '[┌┐└┘╔╗╚╝]'
    $hasHorizLine = $text -match '[─━═]'
    $dataRows = @($lines | Where-Object { $_ -match '^\s*[│║┃].+[│║┃]\s*$' }).Count
    if ($hasCorners -and $hasHorizLine -and $dataRows -ge 1) { return $true }

    return $false
}

function Test-HasTerminalArtifacts([string]$text) {
    # Tier 2: Artifact detection — we THINK this came from a terminal.
    # Score-based, needs >= 3 to trigger.
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }

    $score = 0
    $lines = $text -split "`n"
    $totalLines = $lines.Count

    # Short text (< 2 lines, < 50 chars) — skip to avoid mangling tokens/passwords
    if ($totalLines -le 1 -and $text.Length -lt 50) { return $false }

    # ANSI escape codes (strongest signal)
    if ($text -match '\x1B\[') { $score += 4 }

    # Box-drawing / pipe chars used as terminal wrapping
    if ($text -match '[│┃╏╎▌┌┐└┘├┤┬┴┼╔╗╚╝╠╣╦╩╬═║─━]') { $score += 3 }

    # Shell prompts (PS C:\>)
    $promptLines = ($lines | Where-Object { $_ -match '^\s*(PS\s+[A-Z]:\\[^>]*>)' }).Count
    if ($promptLines -ge 1) { $score += 3 }

    # Terminal-wrapped prose: lines at similar width (within 10 chars of max)
    # 2+ lines at ≥ 60 chars wide is a strong enough signal
    $nonEmpty = $lines | Where-Object { $_.Trim() -ne '' }
    if ($nonEmpty.Count -ge 2) {
        $lengths = $nonEmpty | ForEach-Object { $_.TrimEnd().Length }
        $maxLen = ($lengths | Measure-Object -Maximum).Maximum
        $nearMax = @($lengths | Where-Object { $_ -ge ($maxLen - 10) }).Count
        if ($nearMax -ge 2 -and $maxLen -ge 60) { $score += 3 }
    }

    return $score -ge 3
}

function Test-ShouldClean([string]$text) {
    # Two-tier detection model:
    #   Tier 1: Structural content (tables, etc.) → always clean
    #   Tier 2: Terminal artifacts (ANSI, prompts, wrapping) → score-based
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }

    # Tier 1: "I know what this is"
    if (Test-HasStructuralContent $text) { return $true }

    # Tier 2: "I think this came from a terminal"
    return Test-HasTerminalArtifacts $text
}

# ── Cleaning functions ───────────────────────────────────────────────────────

function Remove-ShellPrompts([string]$text) {
    # Only strip prompts that have a clear path prefix (PS C:\, C:\)
    # Do NOT strip bare > or $ to avoid damaging quoted email or code
    $lines = $text -split "`n"
    $cleaned = foreach ($line in $lines) {
        $line -replace '^\s*PS\s+[A-Z]:\\[^>]*>\s*', '' `
              -replace '^\s*[A-Z]:\\[^>]*>\s*', ''
    }
    return $cleaned -join "`n"
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
            if ($cells.Count -gt 1) {
                [void]$rows.Add($cells)
            }
        }
    }

    # Build HTML table with escaped content
    $html = '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">'
    $isFirst = $true
    foreach ($row in $rows) {
        $tag = if ($isFirst) { 'th' } else { 'td' }
        $cells = ($row | ForEach-Object { "<$tag>$(ConvertTo-HtmlSafe $_)</$tag>" }) -join ''
        $html += "<tr>$cells</tr>"
        $isFirst = $false
    }
    $html += '</table>'

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
            if ($currentType -eq 'List' -and $lineType -eq 'Paragraph') {
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

function Invoke-CleanText([string]$text) {
    # Returns a consistent structured payload:
    #   @{ Kind = 'richtext'|'code'|'passthrough'; Html = ...; PlainText = ...; ShouldRewrite = $true/$false }

    $result = Remove-AnsiEscapes $text
    $result = Remove-LineNumbers $result
    $result = Remove-ShellPrompts $result

    # Code passes through untouched
    if (Test-IsCode $result) {
        return @{
            Kind          = 'code'
            Html          = $null
            PlainText     = $result.Trim()
            ShouldRewrite = $false
        }
    }

    # Parse into blocks and render both HTML and plain text
    $blocks = Parse-Blocks $result

    $htmlParts = [System.Collections.ArrayList]::new()
    $plainParts = [System.Collections.ArrayList]::new()
    $hasSemanticBlocks = $false

    foreach ($block in $blocks) {
        switch ($block.Type) {
            'Blank' {
                # Emit an empty paragraph to preserve spacing in Word/Outlook
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
                # Detect numbered vs bullet list
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
                $hasSemanticBlocks = $true
                $tableResult = Convert-TableToHtml ($block.Lines -join "`n")
                [void]$htmlParts.Add($tableResult.Html)
                [void]$plainParts.Add($tableResult.PlainText)
            }
        }
    }

    $plainText = ($plainParts -join "`n") -replace '(\r?\n){3,}', "`n`n"
    $plainText = $plainText.Trim()
    $htmlText = $htmlParts -join "`n"

    return @{
        Kind          = 'richtext'
        Html          = if ($hasSemanticBlocks) { $htmlText } else { $null }
        PlainText     = $plainText
        ShouldRewrite = $true
    }
}

# ── Clipboard operations ────────────────────────────────────────────────────

function Set-ClipboardRich([hashtable]$payload) {
    # Puts cleaned content on clipboard. Uses HTML format when available.
    if ($payload.Html) {
        $html = $payload.Html

        # Build CF_HTML format (Windows clipboard HTML standard)
        $preamble = "<!DOCTYPE html><html><body><!--StartFragment-->"
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
        $dataObj.SetData("HTML Format", $cfHtml)
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

$notifyIcon = $null
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
Write-Host "  ║         CleanPaste v2.1.0                ║" -ForegroundColor Cyan
Write-Host "  ║   Clipboard monitor is running...        ║" -ForegroundColor Cyan
Write-Host "  ║   Copy terminal text and it gets         ║" -ForegroundColor Cyan
Write-Host "  ║   cleaned automatically.                 ║" -ForegroundColor Cyan
Write-Host "  ║                                          ║" -ForegroundColor Cyan
Write-Host "  ║   Press Ctrl+C to stop.                  ║" -ForegroundColor Cyan
Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

Write-Log "CleanPaste v2.1.0 started (PollMs=$PollMs, NoBalloon=$NoBalloon)"

# Initialize clipboard sequence number
$script:lastSeqNum = [ClipboardApi]::GetClipboardSequenceNumber()

$script:clipFailCount = 0
$script:maxBackoffMs  = 5000

try {
    while ($script:running) {
        Start-Sleep -Milliseconds $PollMs

        try {
            # Skip if we just wrote to the clipboard ourselves
            if ($script:skipNextChange) {
                $script:skipNextChange = $false
                $script:lastSeqNum = [ClipboardApi]::GetClipboardSequenceNumber()
                continue
            }

            # Efficient change detection via Win32 sequence number (no hashing)
            $currentSeqNum = [ClipboardApi]::GetClipboardSequenceNumber()
            if ($currentSeqNum -eq $script:lastSeqNum) { continue }
            $script:lastSeqNum = $currentSeqNum

            # Check if clipboard has text
            if (-not [System.Windows.Forms.Clipboard]::ContainsText()) { continue }

            $currentText = [System.Windows.Forms.Clipboard]::GetText()
            if ([string]::IsNullOrWhiteSpace($currentText)) { continue }

            # Reset backoff on successful read
            $script:clipFailCount = 0

            # Two-tier detection: structural content OR terminal artifacts
            if (-not (Test-ShouldClean $currentText)) { continue }

            # Clean the text
            $payload = Invoke-CleanText $currentText

            # Respect the payload's own decision
            if (-not $payload.ShouldRewrite) { continue }

            # Skip if plain text didn't actually change
            if ($payload.PlainText -eq $currentText) { continue }

            # Write cleaned content to clipboard
            $script:skipNextChange = $true
            $previewText = Set-ClipboardRich $payload
            $script:cleanedCount++

            Write-Log "Cleaned #$($script:cleanedCount) ($($payload.Kind)): $($previewText.Substring(0, [Math]::Min(80, $previewText.Length)))"

            $preview = if ($previewText.Length -gt 80) { $previewText.Substring(0, 80) + "..." } else { $previewText }
            Write-Host "  [$(Get-Date -Format 'HH:mm:ss')] Cleaned clipboard (#$($script:cleanedCount))" -ForegroundColor Green
            Write-Host "    Preview: $preview" -ForegroundColor DarkGray

            if ($notifyIcon) {
                $statusItem.Text = "Cleaned: $($script:cleanedCount) items"
                $notifyIcon.ShowBalloonTip(
                    2000,
                    "CleanPaste",
                    "Clipboard cleaned (#$($script:cleanedCount))",
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
