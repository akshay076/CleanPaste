# CleanPaste 🧹

**Always-on clipboard monitor that auto-cleans terminal output when you copy.**

Copy from any terminal (Claude Code, Copilot CLI, Codex, PowerShell) and paste clean, formatted text into Word, Outlook, Teams, Excel, or any app. Tables become real bordered tables, broken paragraph wraps get re-joined, and bullet lists stay intact.

## Quick Install (one-liner)

```powershell
git clone https://github.com/akshay076/CleanPaste && cd CleanPaste && .\Install.ps1
```

Runs at every login. Install once, never think about it again.

## What It Does

| Copy this...             | Paste this...                         |
|--------------------------|---------------------------------------|
| Broken paragraph wraps   | Single clean paragraph                |
| Box-drawing tables `│┌─` | Real bordered table (HTML clipboard)  |
| Markdown tables `\|---\|`| Real bordered table (HTML clipboard)  |
| Bullet lists `- item`    | Preserved bullet list                 |
| ANSI color codes         | Clean plain text                      |
| Shell prompts `PS C:\>`  | Just the command                      |
| Line numbers `42 │`      | Just the content                      |

## What It Preserves

- **Code** — KQL queries, scripts, and code blocks pass through untouched
- **Tree diagrams** — `├── └──` file trees stay intact
- **Paragraph spacing** — blank lines in the original are preserved
- **Flowcharts** — ASCII box diagrams are not modified

## How It Works

1. Detects clipboard changes via Win32 `GetClipboardSequenceNumber()` API (lightweight, no text hashing)
2. Scores content for terminal artifacts — only rewrites when confidence is high (passwords/tokens pass through)
3. Parses content into **blocks**: paragraphs, tables, lists, blanks
4. Renders each block appropriately:
   - Paragraphs: re-joins broken terminal line wraps
   - Tables: converts to HTML clipboard format (renders with borders in Office)
   - Lists: preserves bullets, joins wrapped continuations
   - Code: passes through unchanged
5. Writes cleaned content back to clipboard
6. Shows a system tray notification on each clean

## Manual Usage

```powershell
# Run directly (foreground, Ctrl+C to stop)
.\CleanPaste.ps1

# Suppress tray notifications
.\CleanPaste.ps1 -NoBalloon
```

## Install / Uninstall

```powershell
# Install as startup task (runs at logon, hidden window)
.\Install.ps1

# Uninstall (removes task + files)
.\Install.ps1 -Uninstall
```

## Safety & Security

CleanPaste is designed to be safe for always-on use:

- **Confidence gating** — only rewrites clipboard when terminal artifacts are detected with high confidence (ANSI codes, box-drawing chars, shell prompts, terminal-width wrapping). Normal text, passwords, and tokens pass through untouched.
- **Short text protection** — single-line text under 50 characters is never modified, preventing accidental rewriting of passwords, tokens, or short commands.
- **HTML injection prevention** — all content is HTML-escaped before rendering. Untrusted clipboard text cannot inject links, images, or markup into your Word/Outlook documents.
- **No data collection** — everything runs locally. No network calls, no telemetry, no clipboard contents are ever transmitted.
- **Safe prompt stripping** — only removes prompts with explicit path prefixes (`PS C:\>`, `C:\>`). Bare `>` characters are preserved, so quoted email text and code are not damaged.
- **Idempotent** — running the cleaner on already-clean text produces identical output. No infinite rewrite loops.

## Known Limitations

- **Code formatting in Word** — code blocks paste as plain text since Word is not a code editor. Indentation is preserved but syntax highlighting is not.
- **Very large clipboard content** — no size cap currently; extremely large text (>1MB) may cause a brief delay.
- **Non-text clipboard** — images, files, and other non-text clipboard content are ignored entirely.

## Logs

Errors and cleaning events are logged to `~\.cleanpaste\cleanpaste.log` (auto-rotates at 1MB).

## Running Tests

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Tests.ps1
```

## Distribution

Share the repo and tell your team:

```
git clone https://github.com/akshay076/CleanPaste && cd CleanPaste && .\Install.ps1
```

That's it — runs at login, no admin needed, no manual steps after install.

## Requirements

- Windows 10/11
- PowerShell 5.1+ (built-in)
- No admin rights needed

## License

MIT
