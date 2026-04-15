# CleanPaste 🧹

**Always-on clipboard monitor that auto-cleans terminal output when you copy.**

Copy from any terminal (Claude Code, Copilot CLI, Codex, PowerShell) and paste clean, formatted text into Word, Outlook, Teams, Excel, or any app. Tables become real bordered tables, broken paragraph wraps get re-joined, and bullet lists stay intact.

Strips ANSI escape codes, re-joins broken terminal line wraps, converts tables to HTML with borders, and preserves bullet lists — so you can paste clean, formatted text into any app.

## Quick Install (one-liner)

```powershell
git clone <your-repo-url> CleanPaste && cd CleanPaste && .\Install.ps1
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

1. Polls the clipboard every 300ms (lightweight, no noticeable CPU impact)
2. Parses content into **blocks**: paragraphs, tables, lists, blanks
3. Renders each block appropriately:
   - Paragraphs: re-joins broken terminal line wraps
   - Tables: converts to HTML clipboard format (renders with borders in Office)
   - Lists: preserves bullets, joins wrapped continuations
   - Code: passes through unchanged
4. Writes cleaned content back to clipboard
5. Shows a system tray notification on each clean

## Manual Usage

```powershell
# Run directly (foreground, Ctrl+C to stop)
.\CleanPaste.ps1

# Suppress tray notifications
.\CleanPaste.ps1 -NoBalloon

# Custom poll interval (ms)
.\CleanPaste.ps1 -PollMs 500
```

## Install / Uninstall

```powershell
# Install as startup task (runs at logon, hidden window)
.\Install.ps1

# Uninstall (removes task + files)
.\Install.ps1 -Uninstall
```

## Logs

Errors and cleaning events are logged to `~\.cleanpaste\cleanpaste.log` (auto-rotates at 1MB).

## Running Tests

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Tests.ps1
```

## Distribution

Share the repo and tell your team:

```
git clone <repo-url> CleanPaste && cd CleanPaste && .\Install.ps1
```

That's it — runs at login, no admin needed, no manual steps after install.

## Requirements

- Windows 10/11
- PowerShell 5.1+ (built-in)
- No admin rights needed

## License

MIT
