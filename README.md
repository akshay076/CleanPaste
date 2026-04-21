# CleanPaste 🧹

**Always-on clipboard monitor that auto-cleans terminal output when you copy.**

Copy from any terminal (Claude Code, Copilot CLI, Agency, Codex, PowerShell) and paste clean, formatted text into Word, Outlook, Teams, Excel, or any app. Tables become real bordered tables, broken paragraph wraps get re-joined, and bullet lists stay intact.

Only touches plain-text clipboard content from terminals — anything you copy from Word, Excel, browser, Teams, or any rich app passes through untouched.

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
- **Flowcharts** — ASCII box diagrams with arrows are not modified
- **Secrets** — tokens, passwords, API keys, PEM/SSH keys, certificates pass through untouched
- **Space-aligned tables** — columnar output from Agency, `Format-Table`, `kubectl`, `docker ps` stays intact
- **Rich content** — anything copied from Word, Excel, browser, Teams, or Outlook is never touched

## Pause / Resume

Right-click the CleanPaste system tray icon to pause or resume cleaning:

- **Pause CleanPaste** — temporarily stops cleaning (icon changes to ⚠️). Useful when you need to paste raw markdown into a GitHub PR or wiki editor.
- **Resume CleanPaste** — re-enables cleaning.
- **Exit CleanPaste** — stops the monitor entirely.

## How It Works

1. Detects clipboard changes via Win32 `GetClipboardSequenceNumber()` API (lightweight, no text hashing)
2. Skips rich clipboard content (HTML Format present = came from an app, not a terminal)
3. Normalizes terminal text (strips ANSI, prompts, line numbers)
4. Classifies content: secret? code? tree? columnar data? table? list? paragraph?
5. Renders cleanable content into HTML + plain text
6. Writes both formats to clipboard — rich apps get HTML, plain editors get text
7. Shows a system tray notification on each clean

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

- **Rich content passthrough** — clipboard content from Word, Excel, browsers, Teams, and other rich apps is never touched. Only plain-text terminal output is processed.
- **Confidence gating** — only rewrites clipboard when terminal artifacts are detected with high confidence (ANSI codes, box-drawing chars, shell prompts, terminal-width wrapping). Normal text, passwords, and tokens pass through untouched.
- **Secret protection** — GitHub PATs, JWTs, API keys, PEM/SSH private keys, certificates, connection strings, and high-entropy tokens are detected and never modified.
- **Short text protection** — single-line text under 50 characters is never modified, preventing accidental rewriting of passwords, tokens, or short commands.
- **HTML injection prevention** — all content is HTML-escaped before rendering. Untrusted clipboard text cannot inject links, images, or markup into your Word/Outlook documents.
- **No data collection** — everything runs locally. No network calls, no telemetry, no clipboard contents are ever logged or transmitted.
- **Safe prompt stripping** — only removes prompts with explicit path prefixes (`PS C:\>`, `C:\>`, `user@host:path$`). Bare `>` and `$` characters are preserved.
- **Idempotent** — running the cleaner on already-clean text produces identical output. No infinite rewrite loops.

## Known Limitations

- **Code formatting in Word** — code blocks paste as plain text since Word is not a code editor. Indentation is preserved but syntax highlighting is not.
- **Very large clipboard content** — content over 512KB is skipped to avoid blocking the monitor.
- **Non-text clipboard** — images, files, and other non-text clipboard content are ignored entirely.
- **Space-aligned tables** — columnar terminal output (e.g., `kubectl`, `Format-Table`) passes through as-is rather than being converted to bordered tables.

## Logs

Errors and cleaning events are logged to `~\.cleanpaste\cleanpaste.log` (auto-rotates at 1MB). No clipboard content is logged.

## Running Tests

```powershell
# Spec tests (68 tests)
powershell -NoProfile -ExecutionPolicy Bypass -File .\Tests.ps1

# Stress tests (227 tests)
powershell -NoProfile -ExecutionPolicy Bypass -File .\StressTest.ps1
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
