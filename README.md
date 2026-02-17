---
Author: Mike Harris
Version: 0.0.2
---

# Outlook ToArchive Exporter

A lightweight PowerShell utility for **Classic Outlook (Windows)** that exports emails categorized with a specific Outlook category (e.g., `ToArchive`) to `.msg` files in a local folder (such as a OneDrive-synced directory).

The script is incremental, efficient, and designed to run automatically (e.g., hourly via Windows Task Scheduler).

## Overview

This tool:

- Scans selected Outlook folders
- Finds messages categorized with a configurable category (default: `ToArchive`)
- Exports matching messages to `.msg` files
- Organizes exports by received date
- Maintains a JSON index to prevent duplicate exports
- Optionally re-categorizes messages after successful export (e.g., `Archived`)

It is designed for personal or small-team use in environments running **Classic Outlook for Windows**.

## Requirements

- Windows 10 or Windows 11
- Classic Microsoft Outlook (COM-based Outlook, not "New Outlook")
- PowerShell 5.1+ (default on Windows)
- Outlook must be running when the script executes

This script does **not** support:
- New Outlook (WebView-based version)
- Outlook for Mac
- Outlook Web App (OWA)

## How It Works

1. You manually apply a category (default: `ToArchive`) to messages you want exported.
2. The script:
   - Uses Outlook COM automation
   - Filters messages using a DASL query
   - Saves each message as a Unicode `.msg` file
   - Writes metadata to `_index.json`
3. Exported messages are not re-exported on future runs.
4. Optionally, the script replaces `ToArchive` with `Archived`.

## Configuration

Edit the configuration block near the top of the script:

```powershell
$CategoryToArchive        = "ToArchive"
$CategoryArchived         = "Archived"
$AlsoRequireOlderThanDays = 0
$ExportRoot               = "$env:USERPROFILE\Documents\Outlook Archive"
$FoldersToScan            = @("Inbox","Inbox\Resolved","Sent Items","Archive","Deleted Items")
```

### Key Settings

**`$CategoryToArchive`**  
The Outlook category that marks messages for export.

**`$CategoryArchived`**  
Category applied after successful export. Set to `""` to disable.

**`$AlsoRequireOlderThanDays`**  
Set to a positive number to export only messages older than X days.

**`$ExportRoot`**  
Destination directory for exported `.msg` files.

**`$FoldersToScan`**  
List of folders to scan. Supports:
- `"Inbox"`
- `"Inbox\SubFolder"`
- `"Sent Items"`
- `"Deleted Items"`
- Any top-level folder
- Nested paths (e.g., `"Inbox\Resolved"`)

Set to `@()` to recursively scan the entire mailbox (slower). Custom folders or subfolders can be added to the `$FoldersToScan` object. As an example, `Inbox\Resolved` represents a folder named `Resolved` that is a subfolder of the main `Inbox` folder. 

## Output Structure

Exports are saved to the output path and organized by received date:

```
Outlook Archive/
    2026-02-10/
        2026-02-10 - Sender Name - Subject.msg
    2026-02-11/
        ...
    _index.json
    _export.log
```

### `_index.json`
Stores exported message identifiers to prevent duplicates.

### `_export.log`
Execution log with timestamps and folder summaries.

## Running Manually

Open PowerShell and run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\Export-ToArchive.ps1"
```

Outlook must already be open. If Outlook is not running, the script exits cleanly.

**Note:** Be sure to update the `\Path\To\` path to the appropriate location where the powershell script has been saved (e.g. `C:\Dev\Export-ToArchive.ps1`)

## Running Automatically (Recommended)

Use **Windows Task Scheduler** to run the script hourly/daily/weekly.

### Suggested Configuration

- Trigger: Daily
- Repeat task every: 1 hour
- For a duration of: Indefinitely
- Action:
  - Program: `powershell.exe`
  - Arguments:
    ```
    -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\Path\To\Export-ToArchive.ps1"
    ```
    **Note:** Be sure to update the `\Path\To\` path to match the file location.
  - Start in:
    ```
    C:\Path\To\
    ```
    **Note:** Be sure to update the `\Path\To\` path to match the file location.

Recommended settings:
- Run only when user is logged on
- Run with highest privileges
- Do not start a new instance if already running

## Design Considerations

- Uses `Items.Restrict()` for performance.
- Iterates backwards to avoid collection mutation issues.
- Uses `InternetMessageID` when available, falls back to `EntryID`.
- JSON index ensures idempotent exports.
- COM-based automation requires Outlook desktop client.

## Limitations

- Requires Classic Outlook.
- Outlook must be open.
- `.msg` format preserves full Outlook item fidelity but is Outlook-specific.
- Does not currently export attachments separately.
- Does not support shared mailboxes without modification.

## Security Considerations

This script:
- Does not transmit data externally.
- Writes files locally only.
- Uses local COM access to Outlook.

If used in corporate environments, ensure compliance with internal email retention policies. This script was originally developed to save important emails beyond the policy minimum retention time.

## License

This project is licensed under the MIT License. See the `LICENSE` file for details.

## Disclaimer

This tool is provided as-is, without warranty. Use at your own risk.
