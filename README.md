# PowerShell Clipboard Utils

Small scripts for reading Windows clipboard history and restoring items back into the active clipboard.

## Requirements

- Windows with Clipboard History enabled.
- Windows PowerShell 5.1. Running from PowerShell 7 is fine; the scripts automatically re-run in `powershell.exe`.
- [`catimg`](https://github.com/jiwonz/catimg) in `PATH` if you want image entries rendered in the terminal.

## Usage

List available history items:

```powershell
.\Get-ClipHistory.ps1 -All
```

Get a text item by index:

```powershell
.\Get-ClipHistory.ps1 -Index 0
```

Render an image item with `catimg`:

```powershell
.\Get-ClipHistory.ps1 -Index 3
```

Restore a history item into the current clipboard:

```powershell
.\Set-ClipHistory.ps1 -Index 1
```

Preview what would be restored without changing the clipboard:

```powershell
.\Set-ClipHistory.ps1 -Index 1 -WhatIf
```
