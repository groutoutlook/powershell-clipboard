# PowerShell Clipboard Utils

Small scripts for reading Windows clipboard history and restoring items back into the active clipboard.

## Requirements

- Windows with Clipboard History enabled.
- Windows PowerShell 5.1. Running from PowerShell 7 is fine; the scripts automatically re-run in `powershell.exe`.
- A terminal image viewer in `PATH` (`imgcat`, `sixel`, or `catimg`) if you want image entries rendered in the terminal.

## Usage

Preview the first 20 history items:

```powershell
.\Get-ClipHistory.ps1
```

Change the preview length:

```powershell
.\Get-ClipHistory.ps1 -PreviewCount 30
```

List all available history items:

```powershell
.\Get-ClipHistory.ps1 -All
```

Get a text item by index:

```powershell
.\Get-ClipHistory.ps1 -Index 0
```

Render an image item in the terminal (imgcat/sixel/catimg):

```powershell
.\Get-ClipHistory.ps1 -Index 3
```

Select by kind instead of absolute history slot. Useful because text entries
push image indexes around. Ordinary `gcbh` listing/index/image reads do not
mutate the active clipboard. `-File` is the special case: Windows history often
hides Explorer file handles behind a Text-only view, so `-File` temporarily
probes via `SetHistoryItemAsContent`, then restores the previous clipboard
contents (path-looking text stays Text unless that probe yields a real drop):

```powershell
gcbh -Image 0              # newest image (interactive sixel/imgcat)
gcbh -Image 0 | qrtool decode
gcbh -File 0               # newest Explorer file-handle entry
qrtool decode (gcbh -Image 0 -Raw)
```

Emit machine-readable output for piping into other tools. Text stays text,
file handles become resolved path(s) when available, images become a temporary
PNG path. Raw image PNGs are flattened onto white and nearest-neighbor upscaled
when small so decoders such as `qrtool` can read soft clipboard QR bitmaps:

```powershell
.\Get-ClipHistory.ps1 -Index 3 -Raw
gcbh 3 | qrtool decode
qrtool decode (gcbh 3 -Raw)
```

The `gcbh` shim adds `-Raw` automatically when it is part of a pipeline, and
leaves interactive stdout alone so images still go through the sixel/imgcat
viewer. Preview lines label `[Image]` and `[File]` separately.

Restore a history item into the current clipboard:

```powershell
.\Set-ClipHistory.ps1 -Index 1
scbh -Image 0    # restore newest image as the active clipboard
scbh -File 0     # restore newest Explorer file-handle entry
```

Preview what would be restored without changing the clipboard:

```powershell
.\Set-ClipHistory.ps1 -Index 1 -WhatIf
scbh -Image 0 -WhatIf
```

Put a file or folder on the clipboard the same way Explorer does when you press Ctrl+C:

```powershell
.\set-fileClipboard.ps1 .\README.md
```

Mark a file or folder as a cut operation instead of copy:

```powershell
.\set-fileClipboard.ps1 .\README.md -Cut
```

Copy multiple paths in one call:

```powershell
.\set-fileClipboard.ps1 .\Get-ClipHistory.ps1, .\Set-ClipHistory.ps1
```

Reuse whatever is already on the clipboard when it contains copied files or valid text paths:

```powershell
.\set-fileClipboard.ps1
```

If the clipboard contains plain text, images, or other content that cannot be resolved to existing file system paths, the script stops with a clear error instead of trying to copy invalid data.
