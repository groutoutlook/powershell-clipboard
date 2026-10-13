<#
.SYNOPSIS
    Reads Windows clipboard history.

.DESCRIPTION
    Returns text entries directly. Interactive image entries are rendered with
    imgcat/sixel/catimg. Plain path text is still Text. Explorer file handles are
    often hidden by Windows history behind Text-only views; only -File probes for
    those (temporarily), then restores the active clipboard.

    Selection modes:
      -Index N   absolute history slot (default positional)
      -Image N   Nth image entry only (0 = newest image)
      -File  N   Nth Explorer file-handle entry only (0 = newest file; probes)

    Use -Raw (or pipe through the gcbh shim) for machine-readable output:
      Text  -> text
      File  -> resolved path(s) when available
      Image -> temporary PNG path (for qrtool decode, etc.)

    The gcbh shim adds -Raw automatically when it is part of a pipeline, and
    leaves interactive stdout alone so images still go through the sixel viewer.
#>

[CmdletBinding(DefaultParameterSetName = 'ByIndex')]
param(
    [Parameter(Position = 0, ParameterSetName = 'ByIndex')]
    [int[]]$Index,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByImage')]
    [int[]]$Image,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByFile')]
    [int[]]$File,

    [Parameter()]
    [switch]$All,

    [Parameter()]
    [int]$PreviewCount = 20,

    [Parameter()]
    [int]$MaxPreviewLines = 8,

    [Parameter()]
    [switch]$Raw
)

. $PSScriptRoot\clipboard-runtime.ps1

[void](Invoke-WindowsPowerShellShim -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters)

$script:AnsiEscape = [char]27
$script:AnsiReset = "$($script:AnsiEscape)[0m"
$script:AnsiColors = @{
    Index   = "$($script:AnsiEscape)[96m"
    Image   = "$($script:AnsiEscape)[95m"
    File    = "$($script:AnsiEscape)[94m"
    Accent  = "$($script:AnsiEscape)[93m"
    Summary = "$($script:AnsiEscape)[92m"
}

function Format-AnsiText {
    param(
        [AllowEmptyString()]
        [string]$Text,
        [string]$Color
    )

    if ([string]::IsNullOrEmpty($Text) -or [string]::IsNullOrEmpty($Color)) {
        return $Text
    }

    return "$Color$Text$($script:AnsiReset)"
}

function Format-ClipboardIndex {
    param([int]$EntryIndex)

    return Format-AnsiText -Text "[$EntryIndex]" -Color $script:AnsiColors.Index
}

function Format-ClipboardImageLabel {
    return Format-AnsiText -Text '[Image]' -Color $script:AnsiColors.Image
}

function Format-ClipboardFileLabel {
    return Format-AnsiText -Text '[File]' -Color $script:AnsiColors.File
}

function Write-ClipboardText {
    param(
        [int]$EntryIndex,
        [string]$Text,
        [switch]$IncludeIndex
    )

    if ($IncludeIndex) {
        Write-Output "$(Format-ClipboardIndex -EntryIndex $EntryIndex) $Text"
        return
    }

    Write-Output $Text
}

function Write-ClipboardEntrySummary {
    param($Entry)

    switch ($Entry.Type) {
        'Text' {
            Write-ClipboardText -EntryIndex $Entry.Index -Text (Get-ClipboardPreview -Text $Entry.Content -MaxLines $MaxPreviewLines) -IncludeIndex
        }
        'File' {
            if ([string]::IsNullOrWhiteSpace($Entry.Content)) {
                Write-Output "$(Format-ClipboardIndex -EntryIndex $Entry.Index) $(Format-ClipboardFileLabel)"
            } else {
                $preview = Get-ClipboardPreview -Text $Entry.Content -MaxLines $MaxPreviewLines
                Write-Output "$(Format-ClipboardIndex -EntryIndex $Entry.Index) $(Format-ClipboardFileLabel) $preview"
            }
        }
        default {
            Write-Output "$(Format-ClipboardIndex -EntryIndex $Entry.Index) $(Format-ClipboardImageLabel)"
        }
    }
}

function Write-ClipboardEntryRaw {
    param($Entry)

    switch ($Entry.Type) {
        'Text' {
            Write-ClipboardText -EntryIndex $Entry.Index -Text $Entry.Content
        }
        'File' {
            if ($Entry.Paths -and $Entry.Paths.Count -gt 0) {
                $Entry.Paths | ForEach-Object { Write-Output $_ }
            } else {
                Write-ClipboardText -EntryIndex $Entry.Index -Text $Entry.Content
            }
        }
        'Image' {
            # ForTools: flatten alpha onto white and NN-upscale small images so
            # decoders like qrtool/rqrr can read soft clipboard QR bitmaps.
            $bytes = Get-ClipboardBitmapPngBytes -DataPackageView $Entry.Item.Content -ForTools
            if ($null -eq $bytes -or $bytes.Length -eq 0) {
                throw "Clipboard history index $($Entry.Index) did not contain readable bitmap data."
            }

            # Emit a temp PNG path so the value survives the powershell.exe shim and
            # works as `qrtool decode (gcbh -Image 0 -Raw)` or `gcbh -Image 0 | qrtool decode`.
            $tempPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("gcbh-{0}-{1}.png" -f $Entry.Index, [guid]::NewGuid().ToString('N'))
            [IO.File]::WriteAllBytes($tempPath, $bytes)
            Write-Output $tempPath
        }
        default {
            throw "Clipboard history index $($Entry.Index) contains $($Entry.Type), which -Raw cannot emit."
        }
    }
}

function Write-ClipboardEntryInteractive {
    param(
        $Entry,
        [int]$SelectionCount
    )

    switch ($Entry.Type) {
        'Text' {
            if ($SelectionCount -gt 1) {
                Write-ClipboardText -EntryIndex $Entry.Index -Text $Entry.Content -IncludeIndex
            } else {
                Write-ClipboardText -EntryIndex $Entry.Index -Text $Entry.Content
            }
        }
        'File' {
            $text = if ($Entry.Paths -and $Entry.Paths.Count -gt 0) {
                $Entry.Paths -join [Environment]::NewLine
            } else {
                $Entry.Content
            }

            if ($SelectionCount -gt 1) {
                Write-ClipboardText -EntryIndex $Entry.Index -Text $text -IncludeIndex
            } else {
                Write-ClipboardText -EntryIndex $Entry.Index -Text $text
            }
        }
        'Image' {
            if ($SelectionCount -gt 1) {
                Write-Host "$(Format-ClipboardIndex -EntryIndex $Entry.Index) $(Format-ClipboardImageLabel)"
            }

            Show-ClipboardHistoryImage -DataPackageView $Entry.Item.Content
        }
        default {
            Write-Error "Unsupported clipboard entry type '$($Entry.Type)' at index $($Entry.Index)."
        }
    }
}

try {
    # Do not probe hidden Explorer file handles here. That path calls
    # SetHistoryItemAsContent and can clobber the active clipboard; save/restore
    # of IDataObject is unreliable on this host. -File only sees StorageItems that
    # are already exposed on the history item.
    $entries = Get-ClipboardHistoryEntries
} catch {
    Write-Warning $_
    return
}

if ($entries.Count -eq 0) {
    return
}

if ($All) {
    $entries | ForEach-Object { Write-ClipboardEntrySummary -Entry $_ }
    return
}

$hasSelection = $PSCmdlet.ParameterSetName -ne 'ByIndex' -or $PSBoundParameters.ContainsKey('Index')
if (-not $hasSelection) {
    $previewEntries = $entries | Select-Object -First $PreviewCount
    $previewEntries | ForEach-Object { Write-ClipboardEntrySummary -Entry $_ }

    $imageCount = @($entries | Where-Object { $_.Type -eq 'Image' }).Count
    $fileCount = @($entries | Where-Object { $_.Type -eq 'File' }).Count

    if ($entries.Count -gt $previewEntries.Count) {
        $footer = "Showing $(Format-AnsiText -Text $previewEntries.Count -Color $script:AnsiColors.Accent) of $(Format-AnsiText -Text $entries.Count -Color $script:AnsiColors.Accent) clipboard entries ($imageCount images, $fileCount files). Use $(Format-AnsiText -Text '-All' -Color $script:AnsiColors.Index) to list everything."
        Write-Host (Format-AnsiText -Text $footer -Color $script:AnsiColors.Summary)
    } else {
        $footer = "Clipboard entries: $(Format-AnsiText -Text $entries.Count -Color $script:AnsiColors.Accent) ($imageCount images, $fileCount files)."
        Write-Host (Format-AnsiText -Text $footer -Color $script:AnsiColors.Summary)
    }

    return
}

$selected = @(Resolve-ClipboardSelection -Entries $entries -ParameterSetName $PSCmdlet.ParameterSetName -Index $Index -Image $Image -File $File)
foreach ($entry in $selected) {
    if ($Raw) {
        Write-ClipboardEntryRaw -Entry $entry
        continue
    }

    Write-ClipboardEntryInteractive -Entry $entry -SelectionCount $selected.Count
}
