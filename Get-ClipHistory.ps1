<#
.SYNOPSIS
    Reads Windows clipboard history.

.DESCRIPTION
    Returns text entries directly and renders image entries with catimg.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [int[]]$Index = 0,

    [Parameter()]
    [switch]$All,

    [Parameter()]
    [int]$MaxPreviewLines = 8
)

. $PSScriptRoot\clipboard-runtime.ps1

[void](Invoke-WindowsPowerShellShim -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters)

function Write-ClipboardText {
    param(
        [int]$EntryIndex,
        [string]$Text,
        [switch]$IncludeIndex
    )

    if ($IncludeIndex) {
        Write-Output "[$EntryIndex] $Text"
        return
    }

    Write-Output $Text
}

try {
    $entries = Get-ClipboardHistoryEntries
} catch {
    Write-Warning $_
    return
}

if ($entries.Count -eq 0) {
    return
}

if ($All) {
    $entries | ForEach-Object {
        if ($_.Type -eq 'Text') {
            Write-ClipboardText -EntryIndex $_.Index -Text (Get-ClipboardPreview -Text $_.Content -MaxLines $MaxPreviewLines) -IncludeIndex
        } else {
            Write-Output "[$($_.Index)] [Image]"
        }
    }
    return
}

foreach ($requestedIndex in $Index) {
    $entry = $entries | Where-Object { $_.Index -eq $requestedIndex } | Select-Object -First 1
    if ($null -eq $entry) {
        Write-Error "Index $requestedIndex out of range. Use -All to list available items."
        continue
    }

    if ($entry.Type -eq 'Text') {
        if ($Index.Count -gt 1) {
            Write-ClipboardText -EntryIndex $entry.Index -Text $entry.Content -IncludeIndex
            continue
        }

        Write-ClipboardText -EntryIndex $entry.Index -Text $entry.Content
        continue
    }

    if ($Index.Count -gt 1) {
        Write-Host "[$($entry.Index)] [Image]"
    }

    Show-ClipboardHistoryImage -DataPackageView $entry.Item.Content
}
