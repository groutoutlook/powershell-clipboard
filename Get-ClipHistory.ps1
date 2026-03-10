<#
.SYNOPSIS
    Reads Windows clipboard history.

.DESCRIPTION
    Returns text entries directly and renders image entries with catimg.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [int[]]$Index,

    [Parameter()]
    [switch]$All,

    [Parameter()]
    [int]$PreviewCount = 20,

    [Parameter()]
    [int]$MaxPreviewLines = 8
)

. $PSScriptRoot\clipboard-runtime.ps1

[void](Invoke-WindowsPowerShellShim -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters)

$script:AnsiEscape = [char]27
$script:AnsiReset = "$($script:AnsiEscape)[0m"
$script:AnsiColors = @{
    Index   = "$($script:AnsiEscape)[96m"
    Image   = "$($script:AnsiEscape)[95m"
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

    if ($Entry.Type -eq 'Text') {
        Write-ClipboardText -EntryIndex $Entry.Index -Text (Get-ClipboardPreview -Text $Entry.Content -MaxLines $MaxPreviewLines) -IncludeIndex
        return
    }

    Write-Output "$(Format-ClipboardIndex -EntryIndex $Entry.Index) $(Format-ClipboardImageLabel)"
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
    $entries | ForEach-Object { Write-ClipboardEntrySummary -Entry $_ }
    return
}

if (-not $PSBoundParameters.ContainsKey('Index')) {
    $previewEntries = $entries | Select-Object -First $PreviewCount
    $previewEntries | ForEach-Object { Write-ClipboardEntrySummary -Entry $_ }

    if ($entries.Count -gt $previewEntries.Count) {
        $footer = "Showing $(Format-AnsiText -Text $previewEntries.Count -Color $script:AnsiColors.Accent) of $(Format-AnsiText -Text $entries.Count -Color $script:AnsiColors.Accent) clipboard entries. Use $(Format-AnsiText -Text '-All' -Color $script:AnsiColors.Index) to list everything."
        Write-Host (Format-AnsiText -Text $footer -Color $script:AnsiColors.Summary)
    } else {
        $footer = "Clipboard entries: $(Format-AnsiText -Text $entries.Count -Color $script:AnsiColors.Accent)"
        Write-Host (Format-AnsiText -Text $footer -Color $script:AnsiColors.Summary)
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
        Write-Host "$(Format-ClipboardIndex -EntryIndex $entry.Index) $(Format-ClipboardImageLabel)"
    }

    Show-ClipboardHistoryImage -DataPackageView $entry.Item.Content
}
