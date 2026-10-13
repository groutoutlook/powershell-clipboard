<#
.SYNOPSIS
    Restores a Windows clipboard history item into the active clipboard.

.DESCRIPTION
    Selection modes match Get-ClipHistory / gcbh:
      -Index N   absolute history slot (default positional, defaults to 0)
      -Image N   Nth image entry only (0 = newest image)
      -File  N   Nth Explorer file-handle entry only (0 = newest file)
#>

[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'ByIndex')]
param(
    [Parameter(Position = 0, ParameterSetName = 'ByIndex')]
    [int]$Index = 0,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByImage')]
    [int]$Image,

    [Parameter(Mandatory = $true, ParameterSetName = 'ByFile')]
    [int]$File,

    [Parameter()]
    [switch]$PassThru
)

. $PSScriptRoot\clipboard-runtime.ps1

[void](Invoke-WindowsPowerShellShim -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters)

try {
    # Avoid SetHistoryItemAsContent probing while enumerating; it mutates the
    # active clipboard and restore is unreliable.
    $entries = Get-ClipboardHistoryEntries
} catch {
    Write-Warning $_
    return
}

$indexValues = if ($PSCmdlet.ParameterSetName -eq 'ByIndex') { @($Index) } else { @() }
$imageValues = if ($PSCmdlet.ParameterSetName -eq 'ByImage') { @($Image) } else { @() }
$fileValues = if ($PSCmdlet.ParameterSetName -eq 'ByFile') { @($File) } else { @() }

$selected = @(Resolve-ClipboardSelection -Entries $entries -ParameterSetName $PSCmdlet.ParameterSetName -Index $indexValues -Image $imageValues -File $fileValues)
$entry = $selected | Select-Object -First 1
if ($null -eq $entry) {
    return
}

$label = switch ($PSCmdlet.ParameterSetName) {
    'ByImage' { "image ordinal $Image (absolute index $($entry.Index))" }
    'ByFile' { "file ordinal $File (absolute index $($entry.Index))" }
    default { "history item $($entry.Index)" }
}

if (-not $PSCmdlet.ShouldProcess('Clipboard', "Restore $label")) {
    return
}

try {
    $restored = Set-ClipboardFromHistoryEntry -Entry $entry
} catch {
    Write-Error $_
    return
}

if ($PassThru) {
    [pscustomobject]@{
        Index   = $restored.Index
        Type    = $restored.Type
        Preview = switch ($restored.Type) {
            'Text' { Get-ClipboardPreview -Text $restored.Content }
            'File' { Get-ClipboardPreview -Text $restored.Content }
            default { '[Image]' }
        }
    }
}
