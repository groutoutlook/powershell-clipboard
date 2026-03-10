<#
.SYNOPSIS
    Restores a Windows clipboard history item into the active clipboard.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]
    [int]$Index = 0,

    [Parameter()]
    [switch]$PassThru
)

. $PSScriptRoot\clipboard-runtime.ps1

[void](Invoke-WindowsPowerShellShim -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters)

try {
    $entries = Get-ClipboardHistoryEntries
} catch {
    Write-Warning $_
    return
}

$entry = $entries | Where-Object { $_.Index -eq $Index } | Select-Object -First 1
if ($null -eq $entry) {
    Write-Error "Index $Index out of range. Use Get-ClipHistory.ps1 -All to list available items."
    return
}

if (-not $PSCmdlet.ShouldProcess('Clipboard', "Restore history item $Index")) {
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
        Preview = if ($restored.Type -eq 'Text') {
            Get-ClipboardPreview -Text $restored.Content
        } else {
            '[Image]'
        }
    }
}