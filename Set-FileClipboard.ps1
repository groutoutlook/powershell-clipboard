<#
.SYNOPSIS
    Places file system paths on the clipboard as Explorer copy or cut data.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Position = 0)]
    [Alias('LiteralPath', 'FullName', 'PSPath')]
    [string[]]$Path,

    [Parameter()]
    [switch]$Cut,

    [Parameter()]
    [switch]$PassThru
)

. $PSScriptRoot\clipboard-runtime.ps1

[void](Invoke-WindowsPowerShellShim -ScriptPath $PSCommandPath -BoundParameters $PSBoundParameters)

$operation = if ($Cut) { 'Cut' } else { 'Copy' }

try {
    if ($PSBoundParameters.ContainsKey('Path')) {
        $inputPaths = [string[]]@($Path)
        if ($inputPaths.Count -eq 0) {
            return
        }

        $resolvedPaths = @(Resolve-ClipboardFileSystemPaths -Path $inputPaths)
    } else {
        $resolvedPaths = @(Get-ClipboardCopyCandidatePaths)
    }
} catch {
    Write-Error $_
    return
}

$target = if ($resolvedPaths.Count -eq 1) {
    $resolvedPaths[0]
} else {
    "$($resolvedPaths.Count) items"
}

if (-not $PSCmdlet.ShouldProcess($target, "$operation to Explorer clipboard")) {
    return
}

try {
    $copiedPaths = @(Set-ClipboardFileDropList -Path $resolvedPaths -Operation $operation)
} catch {
    Write-Error $_
    return
}

if ($PassThru) {
    [pscustomobject]@{
        Operation = $operation
        Count     = $copiedPaths.Count
        Path      = $copiedPaths
    }
}