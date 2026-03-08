<#
.SYNOPSIS
    Retrieves the clipboard history from Windows 10/11 Clipboard History feature.

.DESCRIPTION
    This script accesses the Windows Runtime API to retrieve clipboard history items.
    It requires Windows 10 (1809) or later with Clipboard History enabled.

.PARAMETER Index
    The index of the item to retrieve (0 is the most recent). Default is 0.

.PARAMETER All
    Returns all items in the history as an array of strings.

.EXAMPLE
    .\Get-ClipHistory.ps1 -All
    Retrieves all text items from clipboard history.
#>

[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [int[]]$Index = 0,
    
    [Parameter()]
    [switch]$All
)

# --- PowerShell Core Shim ---
# Windows Clipboard History API (WinRT) is natively accessible in Windows PowerShell 5.1.
# PowerShell 7+ (Core) requires complex interop setup to access these APIs directly.
# To keep this script simple and portable on Windows, we shim to powershell.exe if running in Core.
if ($PSVersionTable.PSVersion.Major -ge 6) {
    if ($IsWindows) {
        # Reconstruct arguments
        $psArgs = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", $PSCommandPath)
        
        # Add bound parameters
        $boundParams = $PSBoundParameters
        foreach ($key in $boundParams.Keys) {
            $val = $boundParams[$key]
            if ($val -is [switch]) {
                if ($val) { $psArgs += "-$key" }
            } elseif ($val -is [Array]) {
                $psArgs += "-$key"
                $psArgs += ($val -join ',')
            } else {
                $psArgs += "-$key"
                $psArgs += $val
            }
        }
        
        Write-Verbose "Shim: Redirecting to Windows PowerShell 5.1..."
        & powershell.exe $psArgs
        exit $LASTEXITCODE
    } else {
        Write-Error "This script relies on Windows Clipboard APIs and requires Windows."
        exit 1
    }
}

# Constants for Type Names
$ClipboardTypeName = "Windows.ApplicationModel.DataTransfer.Clipboard"
# This might be tricky because in PS 5.1 type names need specific assembly qualification sometimes
# but usually base names work if loaded.

try {
    # Try different loading strategies for different PS versions
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        # Core/PS7 requires specific interop to use WinRT.
        # Often it requires a shim/module.
        # However, we can try to load the assembly if present.
        if (-not ($ClipboardTypeName -as [Type])) {
             # No easy way to force load without exact path or module.
             # We rely on user environment.
        }
    } else {
        # Windows PowerShell 5.1
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue
        # Determine if type exists, if not, try the WinRT syntax
    }
} catch {}

# Function to Resolve Types safely across PS versions
function Get-WinRtType {
    param([string]$Name)
    
    $type = $Name -as [Type]
    if ($null -ne $type) { return $type }

    # Try PS 5.1 WinRT Loading Syntax if strict type not found
    if ($PSVersionTable.PSVersion.Major -le 5) {
        try {
            # This expression loads the assembly in PS 5.1
            # We use Invoke-Expression to avoid parse errors in PS 7 if syntax is invalid
            $expr = "[$Name, Windows.ApplicationModel.DataTransfer, ContentType=WindowsRuntime]"
            $loaded = Invoke-Expression $expr
            if ($loaded -is [Type]) { return $loaded }
        } catch {}
    }

    return $null
}

$ClipboardType = Get-WinRtType $ClipboardTypeName
if ($null -eq $ClipboardType) {
    # If we failed to get the type, we can't proceed with this script logic in pure PS7 without a module.
    # But checking if we are on Windows...
    if ($IsWindows -or $PSVersionTable.PSVersion.Major -le 5) {
         # Last attempt: Load assembly by partial name (deprecated but sometimes works)
         try {
            [void][System.Reflection.Assembly]::LoadWithPartialName("Windows.ApplicationModel.DataTransfer")
            $ClipboardType = $ClipboardTypeName -as [Type]
         } catch {}
    }
}

if ($null -eq $ClipboardType) {
    Write-Warning "Could not resolve WinRT Clipboard types."
    Write-Warning "This script is optimized for Windows PowerShell 5.1. On PowerShell 7, ensure WinRT interop is valid."
    return
}

# Check History Enabled
if (-not $ClipboardType::IsHistoryEnabled()) {
    Write-Warning "Windows Clipboard History is disabled. Enable it in Windows Settings > System > Clipboard."
    return
}

# --- Helper to wait for Async Operations ---
# We try to use AsTask if available (PS 5.1 standard), or fallback to simple loop (PS 7 compat)

$asTaskMethod = $null
try {
    $extensions = [System.WindowsRuntimeSystemExtensions]
    # Finding the specific AsTask<TResult>(IAsyncOperation<TResult>)
    $methods = $extensions.GetMethods() | Where-Object { $_.Name -eq 'AsTask' }
    foreach ($m in $methods) {
        try {
            $params = $m.GetParameters()
            if ($params.Count -eq 1 -and $params[0].ParameterType.Name -eq 'IAsyncOperation`1') {
                $asTaskMethod = $m
                break
            }
        } catch { 
            # Ignore reflection errors on incompatible members
        }
    }
} catch { }

function Await-WinRt {
    param($Task, $ResultType)
    
    # Strategy 1: Use AsTask if available
    if ($null -ne $asTaskMethod -and $null -ne $ResultType) {
        try {
            $generic = $asTaskMethod.MakeGenericMethod($ResultType)
            $netTask = $generic.Invoke($null, @($Task))
            $netTask.Wait()
            return $netTask.Result # Wait returns void, Result allows access
        } catch {
            # Fallback if invocation fails
        }
    }

    # Strategy 2: Simple polling (Works if object has Status/GetResults)
    # The Task logic in PS 7 often wraps the WinRT object.
    $op = $Task
    # Most likely it's an IAsyncOperation
    while ($op.Status.ToString() -eq 'Started') {
        Start-Sleep -Milliseconds 10
    }
    
    if ($op.Status.ToString() -eq 'Completed') {
        return $op.GetResults()
    }
    
    throw "Async operation failed with status: $($op.Status)"
}

# --- Main Logic ---

# 1. Get History Items
$op = $ClipboardType::GetHistoryItemsAsync()
# Need result type for AsTask: [Windows.ApplicationModel.DataTransfer.ClipboardHistoryItemsResult]
$resType = Get-WinRtType "Windows.ApplicationModel.DataTransfer.ClipboardHistoryItemsResult"

# If result type not found, we can't use AsTask, but Polling might work without it?
# Yes, simple polling doesn't need the Type object, just property access.
$result = Await-WinRt $op $resType

if ($result.Status.ToString() -ne 'Success') {
    if ($result.Status.ToString() -eq 'AccessDenied') {
        Write-Error "Access to Clipboard History was denied. Focus the application or check permissions."
        return
    }
    Write-Error "Failed to retrieve Clipboard History. Status: $($result.Status)"
    return
}

$items = $result.Items
if ($null -eq $items -or $items.Count -eq 0) {
    return
}


# 2. Iterate and Extract Text
$collection = @()

# We need the StandardDataFormats.Text static property
# Resolve Type 
$fmtType = Get-WinRtType "Windows.ApplicationModel.DataTransfer.StandardDataFormats"
$textFmt = $fmtType::Text 

foreach ($item in $items) {
    if ($item.Content.Contains($textFmt)) {
        try {
            $textOp = $item.Content.GetTextAsync()
            $text = Await-WinRt $textOp ([string])
            $collection += $text
        } catch {
            Write-Verbose "Failed item text extraction: $_"
        }
    }
}

if ($All) {
    for ($i = 0; $i -lt $collection.Count; $i++) {
        Write-Host "[$i]" -ForegroundColor Cyan -NoNewline
        Write-Host " $($collection[$i])"
    }
    return
} else {
    foreach ($idx in $Index) {
        if ($idx -lt $collection.Count) {
            # If multiple indices or formatting requested implicitly by user feedback, show index
            if ($Index.Count -gt 1) {
                Write-Host "[$idx]" -ForegroundColor Cyan -NoNewline
                Write-Host " $($collection[$idx])"
            } else {
                # Single item request -> return raw value for pipeline/clipboard use
                return $collection[$idx]
            }
        } else {
            Write-Error "Index $idx out of range. Use -All to see all $($collection.Count) items."
        }
    }
}
