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

# Configuration
$MaxDisplayLines = 8

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

function Render-Image {
    param($DataPackageView)
    
    # Check for renderer
    $renderer = Get-Command "catimg", "viu", "chafa", "lsc" -ErrorAction SilentlyContinue | Select-Object -First 1
    
    if (-not $renderer) {
        Write-Warning "Image content found, but no image renderer (catimg, viu, chafa, lsc) is available in PATH."
        Write-Warning "Suggestion: Install 'catimg' or 'viu' (cargo install viu)"
        return
    }

    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue

        $bmpOp = $DataPackageView.GetBitmapAsync()
        $ref = Await-WinRt $bmpOp ([Windows.Storage.Streams.RandomAccessStreamReference])
        if ($null -eq $ref) { return }

        $readOp = $ref.OpenReadAsync()
        $rtStream = Await-WinRt $readOp ([Windows.Storage.Streams.IRandomAccessStreamWithContentType])
        
        $dotnetStream = $null
        
        # --- Stream Conversion Strategy ---
        # Strategy 1: Attempt to compile a C# helper to handle strict interface casting avoiding PS binder issues.
        # This is often necessary when WinRT objects are wrapped in System.__ComObject and PS loses type info.
        if (-not ("WinRtAdapter" -as [Type])) {
            $refs = @("System.Runtime.WindowsRuntime")
            # Locate Windows.Foundation.winmd or Windows.WinMD
            $winMdPaths = @(
                "$env:Windir\System32\WinMetadata\Windows.Foundation.winmd",
                "$env:Windir\System32\WinMetadata\Windows.Storage.winmd"
            )
            foreach ($p in $winMdPaths) { if (Test-Path $p) { $refs += $p } }
            
            $csharp = @"
            using System;
            using System.IO;
            using Windows.Storage.Streams;
            using System.Runtime.InteropServices.WindowsRuntime;
            
            public static class WinRtAdapter {
                public static Stream ToStream(object obj) {
                    if (obj is IInputStream) {
                        return WindowsRuntimeStreamExtensions.AsStreamForRead((IInputStream)obj);
                    }
                    return null;
                }
            }
"@
            try {
                Add-Type -TypeDefinition $csharp -ReferencedAssemblies $refs -ErrorAction SilentlyContinue
            } catch {
                Write-Verbose "WinRtAdapter compilation failed (harmless if fallback works): $_"
            }
        }
        
        if ("WinRtAdapter" -as [Type]) {
            try {
                $dotnetStream = [WinRtAdapter]::ToStream($rtStream)
            } catch {
                Write-Verbose "Adapter conversion failed."
            }
        }

        # Strategy 2: Reflection on Extension Method (Bypassing strong typing of PS binder)
        if ($null -eq $dotnetStream) {
            try {
                $extType = [System.IO.WindowsRuntimeStreamExtensions]
                $methods = $extType.GetMethods() | Where-Object { $_.Name -eq 'AsStreamForRead' -and $_.GetParameters().Count -eq 1 }
                foreach ($m in $methods) {
                    try {
                        $dotnetStream = $m.Invoke($null, @($rtStream))
                        if ($dotnetStream) { break }
                    } catch {}
                }
            } catch {}
        }
        
        # Strategy 3: Manual DataReader (Last Resort)
        if ($null -eq $dotnetStream) {
            try {
                $reader = [Windows.Storage.Streams.DataReader]::new($rtStream) # PS Constructor binding magic
                $loadOp = $reader.LoadAsync([uint32]$rtStream.Size)
                [void](Await-WinRt $loadOp ([uint32]))
                $bytes = New-Object byte[] $rtStream.Size
                $reader.ReadBytes($bytes)
                $dotnetStream = [System.IO.MemoryStream]::new($bytes)
            } catch {
                # If this fails, we really can't read it.
                Write-Verbose "Manual DataReader failed: $_"
            }
        }
        
        if ($null -eq $dotnetStream) {
             Write-Error "Could not convert internal WinRT stream to .NET stream. Your PowerShell version might not support this interop."
             return
        }
        
        $tempFile = [System.IO.Path]::GetTempFileName()
        $pngFile = [System.IO.Path]::ChangeExtension($tempFile, ".png")
        
        # Load image from stream and save as PNG
        $img = [System.Drawing.Image]::FromStream($dotnetStream)
        $img.Save($pngFile, [System.Drawing.Imaging.ImageFormat]::Png)
        $img.Dispose()
        $dotnetStream.Dispose()
        try { $rtStream.Dispose() } catch {}
        
        # Run the renderer
        if ($renderer.Name -eq 'chafa') {
            & $renderer.Name -f sixel $pngFile
        } else {
            & $renderer.Name $pngFile
        }
        
        Remove-Item $pngFile -Force -ErrorAction SilentlyContinue
        if (Test-Path $tempFile) { Remove-Item $tempFile -Force -ErrorAction SilentlyContinue }
        
    } catch {
        Write-Error "Failed to render image: $_"
    }
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


# 2. Iterate and Extract Text & Images
$collection = @()

# We need the StandardDataFormats static properties
$fmtType = Get-WinRtType "Windows.ApplicationModel.DataTransfer.StandardDataFormats"
$textFmt = $fmtType::Text 
$bitmapFmt = $fmtType::Bitmap

foreach ($item in $items) {
    if ($item.Content.Contains($textFmt)) {
        try {
            $textOp = $item.Content.GetTextAsync()
            $text = Await-WinRt $textOp ([string])
            $collection += [PSCustomObject]@{
                Type = 'Text'
                Content = $text
                Item = $item
            }
        } catch {
            Write-Verbose "Failed item text extraction: $_"
        }
    } elseif ($item.Content.Contains($bitmapFmt)) {
        $collection += [PSCustomObject]@{
            Type = 'Image'
            Content = $null
            Item = $item
        }
    }
}

if ($All) {
    $ESC = [char]27
    $Cyan = "$ESC[36m"
    $DarkGray = "$ESC[90m"
    $Reset = "$ESC[0m"

    for ($i = 0; $i -lt $collection.Count; $i++) {
        $outputString = $Cyan + "[$i]" + $Reset
        $entry = $collection[$i]
        
        if ($entry.Type -eq 'Text') {
            $lines = $entry.Content -split '\r?\n'
            if ($lines.Count -gt $MaxDisplayLines) {
                $preview = $lines[0..($MaxDisplayLines - 1)] -join "`n"
                $outputString += " $preview`n"
                $outputString += "$DarkGray... ($($lines.Count - $MaxDisplayLines) more lines truncated)$Reset"
            } else {
                $outputString += " $($entry.Content)"
            }
        } elseif ($entry.Type -eq 'Image') {
            $outputString += " [Image] (Bitmap)"
        }

        Write-Output $outputString
    }
    return
} else {
    foreach ($idx in $Index) {
        if ($idx -ge 0 -and $idx -lt $collection.Count) {
            $entry = $collection[$idx]
            
            if ($entry.Type -eq 'Text') {
                if ($Index.Count -gt 1) {
                    Write-Host "[$idx]" -ForegroundColor Cyan -NoNewline
                    Write-Host " $($entry.Content)"
                } else {
                    return $entry.Content
                }
            } elseif ($entry.Type -eq 'Image') {
                if ($Index.Count -gt 1) {
                    Write-Host "[$idx]" -ForegroundColor Cyan -NoNewline
                    Write-Host " [Image]"
                }
                Render-Image $entry.Item.Content
            }
        } else {
            Write-Error "Index $idx out of range. Use -All to see all $($collection.Count) items."
        }
    }
}
