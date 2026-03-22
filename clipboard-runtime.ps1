Set-StrictMode -Version Latest

$script:ClipboardTypeName = 'Windows.ApplicationModel.DataTransfer.Clipboard'
$script:AsTaskMethod = $null

function Convert-BoundParametersToArgumentList {
    param([hashtable]$BoundParameters)

    $arguments = @()
    foreach ($entry in $BoundParameters.GetEnumerator()) {
        if ($entry.Value -is [switch]) {
            if ($entry.Value.IsPresent) {
                $arguments += "-$($entry.Key)"
            }

            continue
        }

        $arguments += "-$($entry.Key)"

        if ($entry.Value -is [System.Array]) {
            $arguments += @($entry.Value | ForEach-Object { [string]$_ })
            continue
        }

        $arguments += [string]$entry.Value
    }

    return $arguments
}

function Invoke-WindowsPowerShellShim {
    param(
        [string]$ScriptPath,
        [hashtable]$BoundParameters
    )

    if ($PSVersionTable.PSVersion.Major -lt 6) {
        return $false
    }

    if (-not $IsWindows) {
        throw 'Windows Clipboard History requires Windows.'
    }

    $arguments = @(
        '-NoProfile'
        '-STA'
        '-ExecutionPolicy'
        'Bypass'
        '-File'
        $ScriptPath
    ) + (Convert-BoundParametersToArgumentList -BoundParameters $BoundParameters)

    & powershell.exe $arguments
    exit $LASTEXITCODE
}

function Get-WinRtType {
    param([string]$TypeName)

    $resolved = $TypeName -as [type]
    if ($null -ne $resolved) {
        return $resolved
    }

    Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue

    try {
        return Invoke-Expression "[$TypeName, Windows, ContentType=WindowsRuntime]"
    } catch {
        return $null
    }
}

function Get-ClipboardType {
    $clipboardType = Get-WinRtType -TypeName $script:ClipboardTypeName
    if ($null -eq $clipboardType) {
        throw 'Unable to load Windows clipboard WinRT types. Run this in Windows PowerShell 5.1.'
    }

    return $clipboardType
}

function Get-WindowsRuntimeAsTaskMethod {
    if ($script:AsTaskMethod) {
        return $script:AsTaskMethod
    }

    foreach ($method in [System.WindowsRuntimeSystemExtensions].GetMethods()) {
        if ($method.Name -ne 'AsTask') {
            continue
        }

        $parameters = $method.GetParameters()
        if ($parameters.Count -ne 1) {
            continue
        }

        if ($parameters[0].ParameterType.Name -ne 'IAsyncOperation`1') {
            continue
        }

        $script:AsTaskMethod = $method
        return $script:AsTaskMethod
    }

    return $null
}

function Await-WinRtOperation {
    param(
        $Operation,
        [type]$ResultType
    )

    if ($null -ne $ResultType) {
        $asTaskMethod = Get-WindowsRuntimeAsTaskMethod
        if ($null -ne $asTaskMethod) {
            $taskMethod = $asTaskMethod.MakeGenericMethod($ResultType)
            $task = $taskMethod.Invoke($null, @($Operation))
            $task.Wait()
            return $task.Result
        }
    }

    if (-not ($Operation.PSObject.Properties.Name -contains 'Status')) {
        throw 'Clipboard async operation does not expose Status and could not be converted to a Task.'
    }

    while ($Operation.Status.ToString() -eq 'Started') {
        Start-Sleep -Milliseconds 15
    }

    switch ($Operation.Status.ToString()) {
        'Completed' {
            if ($Operation.PSObject.Methods.Name -contains 'GetResults') {
                return $Operation.GetResults()
            }

            return $null
        }
        'Canceled' {
            throw 'Clipboard operation was canceled.'
        }
        'Error' {
            if ($Operation.PSObject.Properties.Name -contains 'ErrorCode') {
                throw $Operation.ErrorCode
            }

            throw 'Clipboard operation failed.'
        }
        default {
            throw "Clipboard operation ended in unexpected state '$($Operation.Status)'."
        }
    }
}

function Get-ClipboardHistoryItemsResult {
    $clipboardType = Get-ClipboardType
    $resultType = Get-WinRtType -TypeName 'Windows.ApplicationModel.DataTransfer.ClipboardHistoryItemsResult'

    if (-not $clipboardType::IsHistoryEnabled()) {
        throw 'Windows Clipboard History is disabled. Enable it in Settings > System > Clipboard.'
    }

    $result = Await-WinRtOperation -Operation ($clipboardType::GetHistoryItemsAsync()) -ResultType $resultType
    if ($result.Status.ToString() -ne 'Success') {
        if ($result.Status.ToString() -eq 'AccessDenied') {
            throw 'Access to clipboard history was denied.'
        }

        throw "Failed to read clipboard history. Status: $($result.Status)"
    }

    return $result
}

function Get-ClipboardHistoryEntries {
    $standardFormats = Get-WinRtType -TypeName 'Windows.ApplicationModel.DataTransfer.StandardDataFormats'
    if ($null -eq $standardFormats) {
        throw 'Unable to load standard clipboard formats.'
    }

    $entries = @()
    $result = Get-ClipboardHistoryItemsResult
    $index = 0

    foreach ($item in @($result.Items)) {
        if ($item.Content.Contains($standardFormats::Text)) {
            $entries += [pscustomobject]@{
                Index   = $index
                Type    = 'Text'
                Content = [string](Await-WinRtOperation -Operation ($item.Content.GetTextAsync()) -ResultType ([string]))
                Item    = $item
            }
        } elseif ($item.Content.Contains($standardFormats::Bitmap)) {
            $entries += [pscustomobject]@{
                Index   = $index
                Type    = 'Image'
                Content = $null
                Item    = $item
            }
        }

        $index++
    }

    return @($entries)
}

function Get-ClipboardPreview {
    param(
        [AllowEmptyString()]
        [string]$Text,
        [int]$MaxLines = 8
    )

    if ([string]::IsNullOrEmpty($Text)) {
        return ''
    }

    $lines = $Text -split '\r?\n'
    if ($lines.Count -le $MaxLines) {
        return $Text
    }

    $preview = $lines[0..($MaxLines - 1)] -join [Environment]::NewLine
    return "$preview$([Environment]::NewLine)... ($($lines.Count - $MaxLines) more lines)"
}

function Convert-ClipboardWinRtStreamToDotNetStream {
    param($Stream)

    $method = [System.IO.WindowsRuntimeStreamExtensions].GetMethods() |
        Where-Object { $_.Name -eq 'AsStreamForRead' -and $_.GetParameters().Count -eq 1 } |
        Select-Object -First 1

    if ($null -eq $method) {
        throw 'Unable to resolve WindowsRuntimeStreamExtensions.AsStreamForRead.'
    }

    return $method.Invoke($null, @($Stream))
}

function Get-ClipboardBitmapReference {
    param($DataPackageView)

    $streamReferenceType = Get-WinRtType -TypeName 'Windows.Storage.Streams.RandomAccessStreamReference'
    return Await-WinRtOperation -Operation ($DataPackageView.GetBitmapAsync()) -ResultType $streamReferenceType
}

function Get-ClipboardBitmapPngBytes {
    param($DataPackageView)

    $streamReference = Get-ClipboardBitmapReference -DataPackageView $DataPackageView
    if ($null -eq $streamReference) {
        return $null
    }

    $streamType = Get-WinRtType -TypeName 'Windows.Storage.Streams.IRandomAccessStreamWithContentType'
    $stream = Await-WinRtOperation -Operation ($streamReference.OpenReadAsync()) -ResultType $streamType
    $dotNetStream = $null
    $image = $null
    $memoryStream = $null

    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
        $dotNetStream = Convert-ClipboardWinRtStreamToDotNetStream -Stream $stream
        $image = [System.Drawing.Image]::FromStream($dotNetStream)
        $memoryStream = [System.IO.MemoryStream]::new()
        $image.Save($memoryStream, [System.Drawing.Imaging.ImageFormat]::Png)
        return $memoryStream.ToArray()
    } finally {
        if ($null -ne $image) {
            $image.Dispose()
        }

        if ($null -ne $memoryStream) {
            $memoryStream.Dispose()
        }

        if ($null -ne $dotNetStream) {
            $dotNetStream.Dispose()
        }
    }
}

function Show-ClipboardHistoryImage {
    param($DataPackageView)

    $catimg = Get-Command -Name 'catimg' -ErrorAction SilentlyContinue
    if ($null -eq $catimg) {
        throw 'catimg is not available in PATH.'
    }

    $bytes = Get-ClipboardBitmapPngBytes -DataPackageView $DataPackageView
    if ($null -eq $bytes -or $bytes.Length -eq 0) {
        return
    }

    $tempPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath (([IO.Path]::GetRandomFileName()) + '.png')

    try {
        [IO.File]::WriteAllBytes($tempPath, $bytes)
        & $catimg.Source $tempPath
    } finally {
        Remove-Item -Path $tempPath -Force -ErrorAction SilentlyContinue
    }
}

function Set-ClipboardFromHistoryEntry {
    param($Entry)

    $clipboardType = Get-ClipboardType
    $packageType = Get-WinRtType -TypeName 'Windows.ApplicationModel.DataTransfer.DataPackage'
    if ($null -eq $packageType) {
        throw 'Unable to load the DataPackage WinRT type.'
    }

    $package = $packageType::new()

    switch ($Entry.Type) {
        'Text' {
            $package.SetText($Entry.Content)
        }
        'Image' {
            $package.SetBitmap((Get-ClipboardBitmapReference -DataPackageView $Entry.Item.Content))
        }
        default {
            throw "Unsupported clipboard history item type '$($Entry.Type)'."
        }
    }

    $clipboardType::SetContent($package)
    $clipboardType::Flush()
    return $Entry
}

function Convert-ClipboardTextToPathCandidates {
    param([string]$Text)

    $candidates = @()

    foreach ($line in @($Text -split '\r?\n')) {
        $candidate = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        if ($candidate.Length -ge 2 -and $candidate.StartsWith('"') -and $candidate.EndsWith('"')) {
            $candidate = $candidate.Substring(1, $candidate.Length - 2)
        }

        $candidates += $candidate
    }

    return @($candidates)
}

function Get-ClipboardCopyCandidatePaths {
    Add-Type -AssemblyName System.Windows.Forms

    if ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
        return @([string[]][System.Windows.Forms.Clipboard]::GetFileDropList())
    }

    if (-not [System.Windows.Forms.Clipboard]::ContainsText()) {
        throw 'Clipboard does not contain copied files or folders. Copy items in Explorer, copy text paths, or pass -Path explicitly.'
    }

    $text = [System.Windows.Forms.Clipboard]::GetText()
    if ([string]::IsNullOrWhiteSpace($text)) {
        throw 'Clipboard text is empty and is not suitable for Explorer copy. Copy one or more files or folders, or pass -Path explicitly.'
    }

    $candidates = @(Convert-ClipboardTextToPathCandidates -Text $text)
    if ($candidates.Count -eq 0) {
        throw 'Clipboard text did not contain any usable file or folder paths. Copy items in Explorer, copy text paths, or pass -Path explicitly.'
    }

    try {
        return @(Resolve-ClipboardFileSystemPaths -Path $candidates)
    } catch {
        throw "Clipboard text is not suitable for Explorer copy. Copy existing files or folders, or pass -Path explicitly. $($_.Exception.Message)"
    }
}

function Resolve-ClipboardFileSystemPaths {
    param([string[]]$Path)

    $resolvedPaths = @()

    foreach ($item in $Path) {
        foreach ($resolved in @(Resolve-Path -LiteralPath $item -ErrorAction Stop)) {
            if ($resolved.Provider.Name -ne 'FileSystem') {
                throw "Path '$item' is not a file system path."
            }

            $resolvedPaths += [System.IO.Path]::GetFullPath($resolved.ProviderPath)
        }
    }

    return @($resolvedPaths | Select-Object -Unique)
}

function Set-ClipboardFileDropList {
    param(
        [string[]]$Path,
        [ValidateSet('Copy', 'Cut')]
        [string]$Operation = 'Copy'
    )

    Add-Type -AssemblyName System.Windows.Forms

    $resolvedPaths = @(Resolve-ClipboardFileSystemPaths -Path $Path)
    if ($resolvedPaths.Count -eq 0) {
        throw 'No file system paths were resolved.'
    }

    $dropList = [System.Collections.Specialized.StringCollection]::new()
    $dropList.AddRange([string[]]$resolvedPaths)

    $dataObject = [System.Windows.Forms.DataObject]::new()
    $dataObject.SetFileDropList($dropList)

    $effectBytes = if ($Operation -eq 'Cut') {
        [byte[]](2, 0, 0, 0)
    } else {
        [byte[]](5, 0, 0, 0)
    }

    $dataObject.SetData('Preferred DropEffect', [System.IO.MemoryStream]::new($effectBytes))
    [System.Windows.Forms.Clipboard]::SetDataObject($dataObject, $true)

    return @($resolvedPaths)
}