# PowerShell Clipboard Utils

## Setup 
- Clone the repository. Run the script:
```powershell
.\Get-ClipHistory.ps1 -All
```
**Note**: This script works best in **Windows PowerShell 5.1**.
In PowerShell 7, it attempts to access Windows Runtime APIs but may require environment configuration.

## Usage 
To retrieve clipboard history items, you can use the `Get-ClipboardHistory` function.
```powershell
$historyItem = Get-ClipboardHistory -Index 0
Write-Output $historyItem
```
