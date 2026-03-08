# PowerShell Clipboard Bridge

## Overview
The PowerShell Clipboard Bridge is a module designed to facilitate interaction with the Windows clipboard history. It provides functions to access, manipulate, and display clipboard items, making it easier for users to manage their clipboard data.

## Project Structure
The project consists of the following files:

- **Get-ClipHistory.ps1**: The main script that retrieves items from the Windows Clipboard history.
- **clipboard-bridge.Tests.ps1**: Test scripts.
- **justfile**: Command runner configuration.
- **README.md**: Documentation.

## Setup Instructions
1. Clone the repository.
2. Ensure you are running on Windows.
3. Run the script:
   ```powershell
   .\Get-ClipHistory.ps1 -All
   ```
   **Note**: This script works best in **Windows PowerShell 5.1**. In PowerShell 7, it attempts to access Windows Runtime APIs but may require environment configuration.

## Usage Examples
To retrieve clipboard history items, you can use the `Get-ClipboardHistory` function. For example:
```powershell
$historyItem = Get-ClipboardHistory -Index 0
Write-Output $historyItem
```

## Contributing
Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.

## License
This project is licensed under the MIT License. See the LICENSE file for more details.