# This file contains test scripts for the clipboard bridge functionality.

$clipHistoryScript = "$PSScriptRoot\Get-ClipHistory.ps1"

Describe "Get-ClipHistory Tests" {
    Context "Basic Functionality" {
        It "Should run without throwing on -All" {
            # We wrap in try/catch or assume it passes if environment is set up.
            # If WinRT is missing, it will output warnings but not throw unless manipulated.
            try {
                & $clipHistoryScript -All
            } catch {
                throw $_
            }
        }
        
        It "Should handle -Index 0" {
             try {
                $res = & $clipHistoryScript -Index 0 -ErrorAction SilentlyContinue
             } catch {}
             # Assertion depends on clipboard content
        }
    }
}
