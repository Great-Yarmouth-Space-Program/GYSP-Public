<#
DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.


    Name:               Disconnect-MGSession.PS1
    Version:            1.0
    Date:               23-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Disconnects the Graph session and removes the cached API Token

    Updates:        
#>
# Define a custom function named "Disconnect-MGSession"
Function Disconnect-MGSession {
    try {
        # Disconnect from the Microsoft Graph API 
        Disconnect-MgGraph
        
        # Remove the cached Graph API token
        Remove-Item "$env:USERPROFILE\.graph" -Recurse -Force
        
        Write-Host "Successfully disconnected from MGSession."
    } catch {
        # If an error occurs, catch the exception and display an error message
        Write-Host "Error during MGSession disconnect:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
    }
}