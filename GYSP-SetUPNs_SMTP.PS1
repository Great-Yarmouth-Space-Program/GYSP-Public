<#

DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.


    Name:               GYSP-SetUPNs_SMTP.PS1
    Version:            1.0
    Date:               11-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Sets the UPN and SMTP of users in a CSV file
                        CSV should be in the format: OLDUPN,NEWUPN

    Modules:            ExchangeOnlineManagement
    
    Connections:        Connect-ExchangeOnline
    
    Updates:
#>

# Import user data from a CSV file named 'NewUPNs.csv' located in the current directory
$Users = Import-Csv .\NewUPNs.csv

# Count the total number of users for progress tracking
$totalUsers = $Users.Count
$completedUsers = 0

# Loop through each user entry in the imported CSV
Foreach ($User in $Users) {

    # Calculate the progress percentage
    $progressPercentage = ($completedUsers / $totalUsers) * 100

    # Display a progress bar
    Write-Progress -Activity "Updating Mailboxes" -Status "Processing user $completedUsers of $totalUsers" -PercentComplete $progressPercentage

    # Use a try-catch block to handle potential errors during mailbox updates
    try {
        # Update the mailbox details using the old UPN as identity 
        # Set the Windows Email Address and Microsoft Online Services ID to the new UPN from the CSV
        Set-Mailbox -Identity $User.OldUPN -WindowsEmailAddress $User.NEWUPN -MicrosoftOnlineServicesID $User.NEWUPN 
        
        # Increment the completed users count
        $completedUsers++
    } 
    catch {
        # If an error occurs during the Set-Mailbox operation, write the error message with the relevant user's OldUPN
        Write-Error "Failed to set details for user: $($User.OldUPN). Error: $_"
    }
}

# Clear the progress bar after completion
Write-Progress -Activity "Updating Mailboxes" -Status "Completed" -Completed