<#
DISCLAIMER:

    By running this PowerShell script, you acknowledge and accept that:
    
    - You execute this script at your own risk and discretion.
    - The author of this script is not responsible for any damage, data loss, or unintended consequences 
      that may result from executing this script.
    - It is recommended to test this script in a controlled environment before use in a production scenario.

    If you do not agree with these terms, please refrain from executing this script.


    Name:               GenerateUserPasswords.PS1
    Version:            1.0
    Date:               24-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                This script performs the following tasks:

                        1. Imports user data from a CSV file named `Users.csv` located in the current directory.This file is expected to contain, 
                        at a minimum, columns for `DisplayName` and `Username`.  
                        2. Initializes an empty array, designed to store the generated passwords for each imported user.
                        3. For each user in the `Users.csv` file,  a random 8-character password is created. 
                        The character set for this password includes select special characters, numbers, uppercase, and 
                        lowercase letters, excluding potentially problematic characters like ", ', and \.
                        4. Each user's `DisplayName`, `Username`, and the newly generated `Password` are combined into a custom 
                        PowerShell object and added to the previously initialized password array.
                        5. Exports the updated user information, including the random passwords, to a new CSV file 
                        named `UserPasswords.csv` in the current directory. 

    Updates:        
#>

# Import user data from a CSV file named "Users.csv" from the current directory.
$users = Import-Csv -Path .\Users.csv

# Initialize an empty array to hold the generated passwords for each user.
$PasswordArray = @()

# Loop through each user from the imported CSV.
Foreach($User in $Users) {
    # Generate a random password consisting of 8 characters.
    # The character range includes certain special characters, numbers, uppercase, and lowercase letters.
    # The range skips characters like ", ', \ to avoid potential issues in CSV or passwords.
    $Password = -join ((33,35,36,38+(48..57)+(64..90)+(97..107)+(109..122)) | Get-Random -Count 8 | % {[char]$_})

    # Add the user's display name, username, and the generated password to the password array.
    $PasswordArray += [PSCustomObject]@{
        DisplayName = $User.DisplayName;
        UserName = $User.Username;
        Password = $Password
    } 
}

# Export the array containing display names, usernames, and passwords to a new CSV file named "UserPasswords.csv" in the current directory.
$PasswordArray | Export-Csv -Path .\UserPasswords.csv -NoTypeInformation
