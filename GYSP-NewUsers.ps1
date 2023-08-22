<#
    Name:               GYSP-NewUsers.PS1
    Version:            1.0
    Date:               22-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

    Use:                Creates bulk AzureAD users from CSV
                        CSV should have the following fields:
                            -DisplayName
                            -MailNickname
                            -UserPrincipalName
                            -Password

                            cd "C:\Users\AnsellS\OneDrive - Softcat Plc\Documents\Scratch"


    Updates:        
#>

# Connect to Microsoft Graph API with required scopes
Connect-MgGraph -Scopes 'User.ReadWrite.All'

# Specify the path of the CSV file
$CSVFilePath = ".\Users.csv"



# Import data from CSV file
$NewMGUsers = Import-Csv -Path $CSVFilePath

# Loop through each row containing user details in the CSV file
foreach ($NewMGUser in $NewMGUsers) {
$Password = 

    # Create password profile
    $PasswordProfile = @{
    Password                             = $NewMGUser.Password
    ForceChangePasswordNextSignIn        = $true
    #ForceChangePasswordNextSignInWithMfa = $true
}
    $UserParams = @{
        DisplayName       = $NewMGUser.DisplayName
        MailNickName      = $NewMGUser.MailNickName
        UserPrincipalName = $NewMGUser.UserPrincipalName
        PasswordProfile   = $PasswordProfile
        AccountEnabled    = $true
    }

    try {
        $null = New-MgUser @UserParams -ErrorAction Stop
        Write-Host ("Successfully created the account for {0}" -f $NewMGUser.DisplayName) -ForegroundColor Green
    }
    catch {
        Write-Host ("Failed to create the account for {0}. Error: {1}" -f $NewMGUser.DisplayName, $_.Exception.Message) -ForegroundColor Red
    }
}