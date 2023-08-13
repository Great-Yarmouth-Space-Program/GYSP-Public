<#
    Name:               GYSP-SignOutUsers.PS1
    Version:            1.0
    Date:               13-08-2023
    Original Author:    Si Ansell
    Email:              graph@greatyarmouthspaceprogram.space

   Modules:             Microsoft Graph               Install-Module Microsoft.Graph -Scope AllUsers
   
    Use:                Stops sign-in sessions by revoking tokens for users whose UserPrincipalName does not contain "*onmicrosoft.com"

    Updates:        
#>

# Connect to Microsoft Graph with specified scopes
Connect-MgGraph -Scopes User.ReadWrite.All

# Retrieve users to revoke sessions for
$UserstoRevoke = Get-MGUser | Where-Object { $_.UserPrincipalName -notlike "*onmicrosoft.com" }

# Loop through each user and revoke their sign-in sessions
Foreach ($UsertoRevoke in $UserstoRevoke) {
    Revoke-MgUserSignInSession -UserId $UsertoRevoke.Id 
    Write-Host $UsertoRevoke.UserPrincipalName "token revoked" -ForegroundColor Green
}

# Disconnect from Microsoft Graph
# Disconnect-MGGraph