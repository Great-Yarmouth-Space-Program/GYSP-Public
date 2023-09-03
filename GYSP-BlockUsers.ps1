
Connect-AzureAD


$BlockUsers = Import-Csv .\BlockUsers.csv

Foreach($BlockUser in $BlockUsers) {

Set-AzureADUser -ObjectID $BlockUser.UserPrincipalName -AccountEnabled $false
Revoke-AzureADUserAllRefreshToken -ObjectId $BlockUser.UserPrincipalName
Write-host $BlockUser.UserPrincipalName

}