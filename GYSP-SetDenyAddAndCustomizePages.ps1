






Connect-SPOservice 
$URL = "https://MintecLtd.sharepoint.com/sites/urnerbarryteamsite"
Get-SPOSite -Identity $URL | select DenyAddAndCustomizePages
Set-SPOSite -Identity $URL -DenyAddAndCustomizePages 0
