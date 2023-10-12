Get-MgDomainServiceConfigurationRecord
Get-MgDomainRefDomainNameerenceByRef



v=spf1 include:spf.protection.outlook.com -all


    $Dmarc = Resolve-DnsName -name _dmarc.softcat.com -Type TXT -ErrorAction SilentlyContinue | Where-Object { $_.strings -like 'v=DMARC1*' }    


$DmarcDomain = "_dmarc.softcat.com"

Get-MalwareFilterPolicy | fl Identity, EnableInternalSenderAdminNotifications, InternalSenderAdminAddress

Get-AdminAuditLogConfig | Select-Object AdminAuditLogEnabled, UnifiedAuditLogIngestionEnabled


Get-Mailbox -ResultSize Unlimited | Where-Object {$_.AuditEnabled -ne $true -and ($_.RecipientTypeDetails -ne "UserMailbox" -or $_.RecipientTypeDetails -ne "SharedMailbox")} 

$TeamsGlobalMeetingPolicy | Select identity,RecordingStorageMode



Get-SPOTenant | fl RequireAnonymousLinksExpireInDays



Get-mgpasswordAuthenticationMethod

Get-MgUser -UserId 85606bd7-00b5-4411-839d-f6262e0eced2 | select -ExpandProperty PasswordPolicies

get-mgdomain | FL

get-mgdirectory
