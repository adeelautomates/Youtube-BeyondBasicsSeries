# Install Module & authN to Exchange Online
Install-Module ExchangeOnlineManagement
Connect-ExchangeOnline

# SP that has the Mail.Send Permission. Use its Client ID
$appID = "0000000-0000-000-000-00000000000" 

# Mail-Enabled Security Group that has members you want to only send from. Use its email address
$groupUPN = "name@company.com"

# Add a description for the policy here
$description = ""

# run this cmdlet to apply policy to only let $appID send emails as the members of this group 
New-ApplicationAccessPolicy -AppId $appID -PolicyScopeGroupId $groupUPN -Description $description -AccessRight RestrictAccess 

# To see existing Policies
Get-ApplicationAccessPolicy

# To test and see if its working. Enter UPNS to see what it can send emails as and not. It only only work as members of the group
Test-ApplicationAccessPolicy -Identity "<upn>" -AppID $appId