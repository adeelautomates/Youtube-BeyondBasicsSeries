# Azure Functions profile.ps1
Import-Module Az.Accounts
Import-Module Az.Resources

if ($env:MSI_SECRET) {
    Disable-AzContextAutosave -Scope Process | Out-Null
    Connect-AzAccount -Identity
}
