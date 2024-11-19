# Input bindings are passed in via param block.
param($Timer)

# Get the current universal time in the default string format
$currentUTCtime = (Get-Date).ToUniversalTime()

# The 'IsPastDue' porperty is 'true' when the current function invocation is later than scheduled.
if ($Timer.IsPastDue) {
    Write-Host "PowerShell timer is running late!"
}

# Write an information log with the current time.
Write-Host "PowerShell timer trigger function ran! TIME: $currentUTCtime"


## Connect to the management site
Connect-PnPOnline -Url $env:ManagementSitetUrl -ClientId $env:ClientId -Thumbprint $env:Thumbprint -Tenant $env:TenantId
#Connect-PnPOnline -Url $env:ManagementSitetUrl -ManagedIdentity

## Get list of sites to check from the management site
$sites = Get-PnPListItem -List "SitesWithDenyAddAndCustomizePagesDisabled"

## Connect to the admin site
Connect-PnPOnline -Url $env:AdminUrl -ClientId $env:ClientId -Thumbprint $env:Thumbprint -Tenant $env:TenantId
#Connect-PnPOnline -Url $env:AdminUrl -ManagedIdentity

## Get each site using Get-PnPTenantSite and check if it has DenyAddAndCustomizePages enabled, if so reset
foreach ($site in $sites) {
    $siteDetails = Get-PnPTenantSite -Identity $site["Url"]

    "Checking site $($site["Url"])"

    if ($siteDetails.DenyAddAndCustomizePages -eq "Enabled") {
        Set-PnPTenantSite -DenyAddAndCustomizePages Disabled
        "Resetting site $($site["Url"])"
    }
}