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

## connect to admin site
Connect-PnPOnline -Url $env:AdminUrl -ClientId $env:ClientId -Thumbprint $env:Thumbprint -Tenant $env:TenantId
#Connect-PnPOnline -Url $env:AdminUrl -ManagedIdentity

## Sites with DenyAddAndCustomizePages Disabled i.e. can run scripts
$sites = Get-PnPTenantSite
$sitesWithDenyAddAndCustomizePagesDisabled = $sites | where { $_.DenyAddAndCustomizePages -eq "Disabled" }
$sitesWithDenyAddAndCustomizePagesEnabled = $sites | where { $_.DenyAddAndCustomizePages -eq "Enabled" } 

## Create an object to hold the totlas
$totals = @{
    "TotalSites" = $sites.Count
    "SitesWithCustomisations" = $sitesWithDenyAddAndCustomizePagesDisabled.Count
    "SitesWithout" = $sitesWithDenyAddAndCustomizePagesEnabled.Count
    "LastUpdated" = $currentUTCtime
}

## Connect to the management site
Connect-PnPOnline -Url $env:ManagementSitetUrl -ClientId $env:ClientId -Thumbprint $env:Thumbprint -Tenant $env:TenantId
#Connect-PnPOnline -Url $env:ManagementSitetUrl -ManagedIdentity

## Add the totals to the list
Add-PnPListItem -List "StatusTotals" -Values $totals

if ($env:GetSitesWithCustomizations)
{
    # Create an array to hold list items
    $listItems = $sitesWithDenyAddAndCustomizePagesDisabled | ForEach-Object {
        @{
            "Url" = $_.Url
            "OwnerEmail" = $_.Owner
        }
    }

    # Go through each list item, get if item already exists, if not, add it
    foreach ($listItem in $listItems) {
        $itemExists = Get-PnPListItem -List "SitesWithDenyAddAndCustomizePagesDisabled" -Query "<View><Query><Where><Eq><FieldRef Name='Url'/><Value Type='Text'>$($listItem.Url)</Value></Eq></Where></Query></View>"
        if ($itemExists -eq $null) {
            Add-PnPListItem -List "SitesWithDenyAddAndCustomizePagesDisabled" -Values $listItem
        }
        else {
            Set-PnPListItem -List "SitesWithDenyAddAndCustomizePagesDisabled" -Identity $itemExists.Id -Values $listItem
        }
    }
}




