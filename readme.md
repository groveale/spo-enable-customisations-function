# SPO Enable Customizations

This solution uses **PnP PowerShell** to manage the **Deny Add and Customize Pages** settings effectively. Below are the key components:

## Functions

### 1. GetDenyAddAndCustomizePagesSiteStatus
- **Purpose**: Retrieves the current status of the custom script setting for all sites and writes these totals to an SPO list. This script is used to understand when the SharePoint service resets the `DenyAddandCustomizePages` setting. There is also an option setting to pull all sites that have this disabled currently and write these sites to another list.
- **Schedule**: Configured to run hourly on the hour


### 2. SetDenyAddAndCustomizeStatus
- **Purpose**: Checks the `DenyAddandCustomizePages` setting for SharePoint sites that are details in a SharePoint list and confirms that the the `DenyAddandCustomizePages` setting is disabled, if enabled then disable.
- **Schedule**: Configured to daily at 3am


## Prerequisites

### SharePoint Site Setup

1. **Management Site**: Create a SharePoint site that will be used to manage the site status. This site will contain two lists:
    - **SitesWithDenyAddAndCustomizePagesDisabled**: This list will store the sites that have the `DenyAddAndCustomizePages` setting disabled.
    - **StatusTotals**: This list will store the totals of sites with and without customizations. 

### Create Self Signed Certificate

1. **Certificate creation**: Use the following PowerShell to create a self signed certificate that will be used to authenticate the app registration with SharePoint

```powershell
$certname = "spo-enable-customizations"    ## Replace {certificateName}
$cert = New-SelfSignedCertificate -Subject "CN=$certname" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256
Export-Certificate -Cert $cert -FilePath "$certname.cer"   ## Specify your preferred location

Write-Host "Enter password for private cert - This will be needed if you are planning to run the script from an Azure App service"
$pass = Read-Host -AsSecureString
# Export cert to PFX - uploaded to Azure App Service
Export-PfxCertificate -cert $cert -FilePath "$certname.pfx" -Password $pass

## Note the thumbprint
$cert.Thumbprint
```

### Azure Function App Setup

1. **App Registration**: Create an app registration in Azure AD with the necessary application permissions to access SharePoint Administrations (`Sites.FullControl.All`). Upload the `.cer` certificate to the app registration

2. **Function App**: Create a PowerShell function app, deploy the code and upload the `.pfx` certificate to the Azure Function.

3. **Function App Configuration**: Configure the Azure Function App with the following settings:
    - `ClientId`: The client ID of the app registration.
    - `TenantId`: The tenant ID of your Azure AD.
    - `Thumbprint`: The thumbprint of the certificate used for authentication.
    - `AdminUrl`: The URL of the SharePoint admin site.
    - `ManagementSitetUrl`: The URL of the management site created in step 1.
    - `GetSitesWithCustomizations`: Set to `true` to enable the option to pull all sites with customizations and write them to the list.
    - `WEBSITE_LOAD_CERTIFICATES`: The thumbprint of the certificate used for authentication. Or *. Let's the function know to load the certificate

## Outcome

