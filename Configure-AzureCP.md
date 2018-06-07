## Configure AzureCP

AzureCP needs to be registered as an application in Azure Active Directory, [this page](Create-App-In-AAD.html) explains how to do this.

### Configure with administration pages

AzureCP comes with 2 pages, provisioned in central administration > Security:

- Global configuration: Register Azure AD tenants and configure general settings
- Claim types configuration: Define the claim types, and their mapping with users and groups

### Configure with PowerShell

Starting with v12, AzureCP can be configured with PowerShell:

```powershell
Add-Type -AssemblyName "AzureCP, Version=1.0.0.0, Culture=neutral, PublicKeyToken=65dc6b5903b51636"
$config = [azurecp.AzureCPConfig]::GetConfiguration("AzureCPConfig")

# To view current configuration
$config
$config.ClaimTypes

# Update some settings, e.g. enable augmentation:
$config.EnableAugmentation = $true
$config.Update()

# Reset claim types configuration list to default
$config.ResetClaimTypesList()
$config.Update()

# Reset the whole configuration to default
$config.ResetCurrentConfiguration()
$config.Update()

# Add a new Azure AD tenant
$newAADTenant = New-Object azurecp.AzureTenant
$newAADTenant.TenantName = "xxx.onMicrosoft.com"
$newAADTenant.ClientId = "Application ID"
$newAADTenant.ClientSecret = "XXX"
$config.AzureTenants.Add($newAADTenant)
$config.Update()

# Add a new entry to the claim types configuration list
$newCTConfig = New-Object azurecp.ClaimTypeConfig
$newCTConfig.ClaimType = "ClaimTypeValue"
$newCTConfig.EntityType = [azurecp.DirectoryObjectType]::Group
$newCTConfig.DirectoryObjectProperty = [azurecp.AzureADObjectProperty]::Department
$claimTypes.Add($newCTConfig)
$config.Update()

# Remove a claim type from the claim types configuration list
$claimTypes.Remove("ClaimTypeValue")
$config.Update()
```

AzureCP configuration is stored in SharePoint configuration database, in persisted object "AzureCPConfig", that can be displayed with this SQL command:

```sql
SELECT Id, Name, cast (properties as xml) AS XMLProps FROM Objects WHERE Name = 'AzureCPConfig'
```

## Configure proxy for internet access

AzureCP makes HTTP requests to access Azure AD, and may run in all SharePoint processes (w3wp of the site, STS, central administration, and also in owstimer.exe).  
Besides this, connection is secured so Windows will validate the certificate chain returned by Azure.  
If SharePoint servers connect to internet through a proxy, additional configuration is required.

### To allow AzureCP to connect to Azure

If needed, add the [proxy configuration](https://msdn.microsoft.com/en-us/library/kd3cf2ex.aspx) in the web.config of:

- SharePoint sites that use AzureCP
- SharePoine central administration site
- SharePoint STS located in 15\WebServices\SecurityToken
- SharePoint Web Services root site
- Also create file owstimer.exe.config in 15\BIN of each SharePoint server to put proxy configuration

```xml
<system.net>
    <defaultProxy>
        <proxy proxyaddress="http://proxy.contoso.com:3128" bypassonlocal="true" />
    </defaultProxy>
</system.net>
```

### To allow certificate chain validation

If Windows cannot validate certificates, the usual symptom is a hang during 1 minute upon sign-in, and errors are recorded in CAPI2 event log.  
Certificate validation is performed by lsass.exe, which uses the proxy configured with netsh:

- Display proxy configuration:

```text
netsh winhttp show proxy
```

- Set proxy:

```text
netsh winhttp set proxy proxy-server="http=myproxy;https=sproxy:88" bypass-list="*.foo.com"
```
