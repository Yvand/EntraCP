## Configure AzureCP

AzureCP needs to be registered as an application in Azure Active Directory to be able to run. (This page)[])Create-App-In-AAD.html) explains how to create the application.

### Configure with administration pages

2 administration are provisioned in central administration > Security:

- Global configuration: Register Azure AD tenants and configure general settings
- Claim types configuration: Define the claim types, and their mapping with users and groups

### Configure with PowerShell

Starting with v12, AzureCP can be configured with PowerShell:

```powershell
Add-Type -AssemblyName "AzureCP, Version=1.0.0.0, Culture=neutral, PublicKeyToken=65dc6b5903b51636"
$config = [ldapcp.LDAPCPConfig]::GetConfiguration("AzureCPConfig")

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

# Add a new entry to the claim types configuration list
$newCTConfig = New-Object ldapcp.ClaimTypeConfig
$newCTConfig.ClaimType = "ClaimTypeValue"
$newCTConfig.EntityType = [ldapcp.DirectoryObjectType]::User
$newCTConfig.LDAPClass = "LDAPClassVALUE"
$newCTConfig.LDAPAttribute = "LDAPAttributeVALUE"
$config.ClaimTypes.Add($newCTConfig)
$config.Update()

# Remove a claim type from the claim types configuration list
$claimTypes.Remove("ClaimTypeValue")
$config.Update()
```

AzureCP configuration is stored in SharePoint configuration database, in persisted object "AzureCPConfig", that can be displayed with this SQL command:

```sql
SELECT Id, Name, cast (properties as xml) AS XMLProps FROM Objects WHERE Name = 'AzureCPConfig'
```

## Configure the proxy

AzureCP makes HTTP requests to Azure under the process of the site (for people picker requests) and the SharePoint STS (for augmentation), but there are also requests made by lsass.exe to validate CRL and certificate chain of certificates returned by Azure.  
If SharePoint servers need to connect through a HTTP proxy, additional configuration is required to configure it:

### For AzureCP to be able to connect to Azure

Add the following [proxy configuration](https://msdn.microsoft.com/en-us/library/kd3cf2ex.aspx) in the web.config of:

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

### For certificate validation (CRL) to succeed

Certificate validation is performed by lsass.exe

- Display proxy configuration with this command:

```text
netsh winhttp show proxy
```

- Set proxy configuration with this command:  

```text
netsh winhttp set proxy proxy-server="http=myproxy;https=sproxy:88" bypass-list="*.foo.com"
```
