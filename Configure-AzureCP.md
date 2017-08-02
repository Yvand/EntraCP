## Proxy configuration
AzureCP makes HTTP requests to Azure under the process of the site (for people picker requests) and the SharePoint STS (for augmentation), but there are also requests made by lsass.exe to validate CRL and certificate chain of certificates returned by Azure.<br>
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

### For certificate validation to succeed
- Display proxy configuration with this command:<br>
netsh winhttp show proxy
- Set proxy configuration with this command:<br>
netsh winhttp set proxy proxy-server="http=myproxy;https=sproxy:88" bypass-list="*.foo.com"


## Claims supported
Azure AD binds property UserPrincipalName (which identifies the user) to claim type "_http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name_". Unfortunately this claim type is reserved by SharePoint and cannot be used.

So the STS (whatever it is ACS, ADFS or another one) must transform this claim type into another one.  
AzureCP assumes it will be either "_http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress_" or "_http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn_" (you should define only 1 in the SPTrust), so both are binded to property **UserPrincipalName**.  
Properties **DisplayName, GivenName, Surname** are also used to query users, and **Mail, Mobile, JobTitle** are used as metadata for the permission created.  
For groups, property **DisplayName** is linked to _http://schemas.microsoft.com/ws/2008/06/identity/claims/role_  

AzureCP can augment Azure AD users with their group membership. This is configurable and enabled by default. Groups are stored in claim type "_http://schemas.microsoft.com/ws/2008/06/identity/claims/role_", but this can be changed. However there can be only 1 claim type used for the roles, otherwise augmentation will be disabled.

**All of this can be configured** in AzureCP pages added in central administration > security.
