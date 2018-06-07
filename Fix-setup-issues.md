## Fix setup issues

Sometimes, install/uninstall/update of AzureCP solution fails. Most of the time, it occurs when cmdlets were executed in an old PowerShell console that had stale persisted objects. This caused concurrency update errors and SharePoint cancelled operation in the middle of the process.  
When this happens, some AzureCP features are in an inconsistent state that must be fixed, this page will walk you through the steps to clean this.

> **Important:**  
> Start a **new PowerShell console** to ensure the use of up to date persisted objects, which avoids concurrency update errors.  
> Make all operations in the server **running central administration**, in this order.

### Remove AzureCP claims provider

```powershell
Get-SPClaimProvider| ?{$_.DisplayName -like "AzureCP"}| Remove-SPClaimProvider
```

### Identify AzureCP features still installed

```powershell
# Identify all AzureCP features still installed on the farm, and that need to be manually uninstalled
Get-SPFeature| ?{$_.DisplayName -like 'AzureCP*'}| fl DisplayName, Scope, Id, RootDirectory
```

Usually, only AzureCP farm feature is listed:

```text
DisplayName   : AzureCP
Scope         :
Id            : d1817470-ca9f-4b0c-83c5-ea61f9b0660d
RootDirectory : C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\Template\Features\AzureCP
```

### Recreate missing feature folders and add feature.xml

For each feature listed, check if its "RootDirectory" actually exists in the file system of the current server. If it does not exist:

* Create the "RootDirectory" (e.g. "AzureCP" in "Features" folder)
* Use [7-zip](http://www.7-zip.org/) to open AzureCP.wsp and extract the feature.xml of the corresponding feature
* Copy the feature.xml into the "RootDirectory"

### Deactivate and remove the features

```powershell
# Deactivate AzureCP features (it may thrown an error if feature is already deactivated)
Get-SPFeature| ?{$_.DisplayName -like 'AzureCP*'}| Disable-SPFeature -Confirm:$false
# Uninstall AzureCP features
Get-SPFeature| ?{$_.DisplayName -like 'AzureCP*'}| Uninstall-SPFeature -Confirm:$false
```

### Delete the AzureCP persisted object

AzureCP stores its configuration is its own persisted object, and sometimes this object may not be deleted. In such scenario, this stsadm command can delete it:

```
stsadm -o deleteconfigurationobject -id 0E9F8FB6-B314-4CCC-866D-DEC0BE76C237
```

### If desired, AzureCP solution can now be safely removed

```powershell
Remove-SPSolution "AzureCP.wsp" -Confirm:$false
```
