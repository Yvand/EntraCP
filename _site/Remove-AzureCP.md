## How to remove AzureCP
For an unknown reason, randomly SharePoint 2013 doesn’t uninstall correctly the solution because it removes assembly from the GAC before calling the feature receiver... When this happens, the claims provider is not removed and that causes issues when you re-install it.  
To uninstall safely, deactivate the farm feature before retracting the solution:
```powershell
Disable-SPFeature -identity "AzureCP"
Uninstall-SPSolution -Identity "AzureCP.wsp"
Remove-SPSolution -Identity "AzureCP.wsp"
```
