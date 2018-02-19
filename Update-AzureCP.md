## How to update AzureCP

> **Important:** Always start a new PowerShell console to ensure it uses up to date persisted objects. This avoids painful update errors.

Run Update-SPSolution cmdlet to start a timer job that that will deploy the update on all servers. You can monitor the progression in farm solutions page in central administration.

```powershell
Update-SPSolution -GACDeployment -Identity "AzureCP.wsp" -LiteralPath "PATH TO WSP FILE"
```

Then run iisreset on each server. This is mandatory to ensure that all w3wp processes, including SecurityTokenServiceApplicationPool, get restarted (this one is not restarted by solution update job).

> **Important:** If some SharePoint servers do not have SharePoint service "Microsoft SharePoint Foundation Web Application" started, you need to deploy AzureCP.dll on their GAC manually as shown in [this page](Install-AzureCP.html).