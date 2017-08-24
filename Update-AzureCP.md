## How to update AzureCP

Run Update-SPSolution cmdlet to start a timer job that that will deploy the update. You can monitor the progression in farm solutions page in central administration.

```powershell
# Run this on a new PowerShell console (it tends to avoid issues with local cache of persisted objects, that could cause errors on such operations)
Update-SPSolution -GACDeployment -Identity "AzureCP.wsp" -LiteralPath "PATH TO WSP FILE"
```

Then run iisreset on each server. This is mandatory to ensure that all w3wp processes, including SecurityTokenServiceApplicationPool, get restarted (this one is not restarted by solution update job).
