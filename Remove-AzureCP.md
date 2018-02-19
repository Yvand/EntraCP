# How to remove AzureCP

## Step 1: Reset property ClaimProviderName in the SPTrustedIdentityTokenIssuer

Unfortunately, the only supported way to reset property ClaimProviderName is to remove and recreate the SPTrustedIdentityTokenIssuer object, which requires to remove the trust from all the zones where it is used first, which is time consuming.

Alternatively, it's possible to use reflection to reset this property, but it is not supported and you do this at your own risks. Here is the script:

```powershell
$trust = Get-SPTrustedIdentityTokenIssuer "SPTRUST NAME"
$trust.GetType().GetField("m_ClaimProviderName", "NonPublic, Instance").SetValue($trust, $null)
$trust.Update()
```

## Step 2: Uninstall AzureCP

Randomly, SharePoint doesn’t uninstall the solution correctly: it removes the assembly too early and fails to call the feature receiver... When this happens, the claims provider is not removed and that causes issues when you re-install it.

> **Important**: Always start a new PowerShell console to ensure it uses up to date persisted objects and avoid concurrency update errors.

```powershell
Disable-SPFeature -identity "AzureCP"
Uninstall-SPSolution -Identity "AzureCP.wsp"
# Wait for the timer job to complete
Remove-SPSolution -Identity "AzureCP.wsp"
```

Validate that claims provider was removed:

```powershell
Get-SPClaimProvider| ft DisplayName
# If AzureCP appears in cmdlet above, remove it:
Remove-SPClaimProvider AzureCP
```
