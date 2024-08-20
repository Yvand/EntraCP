# Sample with a hard-coded configuration and a manual reference to EntraCP

This project shows how to create a claims provider that inherits EntraCP. It uses a simple, hard-coded configuration to specify the tenant.

> [!WARNING]
> Do NOT deploy this solution in a SharePoint farm that already has EntraCP deployed, unless both use **exactly** the same versions of NuGet dependencies. If they use different versions, that may cause errors when loading DLLs, due to mismatches with the assembly bindings in the machine.config file.

> [!IMPORTANT]  
> You need to manually add a reference to `Yvand.EntraCP.dll`, and its version will determine the version of the Nuget packages `Azure.Identity` and `Microsoft.Graph` you should use in your project, because that will allow you to reuse the same assembly bindings provided in the file `assembly-bindings.config` and avoid the tedious task of figuring them all out by yourself.
