# Sample with a hard-coded configuration

This project shows how to create a claims provider that inherits EntraCP. It uses a simple, hard-coded configuration to specify the tenant.

Do NOT deploy this solution in a SharePoint farm that already has EntraCP deployed, unless both use **exactly** the same versions of NuGet dependencies. If they use different versions, that may cause errors when loading DLLs, due to mismatches with the assembly bindings in the machine.config file.
