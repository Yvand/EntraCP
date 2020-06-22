# Change log for AzureCP

## Unreleased

* Fix error when earching users with a single quote in the people picker - https://github.com/Yvand/AzureCP/issues/88
* Use batch requests when relevant to optimize network performance
* Avoid potential null reference exception in method AzureCP.ProcessAzureADResults() and ensure that method AzureCP.QueryAzureADTenantAsync() never returns null
* Update target framework to v4.6.1
* Update NuGet package Microsoft.Graph 1.21.0 -> Microsoft.Graph 3.8.0
* Update NuGet package Microsoft.IdentityModel.Clients.ActiveDirectory 5.2.6 -> Microsoft.IdentityModel.Clients.ActiveDirectory 5.2.7

## AzureCP 16.0.20200303.1092 enhancements & bug-fixes - Published in March 17, 2020

* Fix hang while using "check permissions" feature - https://github.com/Yvand/AzureCP/issues/78
* Update NuGet package Microsoft.Graph.1.18.0 -> Microsoft.Graph.1.21.0
* Update NuGet package Microsoft.Graph.Core.1.18.0 -> Microsoft.Graph.Core.1.19.0
* Update NuGet package Microsoft.IdentityModel.Clients.ActiveDirectory.5.2.3 -> Microsoft.IdentityModel.Clients.ActiveDirectory.5.2.6
* Update NuGet package NUnit3TestAdapter.3.15.1 -> NUnit3TestAdapter.3.16.1
* Update NuGet package Newtonsoft.Json.12.0.2 -> Newtonsoft.Json.12.0.3 in AzureCP.Tests project only
* Add assembly dependency introduced by Microsoft.Graph.Core.1.19.0 -> System.Buffers.4.4.0
* Add assembly dependency introduced by Microsoft.Graph.Core.1.19.0 -> System.Diagnostics.DiagnosticSource.4.6.0
* Add assembly dependency introduced by Microsoft.Graph.Core.1.19.0 -> System.Memory.4.5.3
* Add assembly dependency introduced by Microsoft.Graph.Core.1.19.0 -> System.Runtime.CompilerServices.Unsafe.4.5.2

## AzureCP 15.3.20191120.1028 enhancements & bug-fixes - Published in November 20, 2019

* Fix a bug where search fails if input contains a '{'

## AzureCP 15.2.20191031.1012 enhancements & bug-fixes - Published in October 31, 2019

* Fix a bug where guest users are not returned correctly due to bad initialization of IdentityClaimTypeConfig
* Add method AzureCPConfig.CreateDefaultConfiguration
* Update NuGet package Microsoft.Graph to v1.18
* Update NuGet package Microsoft.IdentityModel.Clients.ActiveDirectory to v5.2.3
* Update NuGet package NUnit3TestAdapter to v3.15.1

## AzureCP 15.1.20190715.926 enhancements & bug-fixes - Published in July 16, 2019

* Add missing assembly System.ValueTuple.dll as it became a dependency of Microsoft.Graph.Core 1.15.0

## AzureCP 15.0.20190621.906 enhancements & bug-fixes - Published in June 21, 2019

* Add a default mapping to populate the email of groups
* Update text in claims mapping page to better explain settings
* Add an option to return only security-enabled groups
* Update DevOps build pipelines
* Improve code quality as per Codacy's static code analysis
* Update NuGet package Microsoft.Graph to 1.15
* Update NuGet package Microsoft.IdentityModel.Clients.ActiveDirectory to 5.0.5
* Update NuGet package NUnit from 3.11 to 3.12
* Make most of public members privates and replace them with public properties, to meet best practices
* Use reflection to copy configuration objects, whenever possible, to avoid misses when new properties are added

## AzureCP 14.0.20190307.660 enhancements & bug-fixes - Published in March 13, 2019

* Fix augmentating failing for guest accounts due to non-URL encoded filter
* Add more strict checks on the claim type passed during augmentation and validation, to record a more meaningful error if needed
* Add test to ensure that AzureCP augments only entities issued from the TrustedProvider it is associated with
* Fix sign-in of users failing if AzureCP configuration does not exist
* Fix msbuild warnings
* Improve tests
* Use Azure DevOps to build AzureCP
* Cache result returned by FileVersionInfo.GetVersionInfo() to avoid potential hangs
* Add property AzureCPConfig.MaxSearchResultsCount to set max number of results returned to SharePoint during a search
* Remove reference on Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll
* Update NuGet package Microsoft.Graph to v1.12
* Update NuGet package Microsoft.IdentityModel.Clients.ActiveDirectory to v4.4.2
* Update NuGet package System.Net.Http to v4.3.4
* Update NuGet package NUnit to v3.11
* Update NuGet package NUnit3TestAdapter to v3.13
* Update NuGet package CsvTools to v1.0.12

## AzureCP v13 enhancements & bug-fixes - Published in August 30, 2018

* Guest users are now fully supported. By default, AzureCP will use the Mail property to search Guest accounts and create their permissions in SharePoint
* The identity claim type set in the SPTrustedIdentityTokenIssuer is now automatically detected and associated with the property UserPrincipalName
* Fixed no result returned under high load, caused by a thread safety issue where the same filter was used in all threads regardless of the actual input
* Improved validation of changes made to ClaimTypes collection
* Added method ClaimTypeConfigCollection.GetByClaimType()
* Implemented unit tests
* Explicitly encode HTML messages shown in admin pages and rendered from server side code to comply with tools scanning code to detect security vulnerabilities
* Fixed display text of groups that were not using the expected format "(GROUP) groupname"
* Deactivating farm-scoped feature "AzureCP" removes the claims provider from the farm, but it does not delete its configuration anymore. Configuration is now deleted when feature is uninstalled (typically when retracting the solution)
* Added user identifier properties in global configuration page

## AzureCP v12 enhancements & bug-fixes - Published in June 7, 2018

* Improved: AzureCP now uses the unified [Microsoft Graph API](https://developer.microsoft.com/en-us/graph/) instead of the old [Azure AD Graph API](https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-graph-api).
* New: AzureCP can be entirely configured with PowerShell, including claim types configuration
* Updated: AzureCP administration pages are now created as User Controls and are reusable by developers.
* Improved: AzureCP claims mapping page is easier to understand
* Updated: Logging is more relevant and generates less messages.
* New: Timeout of connection is 4 secs and can be customized
* Improved: When multiple tenants are set, they are queried in parallel
* Updated: Tenant ID is no longer needed to register a tenant
* Improved: Nested groups are now supported, if group permissions are created using the ID of groups (new default setting).
* **Beaking change**: By default, AzureCP now creates group permissions using the Id of the group instead of its name (group IDs are unique, not names). There are 2 ways to deal with this: Modify group claim type configuration to reuse the name, or migrate existing groups permissions to set their values with their group ID
* **Beaking change**: Due to the amount of changes in this area, the claim types configuration will be reset if you update from an earlier version.
* Many bug fixes and optimizations
