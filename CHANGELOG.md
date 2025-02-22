# Change log for ~~AzureCP~~ EntraCP

## EntraCP v28.0 - enhancements & bug-fixes - Published in December 12, 2024

* Address security advisory CVE-2024-43485 (related to `System.Text.Json` 6.0.0) by bumping dependencies `Azure.Identity` and `Microsoft.Graph` to their latest version - https://github.com/Yvand/EntraCP/pull/294
* Fix a validation issue that impacted guest users, when the identifier property for guest users is the UserPrincipalName - https://github.com/Yvand/EntraCP/pull/295
* Fix the noisy logs of category "Azure Identity" due to new level `LogMsalAlways`, by recording logs with level `LogMsalAlways` as `VerboseEx` - https://github.com/Yvand/EntraCP/pull/296
* Remove the parameter EventSeverity in method Log(), as it can be deducted from parameter TraceSeverity - https://github.com/Yvand/EntraCP/pull/296

## EntraCP v27.0.20240820.36 - enhancements & bug-fixes - Published in August 21, 2024

* Ensure that restrict searchable users feature works for all members, instead of only 100 members maximum - https://github.com/Yvand/EntraCP/issues/264
* Update the script that provisions tenant with test users and groups, to be more reliable and provision 999 users (instead of 50), so tests are more realistics
* Improve tests, to make them much easier to replay from a new test environment, and with many more users and groups
* Publish a sample project that developers can use to create a custom version of EntraCP, for specific needs
* Add a [Bruno](https://www.usebruno.com/) collection to replay the requests sent to Microsoft Graph by EntraCP

## EntraCP v26.0.20240627.35 enhancements & bug-fixes - Published in June 27, 2024

* Fix an NullReferenceException in a very rare scenario where ClaimsPrincipal.Identity is null
* Add helper methods to get/delete a tenant in the configuration
* Update Azure.Identity from 1.11.2 to 1.12.0
* Update Microsoft.Graph from 5.49.0 to 5.56.0

## EntraCP v25.0.20240503.33 enhancements & bug-fixes - Published in May 3, 2024

* Update the global configuration page to give the possibility to update the credentials of tenants already registered - https://github.com/Yvand/EntraCP/pull/248
* Fix an issue where tenant credentials could not be changed from a client certificate to a client secret - https://github.com/Yvand/EntraCP/pull/248
* Prompt user for confirmation before actually deleting a tenant - https://github.com/Yvand/EntraCP/pull/248
* New feature: It is now possible to configure EntraCP, to return only users that are members of some Entra groups, configured by the administrator - https://github.com/Yvand/EntraCP/pull/243
* Fix the connection to Microsoft Graph not working when the tenant is hosted in a national cloud
* Update Azure.Identity from 1.10.4 to 1.11.2
* Update Microsoft.Graph from 5.44.0 to 5.49.0

## EntraCP v24.0.20240318.32 enhancements & bug-fixes - Published in March 18, 2024

* Fix "Edit" links fail to work on the "Claim types configuration" page - https://github.com/Yvand/EntraCP/pull/220
* Fix the verbosity of the logging which could no longer be changed and optimized it
* Improve the global configuration page, to make it various settings easier to understand
* Multiple optimizations
* Add multiple helper methods
* Update tests
* Bump deps

## EntraCP v23.0.20231121.30 enhancements & bug-fixes - Published in November 21, 2023

* Fix a bug causing a hang on the web application, when using the check permissions feature to verify the permissions of a trusted user
* SharePoint 2016: The new minimum build version required to run EntraCP is 16.0.4483.1001 - January 2017 cumulative update
* Improve the management of tenant credentials, including new helpers to renew the client secret of certificate using PowerShell
* Improve page TroubleshootEntraCP.aspx
* Add a mechanism to use custom settings instead of settings from the persisted objects
* Fix issues causing logging to not be actually written in SharePoint logs
* Add logging categories "Azure Identity" and "Graph Requests"
* Improve the readability of the errors recorded in the log
* Better verify the responses from Graph, to avoid causing additional exceptions if an error already happened
* Bump dependencies

## EntraCP v22.0.20230927.29 enhancements & bug-fixes - Published in September 27, 2023

* Rename project AzureCP to EntraCP
* Rewrite AzureCP from scratch, with many breaking changes and improvements
* Update dependencies to use latest versions of NuGet packages Microsoft.Graph and Azure.Identity
* Update .NET Framework dependency to .NET 4.8 (https://github.com/Yvand/AzureCP/pull/184)
* Update reference on Microsoft.SharePoint.dll to use the one published with SharePoint Subscription RTM (https://github.com/Yvand/AzureCP/pull/184)
* Update minimum SharePoint Product Version from 15.0 to 16.0 (https://github.com/Yvand/AzureCP/pull/184)
* Change the deployment server type of the wsp solution to "ApplicationServer" (https://github.com/Yvand/AzureCP/pull/185)

## AzureCP v21.0.20230703.25 enhancements & bug-fixes - Published in July 3, 2023

* Fixed "Unsupported or invalid query filter" when AzureCP queries some properties like 'companyName' - (https://github.com/Yvand/AzureCP/pull/172 and https://github.com/Yvand/AzureCP/pull/177)
* Add support for extension attributes ([#174](https://github.com/Yvand/AzureCP/pull/174))
* Fix multiple reliability issues when using a certificate to authenticate in Azure AD
* Fix the properties of the identity claim type not updated (https://github.com/Yvand/AzureCP/pull/171)
* Switch from Azure DevOps to GitHub Actions to build the project and create the unit tests environments

## AzureCP 20.0.20220421.1391 enhancements & bug-fixes - Published in April 21, 2022

* Fix: In claims configurfation page, the values in the list of "PickerEntity metadata" was not populated correctly, which caused an issue with the "Title" (and a few others)
* Fix: Ensure augmentation can continue even if a tenant has a problem
* Reorganize AzureCP.csproj file
* Add a link to the privacy policy
* Explicitely set build property DependsOnNETStandard to false to try to get rid of occasional FileNotFoundException error "Could not load file or assembly 'netstandard, Version=2.0.0.0"
* Update NuGet package Microsoft.Graph to 3.35
* Update NuGet package Microsoft.Identity.Client to 4.42.1
* Update NuGet package Nito.AsyncEx to 5.1.2
* Update NuGet package NUnit to 3.13.3
* Update NuGet package NUnit3TestAdapter to 4.2.1

## AzureCP 19.0.20210211.1285 enhancements & bug-fixes - Published in February 11, 2021

* Fix bug: No Azure AD group was returned when FilterSecurityEnabledGroupsOnly is set to true - https://github.com/Yvand/AzureCP/issues/109
* Fix bug: Setting "Filter out Guest users" in new AAD teant dialog excludes members instead of guests - https://github.com/Yvand/AzureCP/issues/107
* Add ConfigureAwait(false) to Task.WhenAll()
* Update the text in AzureCP global admin page
* Better handle scenario when AzureCP cannot get an access token for a tenant
* Update NuGet package Microsoft.Graph.3.19.0 -> Microsoft.Graph.3.23.0
* Update NuGet package Microsoft.Identity.Client.4.22.0 -> Microsoft.Identity.Client.4.25.0
* Update NuGet package NUnit.3.12.0 -> Microsoft.Identity.Client.3.13.1

## AzureCP 18.0.20201120.1245 enhancements & bug-fixes - Published in November 24, 2020

* IMPORTANT: due to its dependencies and issues specific to .NET 4.6.1 (and lower) with .NET Standard 2.0, AzureCP 18 requires at least .NET 4.7.2.
* Replace authentication library ADAL (deprecated) with MSAL.NET (recommended), using Nuget package Microsoft.Identity.Client 4.22.0
* Replace NuGet package Nito.AsyncEx.StrongNamed (outdated) with up to date Nuget package Nito.AsyncEx 5.1.0.
* Add support for national clouds
* Fix queries on multiple Azure AD tenants that were not actually made in parallel
* Add method ValidateEntities to allow developers to inspect the entities generated by AzureCP, and remove some before they are returned to SharePoint.
* Make users searchable by mail in the people picker
* Update NuGet package Microsoft.Graph.3.8.0 -> Microsoft.Graph.3.19.0
* Update NuGet package Microsoft.Graph.Core.1.20.1 -> Microsoft.Graph.Core.1.22.0
* Update NuGet package NUnit3TestAdapter.3.16.1 -> NUnit3TestAdapter.3.17.0
* Remove several useless NuGet packages and references

## AzureCP 17.0.20200618.1133 enhancements & bug-fixes - Published in June 22, 2020

* IMPORTANT: due to its dependency to Microsoft.Graph 3+, AzureCP 17 requires at least .NET 4.6.1 (SharePoint requires only .NET 4.5)
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
