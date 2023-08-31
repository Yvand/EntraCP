# Yvand.AzureCP

This package allows you to create your own claims provider for SharePoint Server, based on [AzureCP](https://github.com/Yvand/AzureCP). This is useful in the following situations:

* You need to use AzureCP on multiple trusts in the same  SharePoint farm.
* You need more control on the permissions created by AzureCP.

## Prerequisites

* SharePoint Server 2016 or newer.
* .NET Framework 4.8.

## Getting started

Follow these steps to use this package:

* Create a Visual Sutdio project based on the template "SharePoint 2019 * Empty Project"
* Install Yvand.AzureCP via NuGet:
  * Search for `Yvand.AzureCP` in the NuGet Library, or
  * Type `Install-Package Yvand.AzureCP` into the Package Manager Console.
* Create a class that inherits `Yvand.AzureCP`
* Create a farm feature with an event receiver that inherits SPClaimProviderFeatureReceiver, to register your claims provider in the farm when activating the feature

## Documentation and resources

* [Overview](https://azurecp.yvand.net/)
* [Get help](https://github.com/Yvand/AzureCP/issues)
* [Release notes](https://github.com/Yvand/AzureCP/blob/master/CHANGELOG.md)

## License

Copyright © 2023, Yvan Duhamel * yvandev@outlook.fr. All Rights Reserved. Licensed under the Apache 2.0 [license](https://github.com/Yvand/AzureCP/blob/master/LICENSE).
