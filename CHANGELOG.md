# Change log for AzureCP

## AzureCP v13 enhancements & bug-fixes

* Guest users are now fully supported. AzureCP will use the Mail property to search Guest accounts and create their permissions in SharePoint
* The identity claim type set in the SPTrustedIdentityTokenIssuer is now automatically detected and associated with the UserPrincipalName property
* Fixed no result returned under high load, caused by a thread safety issue where the same filter was used in all threads regardless of the actual input
* Improved validation of changes made to ClaimTypes collection
* Added method ClaimTypeConfigCollection.GetByClaimType()
* Implemented unit tests
* Explicitely encode HTML messages shown in admin pages and renderred from server side code to comply with tools scanning code to detect security vulnerabilities