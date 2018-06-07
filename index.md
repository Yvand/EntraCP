This claims provider uses [Microsoft Graph](https://developer.microsoft.com/en-us/graph/) to connect SharePoint 2013 and 2016 with Azure Active Directory and provide a great search experience in the people picker.  
![People picker with AzureCP](https://github.com/Yvand/AzureCP/raw/gh-pages/assets/people%20picker%20AzureCP_2.png)

[This article](https://docs.microsoft.com/en-us/office365/enterprise/using-azure-ad-for-sharepoint-server-authentication) explains how to federate SharePoint 2013 / 2016 with Azure AD.

## Features

- Query Azure AD to get users and groups based on the user input.
- Get group membership of Azure AD users.
- Connect to multiple Azure AD tenants in parallel (multi-threaded queries).
- Easy to configure with administration pages or using PowerShell.

## Customization capabilities

- Configure the list of claim types, their mapping with Azure AD users and groups, and many other settings.
- Enable/disable augmentation.
- Enable/disable Azure AD lookup, to keep AzureCP running with limited functionality if connectivity with Azure AD is lost.
- Customize display of results in the people picker.
- Developers can deeply customize AzureCP to meet specific needs.
