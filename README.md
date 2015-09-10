This claims provider for SharePoint 2013 leverages [Azure AD Graph Client Library](http://www.nuget.org/packages/Microsoft.Azure.ActiveDirectory.GraphClient/) to query Azure Active Directory from the people picker. It also gets the groups of Azure users to augment their SAML token, so that permission can be granted on Azure groups.
![People picker with AzureCP](https://cloud.githubusercontent.com/assets/8788631/9786028/28938be0-57b7-11e5-8119-ea759f5c508e.png)
It is highly customizable through administration pages added in Central administration/Security:
- Connect to multiple Azure AD tenants in parallel (multi-threaded queries). 
- Customize display of permissions. 
- Populate properties of users when they are added from the people picker. 
- Customize list of claim types, and their mapping with Azure AD users or groups. 
- Augment Azure AD users with their group membership. 
- Set a keyword to bypass Azure AD lookup. E.g. input "extuser:partner@contoso.com" directly creates permission "partner@contoso.com" on claim type set for this. 
- Supports rehydration for provider hosted apps. 

# How to install AzureCP
Download [AzureCP.wsp](https://raw.githubusercontent.com/Yvand/AzureCP/master/AzureCP.wsp) and please visit [the wiki](https://github.com/Yvand/AzureCP/wiki) for important information and installation instructions.
