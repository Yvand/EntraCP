## Create the app in Azure Active Directory

AzureCP must be registered as an application in Azure AD to be able to connect and make queries. It requires only permission "Read directory data", and it must not have any deleated permission.

Once AzureCP is installed in SharePoint farm, connection to Azure AD can be made and tested in Central Administration > Security > AzureCP Glogal configuration > "New Azure tenant" section.

To create the app in Azure AD:

- Sign-in to the [Azure portal](https://portal.azure.com/) and browse to your Azure AD directory.
- Go to Properties and note the "Directory ID"
    > **Note:** "Directory ID" corresponds to the "Tenant ID" in new tenant form in AzureCP.
- Go to "App Registrations" > "New application registration" > Fill following information:
    > Name: e.g. AzureCP<br>
    > Application type: "Web app / API"<br>
    > Sign-out URL: e.g. http://fake
- Click Create
- Click on the app created
    > **Note:** Copy the "Application ID" to use it in new tenant form in AzureCP.
- Click on "Required permissions" > Click on "Windows Azure Active Directory": Make sure **only** "Read directory data" permission under "Application permissions" is selected > Save
- Click "Grant Permissions" > Yes
- Click on Keys: In "Password" section type a key description (e.g. "AzureCP for SP"), choose a Duration and Save.
    > **Note:** "Password" corresponds to the "Application Secret" in new tenant form in AzureCP.

> **Note:** If you did a mistake in permissions configuration and fix it afterwards, it may take some time before this is actually applied. Feel free to go for a coffee break before testing connection again.
