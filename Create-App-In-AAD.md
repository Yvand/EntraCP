## Create the app in Azure Active Directory

AzureCP must be registered as an application in Azure AD to be able to connect and make queries. It only requires permission "Read directory data" and must not have any delegated permission.

To register the application in Azure Active Directory:

- Sign-in to the [Azure portal](https://portal.azure.com/) and browse to your Azure Active Directory tenant
- Go to "App Registrations" > "New application registration" > Add following information:
    > Name: e.g. AzureCP  
    > Application type: "Web app / API"  
    > Sign-out URL: e.g. http://fake
- Click Create
- Click on the app created
    > **Note:** Copy the "Application ID" to use it in new tenant form in AzureCP.
- Click on "Required permissions" > Click on "Windows Azure Active Directory": Make sure **only** "Read directory data" permission under "Application permissions" is selected > Save
- Click "Grant Permissions" > Yes
- Click on Keys: In "Password" section type a key description (e.g. "AzureCP for SP"), choose a Duration and Save.
    > **Note:** "Password" corresponds to the "Application Secret" in new tenant form in AzureCP.

> **Note:** If you did a mistake in permissions configuration and fix it afterwards, it may take some time before this is actually applied. Feel free to go for a coffee break before testing connection again.
