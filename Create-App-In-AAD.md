## Create the app in Azure Active Directory

AzureCP must be registered as an application in Azure AD to be able to connect and make queries. It only needs permission "Read directory data" with no deleated permission.

- Sign-in to the [Azure portal](https://portal.azure.com/) and browse to your Azure AD directory.
- Go to Properties and note the "Directory ID": this is the "Tenant ID" in AzureCP
- Go to "App Registrations" > "New application registration" > Fill following information:

Name: e.g. AzureCP
Application type: "Web app / API"
Sign-out URL: e.g. http://fake

- Click Create
- Click on the app created > Note the "Application ID"
- Click on "Required permissions" > Click on "Windows Azure Active Directory": Make sure **only** "Read directory data" permission under "Application permissions" is selected > Save
- Click "Grant Permissions" > Yes
- Click on Keys: In "Password" section type a key description (e.g. "AzureCP for SP"), choose a Duration and Save. Make a Copy of the secret.
