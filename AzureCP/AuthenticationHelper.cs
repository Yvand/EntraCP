using System;
using System.Threading.Tasks;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Utilities;

namespace azurecp
{
    internal class AuthenticationHelper
    {
        //public static string TokenForUser;

        /// <summary>
        /// Async task to acquire token for Application.
        /// </summary>
        /// <returns>Async Token for application.</returns>
        public static async Task<string> AcquireTokenAsyncForApplication(string tenantName, string clientId, string clientSecret)
        {
            return GetTokenForApplication(tenantName, clientId, clientSecret);
        }

        /// <summary>
        /// Get Token for Application.
        /// </summary>
        /// <returns>Token for application.</returns>
        public static string GetTokenForApplication(string tenantName, string clientId, string clientSecret)
        {
            AuthenticationContext authenticationContext = new AuthenticationContext(String.Format(Constants.AuthString, tenantName), false);
            // Config for OAuth client credentials 
            ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
            Task<AuthenticationResult> authenticationResult = authenticationContext.AcquireTokenAsync(Constants.ResourceUrl, clientCred);
            string token = authenticationResult.Result.AccessToken;
            return token;
        }

        /// <summary>
        /// Get Active Directory Client for Application.
        /// </summary>
        /// <returns>ActiveDirectoryClient for Application.</returns>
        public static ActiveDirectoryClient GetActiveDirectoryClientAsApplication(string tenantName, string tenantId, string clientId, string clientSecret)
        {
            using (new SPMonitoredScope(String.Format("[AzureCP] Getting access token for tenant {0} by connecting to '{1}' ", tenantName, Constants.ResourceUrl), 1000))
            {
                Uri servicePointUri = new Uri(Constants.ResourceUrl);
                Uri serviceRoot = new Uri(servicePointUri, tenantId);
                ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
                    async () => await AcquireTokenAsyncForApplication(tenantName, clientId, clientSecret).ConfigureAwait(false));
                return activeDirectoryClient;
            }
        }

        ///// <summary>
        ///// Async task to acquire token for User.
        ///// </summary>
        ///// <returns>Token for user.</returns>
        //public static async Task<string> AcquireTokenAsyncForUser()
        //{
        //    return GetTokenForUser();
        //}

        ///// <summary>
        ///// Get Token for User.
        ///// </summary>
        ///// <returns>Token for user.</returns>
        //public static string GetTokenForUser()
        //{
        //    if (TokenForUser == null)
        //    {
        //        var redirectUri = new Uri("https://localhost");
        //        AuthenticationContext authenticationContext = new AuthenticationContext(Constants.AuthString, false);
        //        AuthenticationResult userAuthnResult = authenticationContext.AcquireToken(Constants.ResourceUrl,
        //            Constants.ClientIdForUserAuthn, redirectUri, PromptBehavior.Always);
        //        TokenForUser = userAuthnResult.AccessToken;
        //        Console.WriteLine("\n Welcome " + userAuthnResult.UserInfo.GivenName + " " +
        //                          userAuthnResult.UserInfo.FamilyName);
        //    }
        //    return TokenForUser;
        //}

        ///// <summary>
        ///// Get Active Directory Client for User.
        ///// </summary>
        ///// <returns>ActiveDirectoryClient for User.</returns>
        //public static ActiveDirectoryClient GetActiveDirectoryClientAsUser()
        //{
        //    Uri servicePointUri = new Uri(Constants.ResourceUrl);
        //    Uri serviceRoot = new Uri(servicePointUri, Constants.TenantId);
        //    ActiveDirectoryClient activeDirectoryClient = new ActiveDirectoryClient(serviceRoot,
        //        async () => await AcquireTokenAsyncForUser());
        //    return activeDirectoryClient;
        //}
    }
}
