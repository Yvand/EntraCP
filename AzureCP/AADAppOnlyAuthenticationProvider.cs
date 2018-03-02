using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace azurecp
{
    public class AADAppOnlyAuthenticationProvider : IAuthenticationProvider
    {
        static string GraphAPIResource = "https://graph.microsoft.com/";
        private string aadInstance;
        private string tenant;
        private string clientId;
        private string appKey;
        private string authority;

        AuthenticationContext authContext;
        ClientCredential creds;
        private AuthenticationResult authResult;

        public AADAppOnlyAuthenticationProvider(string aadInstance, string tenant, string clientId, string appKey)
        {
            this.aadInstance = aadInstance;
            this.tenant = tenant;
            this.clientId = clientId;
            this.appKey = appKey;
            this.authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            string clientId = this.clientId;
            string clientSecret = this.appKey;

            AzureCPLogging.Log($"Getting new access token for tenant '{tenant}'", TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
            Stopwatch timer = new Stopwatch();
            timer.Start();

            //AuthenticationContext authContext = new AuthenticationContext("https://login.windows.net/yvandev.onmicrosoft.com/oauth2/token");
            authContext = new AuthenticationContext(authority);
            creds = new ClientCredential(clientId, clientSecret);
            authResult = await authContext.AcquireTokenAsync(GraphAPIResource, creds);
            request.Headers.Add("Authorization", "Bearer " + authResult.AccessToken);

            timer.Stop();
            AzureCPLogging.Log($"Got new access token for tenant '{tenant}' valid until {authResult.ExpiresOn.ToLocalTime().ToString()} local time, retrieved in {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
        }        
    }
}
