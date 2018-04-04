using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Administration;
using Nito.AsyncEx;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using static azurecp.AzureCPLogging;

namespace azurecp
{
    public class AADAppOnlyAuthenticationProvider : IAuthenticationProvider
    {
        static string GraphAPIResource = "https://graph.microsoft.com/";
        private string AzureADInstance;
        private string Tenant;
        private string ClientId;
        private string ClientSecret;
        private string Authority;

        AuthenticationContext AuthContext;
        ClientCredential Creds;
        private AuthenticationResult AuthResult;

        private AsyncLock GetAccessTokenLock = new AsyncLock();

        public AADAppOnlyAuthenticationProvider(string aadInstance, string tenant, string clientId, string appKey)
        {
            this.AzureADInstance = aadInstance;
            this.Tenant = tenant;
            this.ClientId = clientId;
            this.ClientSecret = appKey;
            this.Authority = String.Format(CultureInfo.InvariantCulture, aadInstance, tenant);
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            using (GetAccessTokenLock.Lock())
            {
                bool getAccessToken = false;

                if (AuthResult == null)
                {
                    getAccessToken = true;
                }
                else if (DateTime.Now.ToUniversalTime().Ticks > AuthResult.ExpiresOn.UtcDateTime.Subtract(TimeSpan.FromMinutes(1)).Ticks)
                {
                    // Access token will expire within 1 min, let's renew it
                    AzureCPLogging.Log($"Access token for tenant '{Tenant}' expired, renewing it...", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);
                    getAccessToken = true;
                }

                if (getAccessToken) await GetAccessToken();

                request.Headers.Add("Authorization", "Bearer " + AuthResult.AccessToken);
            }
        }

        private async Task GetAccessToken()
        {
            AzureCPLogging.Log($"Getting new access token for tenant '{Tenant}'", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);
            Stopwatch timer = new Stopwatch();
            timer.Start();

            //AuthenticationContext authContext = new AuthenticationContext("https://login.windows.net/yvandev.onmicrosoft.com/oauth2/token");
            AuthContext = new AuthenticationContext(Authority);
            Creds = new ClientCredential(ClientId, ClientSecret);
            AuthResult = await AuthContext.AcquireTokenAsync(GraphAPIResource, Creds);

            timer.Stop();
            TimeSpan duration = new TimeSpan(AuthResult.ExpiresOn.UtcTicks - DateTime.Now.ToUniversalTime().Ticks);
            AzureCPLogging.Log($"Got new access token for tenant '{Tenant}', valid for {Math.Round((duration.TotalHours), 1)} hour(s) and retrieved in {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
        }
    }
}
