using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Administration;
using Nito.AsyncEx;
using System;
using System.Diagnostics;
using System.Globalization;
using System.Net.Http;
using System.Threading.Tasks;
using static azurecp.ClaimsProviderLogging;

namespace azurecp
{
    public class AADAppOnlyAuthenticationProvider : IAuthenticationProvider
    {
        private string Tenant;
        private string ClientId;
        private string ClientSecret;
        private string AuthorityUri;

        private AuthenticationContext AuthContext;
        private ClientCredential Creds;
        private AuthenticationResult AuthNResult;
        private AsyncLock GetAccessTokenLock = new AsyncLock();

        public AADAppOnlyAuthenticationProvider(string authorityUriTemplate, string tenant, string clientId, string appKey)
        {
            this.Tenant = tenant;
            this.ClientId = clientId;
            this.ClientSecret = appKey;
            this.AuthorityUri = String.Format(CultureInfo.InvariantCulture, authorityUriTemplate, tenant);
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            using (GetAccessTokenLock.Lock())
            {
                bool getAccessToken = false;
                if (AuthNResult == null)
                {
                    getAccessToken = true;
                }
                else if (DateTime.Now.ToUniversalTime().Ticks > AuthNResult.ExpiresOn.UtcDateTime.Subtract(TimeSpan.FromMinutes(1)).Ticks)
                {
                    // Access token already expired or will expire within 1 min, let's renew it
                    ClaimsProviderLogging.Log($"Access token for tenant '{Tenant}' expired, renewing it...", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);
                    getAccessToken = true;
                }

                if (getAccessToken)
                {
                    bool success = await GetAccessToken(false);
                }

                if (!String.IsNullOrEmpty(AuthNResult.AccessToken))
                {
                    request.Headers.Add("Authorization", $"Bearer {AuthNResult.AccessToken}");
                }
            }
        }

        public async Task<bool> GetAccessToken(bool throwExceptionIfFail)
        {
            ClaimsProviderLogging.Log($"Getting new access token for tenant '{Tenant}'", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);
            bool success = true;
            Stopwatch timer = new Stopwatch();
            timer.Start();
            try
            {
                AuthContext = new AuthenticationContext(AuthorityUri);
                Creds = new ClientCredential(ClientId, ClientSecret);
                AuthNResult = await AuthContext.AcquireTokenAsync(ClaimsProviderConstants.GraphAPIResource, Creds);

                TimeSpan duration = new TimeSpan(AuthNResult.ExpiresOn.UtcTicks - DateTime.Now.ToUniversalTime().Ticks);
                ClaimsProviderLogging.Log($"Got new access token for tenant '{Tenant}', valid for {Math.Round((duration.TotalHours), 1)} hour(s) and retrieved in {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            }
            catch (AdalServiceException ex)
            {
                ClaimsProviderLogging.Log($"Unable to get access token for tenant '{Tenant}': {ex.Message}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                success = false;
                if (throwExceptionIfFail) throw ex;
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(String.Empty, $"while getting access token for tenant '{Tenant}'", TraceCategory.Lookup, ex);
                if (throwExceptionIfFail) throw ex;
            }
            finally
            {
                timer.Stop();
            }
            return success;
        }
    }
}
