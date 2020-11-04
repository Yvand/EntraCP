using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Administration;
using Nito.AsyncEx;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using static azurecp.ClaimsProviderLogging;

namespace azurecp
{
    public class AADAppOnlyAuthenticationProvider : IAuthenticationProvider
    {
        private readonly string Tenant;
        private readonly string ClientId;
        private readonly string ClientSecret;
        private readonly string ClaimsProviderName;
        private readonly int Timeout;
        private readonly X509Certificate2 ClientCertificate;
        private readonly List<string> Scopes;
        private readonly AzureCloudInstance CloudInstance;
        public readonly string GraphServiceEndpoint;

        private AuthenticationResult AuthNResult;
        private AsyncLock GetAccessTokenLock = new AsyncLock();

        public AADAppOnlyAuthenticationProvider(AzureCloudInstance cloudInstance, string tenant, string clientId, string appKey, string claimsProviderName, int timeout)
        {
            this.Tenant = tenant;
            this.ClientId = clientId;
            this.ClientSecret = appKey;
            this.ClaimsProviderName = claimsProviderName;
            this.Timeout = timeout;
            //Enum.TryParse("", out AzureCloudInstance CloudInstance);
            this.CloudInstance = cloudInstance;

            this.GraphServiceEndpoint = ClaimsProviderConstants.AzureCloudEndpoints.SingleOrDefault(kvp => kvp.Key == cloudInstance).Value;
            UriBuilder scopeBuilder = new UriBuilder(this.GraphServiceEndpoint);
            scopeBuilder.Path = "/.default";
            this.Scopes = new List<string>(1);
            this.Scopes.Add(scopeBuilder.Uri.ToString());

            ////Uri loginEndpointUri = new Uri(loginServiceEndpoint);
            //var authorityUriBuilder = new UriBuilder(loginServiceEndpoint);
            //authorityUriBuilder.Path = $"/{tenant}";
            //this.AuthorityUri = authorityUriBuilder.ToString();
        }

        public AADAppOnlyAuthenticationProvider(AzureCloudInstance cloudInstance, string tenant, string clientId, X509Certificate2 ClientCertificate, string claimsProviderName, int timeout)
        {
            this.Tenant = tenant;
            this.ClientId = clientId;
            this.ClientCertificate = ClientCertificate;
            this.ClaimsProviderName = claimsProviderName;
            this.Timeout = timeout;
            //Enum.TryParse("", out AzureCloudInstance cloudInstance);
            this.CloudInstance = cloudInstance;

            this.GraphServiceEndpoint = ClaimsProviderConstants.AzureCloudEndpoints.SingleOrDefault(kvp => kvp.Key == cloudInstance).Value;
            UriBuilder scopeBuilder = new UriBuilder(this.GraphServiceEndpoint);
            scopeBuilder.Path = "/.default";
            this.Scopes = new List<string>(1);
            this.Scopes.Add(scopeBuilder.Uri.ToString());

            ////Uri loginEndpointUri = new Uri(loginServiceEndpoint);
            //var authorityUriBuilder = new UriBuilder(loginServiceEndpoint);
            //authorityUriBuilder.Path = $"/{tenant}";
            //this.AuthorityUri = authorityUriBuilder.ToString();
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
                    ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Access token for tenant '{Tenant}' expired, renewing it...", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);
                    getAccessToken = true;
                }

                if (getAccessToken)
                {
                    bool success = await GetAccessToken(false).ConfigureAwait(false);
                }

                if (AuthNResult != null && !String.IsNullOrEmpty(AuthNResult.AccessToken))
                {
                    request.Headers.Add("Authorization", $"Bearer {AuthNResult.AccessToken}");
                }
            }
        }

        public async Task<bool> GetAccessToken(bool throwExceptionIfFail)
        {
            bool success = true;
            Stopwatch timer = new Stopwatch();
            timer.Start();
            int timeout = this.Timeout;
            try
            {
                AzureCloudInstance cloudInstance = AzureCloudInstance.AzurePublic;
                ConfidentialClientApplicationBuilder appBuilder = ConfidentialClientApplicationBuilder.Create(ClientId).WithAuthority(cloudInstance, this.Tenant);
                IConfidentialClientApplication app = null;
                if (!String.IsNullOrWhiteSpace(ClientSecret))
                {
                    // Get bearer token using a client secret
                    ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Getting new access token for tenant '{Tenant}' using client ID {ClientId} and a client secret.", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);
                    app = appBuilder.WithClientSecret(ClientSecret).Build();
                }
                else
                {
                    // Get bearer token using a client certificate
                    ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Getting new access token for tenant '{Tenant}' using client ID {ClientId} and a client certificate with thumbprint {ClientCertificate.Thumbprint}.", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);
                    app = appBuilder.WithCertificate(ClientCertificate).Build();
                }
                // Acquire bearer token
                Task<AuthenticationResult> acquireTokenTask = app.AcquireTokenForClient(this.Scopes).ExecuteAsync();
                AuthNResult = await TaskHelper.TimeoutAfter<AuthenticationResult>(acquireTokenTask, new TimeSpan(0, 0, 0, 0, timeout)).ConfigureAwait(false);
                TimeSpan duration = new TimeSpan(AuthNResult.ExpiresOn.UtcTicks - DateTime.Now.ToUniversalTime().Ticks);
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Got new access token for tenant '{Tenant}', valid for {Math.Round((duration.TotalHours), 1)} hour(s) and retrieved in {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            }
            catch (MsalServiceException ex)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Unable to get access token for tenant '{Tenant}': {ex.Message}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                success = false;
                if (throwExceptionIfFail) { throw; }
            }
            catch (TimeoutException)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Could not get access token before timeout of {timeout.ToString()} ms for tenant '{Tenant}'", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                success = false;
                if (throwExceptionIfFail) { throw; }
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ClaimsProviderName, $"while getting access token for tenant '{Tenant}'", TraceCategory.Lookup, ex);
                success = false;
                if (throwExceptionIfFail) { throw; }
            }
            finally
            {
                timer.Stop();
            }
            return success;
        }
    }

    public static class TaskHelper
    {
        /// <summary>
        /// Use extension method documented in https://stackoverflow.com/questions/4238345/asynchronously-wait-for-taskt-to-complete-with-timeout
        /// </summary>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="task"></param>
        /// <param name="timeout"></param>
        /// <returns></returns>
        public static async Task<TResult> TimeoutAfter<TResult>(this Task<TResult> task, TimeSpan timeout)
        {
            using (var timeoutCancellationTokenSource = new CancellationTokenSource())
            {
                var completedTask = await Task.WhenAny(task, Task.Delay(timeout, timeoutCancellationTokenSource.Token)).ConfigureAwait(false);
                if (completedTask == task)
                {
                    timeoutCancellationTokenSource.Cancel();
                    return await task.ConfigureAwait(false);  // Very important in order to propagate exceptions
                }
                else
                {
                    throw new TimeoutException("The operation has timed out.");
                }
            }
        }
    }
}
