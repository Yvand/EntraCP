using Microsoft.SharePoint.Administration;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace Yvand.EntraClaimsProvider
{
    /// <summary>
    /// Capture the requests to Graph to record them in the SharePoint logs.
    /// This does not capture the authentication traffic.
    /// Doc: https://github.com/microsoftgraph/msgraph-sdk-dotnet-core/blob/dev/docs/logging-requests.md
    /// </summary>
    public class GraphRequestsLogging : DelegatingHandler
    {

        public GraphRequestsLogging()
        {
        }

        /// <summary>
        /// Sends a HTTP request.
        /// </summary>
        /// <param name="httpRequest">The <see cref="HttpRequestMessage"/> to be sent.</param>
        /// <param name="cancellationToken">The <see cref="CancellationToken"/> for the request.</param>
        /// <returns></returns>
        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage httpRequest, CancellationToken cancellationToken)
        {
            HttpResponseMessage response = await base.SendAsync(httpRequest, cancellationToken).ConfigureAwait(false);
            if (response.IsSuccessStatusCode == false)
            {
                string requestBody = await httpRequest.Content.ReadAsStringAsync().ConfigureAwait(false);
                Logger.Log($"[{EntraCP.ClaimsProviderName}] Graph returned error {response.StatusCode} {response.ReasonPhrase} on request '{httpRequest.RequestUri}' with JSON payload \"{requestBody}\"", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.GraphRequests);
            }
            else
            {
                // Log in VerboseEx because as is, those messages aren't really useful
                Logger.Log($"[{EntraCP.ClaimsProviderName}] Graph returned success {response.StatusCode} on request '{httpRequest.RequestUri}'", TraceSeverity.VerboseEx, EventSeverity.Verbose, TraceCategory.GraphRequests);
            }
            return response;
        }
    }
}
