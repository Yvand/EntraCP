using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Yvand.EntraClaimsProvider
{
    /// <summary>
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
            string requestBody = await httpRequest.Content.ReadAsStringAsync().ConfigureAwait(false);
            Logger.Log($"Sent Graph request to {httpRequest.RequestUri}: '{requestBody}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
            HttpResponseMessage response = await base.SendAsync(httpRequest, cancellationToken);
            if (response.IsSuccessStatusCode == false)
            {
                Logger.Log($"Graph returned an error: {response.ReasonPhrase}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
            }
            Logger.Log($"Graph response status: {response.ReasonPhrase}", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
            return response;
        }
    }
}
