using Azure.Core;
using Azure.Core.Pipeline;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace Yvand.EntraClaimsProvider.Config
{
    public class EntraIDTenant : SPAutoSerializingObject
    {
        public Guid Identifier
        {
            get => _Identifier;
            set => _Identifier = value;
        }
        [Persisted]
        private Guid _Identifier = Guid.NewGuid();

        /// <summary>
        /// Tenant Name or ID
        /// </summary>
        public string Name
        {
            get => _Name;
            set => _Name = value;
        }
        [Persisted]
        private string _Name;

        /// <summary>
        /// Application (client) ID of the app registration created for EntraCP in the Microsoft Entra ID tenant
        /// </summary>
        public string ClientId
        {
            get => _ClientId;
            set => _ClientId = value;
        }
        [Persisted]
        private string _ClientId;

        /// <summary>
        /// Client secret of the app registration created for EntraCP in the Microsoft Entra ID tenant
        /// </summary>
        public string ClientSecret
        {
            get => _ClientSecret;
            set => _ClientSecret = value;
        }
        [Persisted]
        private string _ClientSecret;

        public bool ExcludeMemberUsers
        {
            get => _ExcludeMemberUsers;
            set => _ExcludeMemberUsers = value;
        }
        [Persisted]
        private bool _ExcludeMemberUsers = false;

        public bool ExcludeGuestUsers
        {
            get => _ExcludeGuestUsers;
            set => _ExcludeGuestUsers = value;
        }
        [Persisted]
        private bool _ExcludeGuestUsers = false;

        /// <summary>
        /// Client ID of AD Connect used in extension attribues
        /// </summary>
        [Persisted]
        private Guid _ExtensionAttributesApplicationId;

        public Guid ExtensionAttributesApplicationId
        {
            get => _ExtensionAttributesApplicationId;
            set => _ExtensionAttributesApplicationId = value;
        }

        /// <summary>
        /// Gets or set the client certificate with its private key, configured in the app registration for EntraCP
        /// </summary>
        public X509Certificate2 ClientCertificateWithPrivateKey
        {
            get
            {
                return _ClientCertificateWithPrivateKey;
            }
            set
            {
                if (value == null) { return; }
                if (!value.HasPrivateKey) { throw new ArgumentException("The certificate cannot be imported because it does not have a private key"); }
                _ClientCertificateWithPrivateKey = value;
                try
                {
                    // https://stackoverflow.com/questions/32354790/how-to-check-is-x509certificate2-exportable-or-not
                    _ClientCertificateWithPrivateKeyRawData = value.Export(X509ContentType.Pfx, ClaimsProviderConstants.ClientCertificatePrivateKeyPassword);
                }
                catch (CryptographicException ex)
                {
                    // X509Certificate2.Export() is expected to fail if the private key is not exportable, which depends on the X509KeyStorageFlags used when creating the X509Certificate2 object
                    //ClaimsProviderLogging.LogException(EntraCP._ProviderInternalName, $"while setting the certificate for tenant '{this.Name}'. Is the private key of the certificate exportable?", TraceCategory.Core, ex);
                }
            }
        }
        private X509Certificate2 _ClientCertificateWithPrivateKey;
        [Persisted]
        private byte[] _ClientCertificateWithPrivateKeyRawData;

        public Uri AzureAuthority
        {
            get => new Uri(this._AzureAuthority);
            set => _AzureAuthority = value.ToString();
        }
        [Persisted]
        private string _AzureAuthority = AzureAuthorityHosts.AzurePublicCloud.ToString();
        public AzureCloudInstance CloudInstance
        {
            get
            {
                if (AzureAuthority == null) { return AzureCloudInstance.AzurePublic; }
                KeyValuePair<AzureCloudInstance, Uri> kvp = ClaimsProviderConstants.AzureCloudEndpoints.FirstOrDefault(item => item.Value.Equals(this.AzureAuthority));
                return kvp.Equals(default(KeyValuePair<AzureCloudInstance, Uri>)) ? AzureCloudInstance.AzurePublic : kvp.Key;
            }
        }

        public string AuthenticationMode
        {
            get
            {
                return String.IsNullOrWhiteSpace(this.ClientSecret) ? "Client certificate" : "Client secret";
            }
        }

        public GraphServiceClient GraphService { get; private set; }
        public string UserFilter { get; set; }
        public string GroupFilter { get; set; }
        public string[] UserSelect { get; set; }
        public string[] GroupSelect { get; set; }

        public EntraIDTenant() { }

        protected override void OnDeserialization()
        {
            if (_ClientCertificateWithPrivateKeyRawData != null)
            {
                try
                {
                    // Sets the local X509Certificate2 object from the persisted raw data stored in the configuration database
                    // EphemeralKeySet: Keep the private key in-memory, it won't be written to disk - https://www.pkisolutions.com/handling-x509keystorageflags-in-applications/
                    _ClientCertificateWithPrivateKey = ImportPfxCertificateBlob(_ClientCertificateWithPrivateKeyRawData, ClaimsProviderConstants.ClientCertificatePrivateKeyPassword, X509KeyStorageFlags.EphemeralKeySet);
                }
                catch (CryptographicException ex)
                {
                    // It may fail with CryptographicException: The system cannot find the file specified, but it does not have any impact
                    Logger.LogException(EntraCP.ClaimsProviderName, $"while deserializating the certificate for tenant '{this.Name}'.", TraceCategory.Core, ex);
                }
            }
        }

        /// <summary>
        /// Initializes the authentication to Microsoft Graph
        /// </summary>
        public void InitializeAuthentication(int timeout, string proxyAddress)
        {
            if (String.IsNullOrWhiteSpace(this.ClientSecret) && this.ClientCertificateWithPrivateKey == null)
            {
                Logger.Log($"[{EntraCP.ClaimsProviderName}] Cannot initialize authentication for tenant {this.Name} because both properties {nameof(ClientSecret)} and {nameof(ClientCertificateWithPrivateKey)} are not set.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                return;
            }
            if (String.IsNullOrWhiteSpace(this.ClientId))
            {
                Logger.Log($"[{EntraCP.ClaimsProviderName}] Cannot initialize authentication for tenant {this.Name} because the property {nameof(ClientId)} is not set.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                return;
            }
            if (String.IsNullOrWhiteSpace(this.Name))
            {
                Logger.Log($"[{EntraCP.ClaimsProviderName}] Cannot initialize authentication because the property {nameof(Name)} of current tenant is not set.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                return;
            }

            WebProxy webProxy = null;
            HttpClientTransport clientTransportProxy = null;
            if (!String.IsNullOrWhiteSpace(proxyAddress))
            {
                webProxy = new WebProxy(new Uri(proxyAddress));
                HttpClientHandler clientProxy = new HttpClientHandler { Proxy = webProxy };
                clientTransportProxy = new HttpClientTransport(clientProxy);
            }

            var handlers = GraphClientFactory.CreateDefaultHandlers();
#if DEBUG
            //handlers.Add(new ChaosHandler());
#endif

            ClientSecretCredentialOptions options = new ClientSecretCredentialOptions();
            options.AuthorityHost = this.AzureAuthority;
            if (clientTransportProxy != null) { options.Transport = clientTransportProxy; }

            TokenCredential tokenCredential;
            if (!String.IsNullOrWhiteSpace(this.ClientSecret))
            {
                tokenCredential = new ClientSecretCredential(this.Name, this.ClientId, this.ClientSecret, options);
            }
            else
            {
                tokenCredential = new ClientCertificateCredential(this.Name, this.ClientId, this.ClientCertificateWithPrivateKey, options);
            }

            var scopes = new[] { "https://graph.microsoft.com/.default" };
            HttpClient httpClient = GraphClientFactory.Create(handlers: handlers, proxy: webProxy);
            if (timeout > 0 && timeout < Int32.MaxValue)
            {
                httpClient.Timeout = TimeSpan.FromMilliseconds(timeout);
            }

            // https://learn.microsoft.com/en-us/graph/sdks/customize-client?tabs=csharp
            var authProvider = new Microsoft.Graph.Authentication.AzureIdentityAuthenticationProvider(
                credential: tokenCredential,
                scopes: new[] { "https://graph.microsoft.com/.default",
            });
            this.GraphService = new GraphServiceClient(httpClient, authProvider);
        }

        public async Task<bool> TestConnectionAsync(string proxyAddress)
        {
            if (this.GraphService == null)
            {
                this.InitializeAuthentication(Int32.MaxValue, proxyAddress);
            }
            if (this.GraphService == null)
            {
                return false;
            }
            bool success = true;
            try
            {
                await GraphService.Users.GetAsync((config) =>
                {
                    config.QueryParameters.Select = new[] { "Id" };
                    config.QueryParameters.Top = 1;
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                success = false;
                throw ex;
            }
            return success;
        }

        /// <summary>
        /// Returns a copy of the current object. This copy does not have any member of the base SharePoint base class set
        /// </summary>
        /// <returns></returns>
        public EntraIDTenant CopyConfiguration()
        {
            EntraIDTenant copy = new EntraIDTenant();
            copy = (EntraIDTenant)Utils.CopyPersistedFields(typeof(EntraIDTenant), this, copy);
            return copy;
        }

        public EntraIDTenant CopyPublicProperties()
        {
            EntraIDTenant copy = new EntraIDTenant();
            // Copy non-inherited public properties
            PropertyInfo[] propertiesToCopy = this.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
            foreach (PropertyInfo property in propertiesToCopy)
            {
                if (property.CanWrite)
                {
                    object value = property.GetValue(this);
                    if (value != null)
                    {
                        property.SetValue(copy, value);
                    }
                }
            }
            return copy;
        }

        /// <summary>
        /// Sets the credentials used to connect to the Microsoft Entra ID tenant
        /// </summary>
        /// <param name="clientId">Application (client) ID</param>
        /// <param name="clientSecret">Application (client) secret</param>
        public void SetCredentials(string clientId, string clientSecret)
        {
            this.ClientId = clientId;
            this.ClientSecret = clientSecret;
            this.ClientCertificateWithPrivateKey = null;
        }

        /// <summary>
        /// Sets the credentials used to connect to the Microsoft Entra ID tenant
        /// </summary>
        /// <param name="clientId">Application (client) secret</param>
        /// <param name="clientCertificateWithPrivateKey">Client certificate with its private key</param>
        public void SetCredentials(string clientId, X509Certificate2 clientCertificateWithPrivateKey)
        {
            this.ClientId = clientId;
            this.ClientSecret = String.Empty;
            this.ClientCertificateWithPrivateKey = clientCertificateWithPrivateKey;
        }

        /// <summary>
        /// Imports the input blob into a pfx X509Certificate2 object with its private key
        /// </summary>
        /// <param name="blob"></param>
        /// <param name="certificatePassword"></param>
        /// <param name="keyStorageFlags"></param>
        /// <returns></returns>
        public static X509Certificate2 ImportPfxCertificateBlob(byte[] blob, string certificatePassword, X509KeyStorageFlags keyStorageFlags)
        {
            if (X509Certificate2.GetCertContentType(blob) != X509ContentType.Pfx)
            {
                return null;
            }

            if (String.IsNullOrWhiteSpace(certificatePassword))
            {
                // If passwordless, import the private key as documented in https://support.microsoft.com/en-us/topic/kb5025823-change-in-how-net-applications-import-x-509-certificates-bf81c936-af2b-446e-9f7a-016f4713b46b
                return new X509Certificate2(blob, (string)null, keyStorageFlags);
            }
            else
            {
                return new X509Certificate2(blob, certificatePassword, keyStorageFlags);
            }
        }
    }
}
