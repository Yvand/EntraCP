using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using static Yvand.ClaimsProviders.ClaimsProviderLogging;
using Azure.Core.Pipeline;
using System.Security.Cryptography.Xml;

namespace Yvand.ClaimsProviders.Configuration.AzureAD
{
    public class AzureTenant : SPAutoSerializingObject
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
        /// Application (client) ID of the app registration created for AzureCP in the Azure AD tenant
        /// </summary>
        public string ClientId
        {
            get => _ClientId;
            set => _ClientId = value;
        }
        [Persisted]
        private string _ClientId;

        /// <summary>
        /// Client secret of the app registration created for AzureCP in the Azure AD tenant
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
        /// Client certificate with its private key, of the app registration created for AzureCP
        /// </summary>
        public X509Certificate2 ClientCertificatePrivateKey
        {
            get
            {
                return _ClientCertificatePrivateKey;
            }
            set
            {
                if (value == null) { return; }
                _ClientCertificatePrivateKey = value;
                try
                {
                    // https://stackoverflow.com/questions/32354790/how-to-check-is-x509certificate2-exportable-or-not
                    ClientCertificatePrivateKeyRawData = value.Export(X509ContentType.Pfx, ClaimsProviderConstants.ClientCertificatePrivateKeyPassword);
                }
                catch (CryptographicException ex)
                {
                    // X509Certificate2.Export() is expected to fail if the private key is not exportable, which depends on the X509KeyStorageFlags used when creating the X509Certificate2 object
                    //ClaimsProviderLogging.LogException(AzureCP._ProviderInternalName, $"while setting the certificate for tenant '{this.Name}'. Is the private key of the certificate exportable?", TraceCategory.Core, ex);
                }
            }
        }
        private X509Certificate2 _ClientCertificatePrivateKey;
        [Persisted]
        private byte[] ClientCertificatePrivateKeyRawData;

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

        public GraphServiceClient GraphService { get; private set; }

        public string UserFilter { get; set; }
        public string GroupFilter { get; set; }
        public string[] UserSelect { get; set; }
        public string[] GroupSelect { get; set; }

        public AzureTenant()
        {
        }

        protected override void OnDeserialization()
        {
            if (ClientCertificatePrivateKeyRawData != null)
            {
                try
                {
                    // EphemeralKeySet: Keep the private key in-memory, it won't be written to disk - https://www.pkisolutions.com/handling-x509keystorageflags-in-applications/
                    _ClientCertificatePrivateKey = ImportPfxCertificateBlob(ClientCertificatePrivateKeyRawData, ClaimsProviderConstants.ClientCertificatePrivateKeyPassword, X509KeyStorageFlags.EphemeralKeySet);
                }
                catch (CryptographicException ex)
                {
                    ClaimsProviderLogging.LogException(AzureCPSE.ClaimsProviderName, $"while deserializating the certificate for tenant '{this.Name}'.", TraceCategory.Core, ex);
                }
            }
        }

        /// <summary>
        /// Set properties AuthenticationProvider and GraphService
        /// </summary>
        public void InitializeGraphForAppOnlyAuth(int timeout, string proxyAddress)
        {
            try
            {
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
                handlers.Add(new ChaosHandler());
#endif

                ClientSecretCredentialOptions options = new ClientSecretCredentialOptions();
                options.AuthorityHost = this.AzureAuthority;
                if (clientTransportProxy != null) { options.Transport = clientTransportProxy; }

                TokenCredential tokenCredential;
                if (!String.IsNullOrWhiteSpace(_ClientSecret))
                {
                    tokenCredential = new ClientSecretCredential(this.Name, this.ClientId, this.ClientSecret, options);
                }
                else
                {
                    tokenCredential = new ClientCertificateCredential(this.Name, this.ClientId, this.ClientCertificatePrivateKey, options);
                }

                var scopes = new[] { "https://graph.microsoft.com/.default" };
                HttpClient httpClient = GraphClientFactory.Create(handlers: handlers, proxy: webProxy);
                httpClient.Timeout = TimeSpan.FromMilliseconds(timeout);

                // https://learn.microsoft.com/en-us/graph/sdks/customize-client?tabs=csharp
                var authProvider = new Microsoft.Graph.Authentication.AzureIdentityAuthenticationProvider(
                    credential: tokenCredential,
                    scopes: new[] { "https://graph.microsoft.com/.default",
                });
                this.GraphService = new GraphServiceClient(httpClient, authProvider);
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(AzureCPSE.ClaimsProviderName, $"while setting client context for tenant '{this.Name}'.", TraceCategory.Core, ex);
            }
        }

        /// <summary>
        /// Returns a copy of the current object. This copy does not have any member of the base SharePoint base class set
        /// </summary>
        /// <returns></returns>
        internal AzureTenant CopyConfiguration()
        {
            AzureTenant copy = new AzureTenant();
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
        /// Update the credentials used to connect to the Azure AD tenant
        /// </summary>
        /// <param name="newApplicationSecret">New application (client) secret</param>
        public void UpdateCredentials(string newApplicationSecret)
        {
            SetCredentials(this.ClientId, newApplicationSecret);
        }

        /// <summary>
        /// Set the credentials used to connect to the Azure AD tenant
        /// </summary>
        /// <param name="applicationId">Application (client) ID</param>
        /// <param name="applicationSecret">Application (client) secret</param>
        public void SetCredentials(string applicationId, string applicationSecret)
        {
            this.ClientId = applicationId;
            this.ClientSecret = applicationSecret;
            this.ClientCertificatePrivateKey = null;
        }

        /// <summary>
        /// Update the credentials used to connect to the Azure AD tenant
        /// </summary>
        /// <param name="newCertificate">New certificate with its private key</param>
        public void UpdateCredentials(X509Certificate2 newCertificate)
        {
            SetCredentials(this.ClientId, newCertificate);
        }

        /// <summary>
        /// Set the credentials used to connect to the Azure AD tenant
        /// </summary>
        /// <param name="applicationId">Application (client) secret</param>
        /// <param name="certificate">Certificate with its private key</param>
        public void SetCredentials(string applicationId, X509Certificate2 certificate)
        {
            this.ClientId = applicationId;
            this.ClientSecret = String.Empty;
            this.ClientCertificatePrivateKey = certificate;
        }

        /// <summary>
        /// Import the input blob certificate into a pfx X509Certificate2 object
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
                // If passwordless, import private key as documented in https://support.microsoft.com/en-us/topic/kb5025823-change-in-how-net-applications-import-x-509-certificates-bf81c936-af2b-446e-9f7a-016f4713b46b
                return new X509Certificate2(blob, (string)null, keyStorageFlags);
            }
            else
            {
                return new X509Certificate2(blob, certificatePassword, keyStorageFlags);
            }
        }
    }
}
