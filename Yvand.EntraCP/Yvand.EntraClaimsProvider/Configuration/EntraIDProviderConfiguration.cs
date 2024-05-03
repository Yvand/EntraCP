using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Yvand.EntraClaimsProvider.Logging;

namespace Yvand.EntraClaimsProvider.Configuration
{
    public interface IEntraIDProviderSettings
    {
        #region Base settings
        /// <summary>
        /// Gets the version of the settings
        /// </summary>
        long Version { get; }

        /// <summary>
        /// Gets the claim types and their mapping with a DirectoryObject property
        /// </summary>
        ClaimTypeConfigCollection ClaimTypes { get; }

        /// <summary>
        /// Gets or sets whether to skip Microsoft Entra ID lookup and consider any input as valid.
        /// This can be useful to keep people picker working even if connectivity with the Azure tenant is lost.
        /// </summary>
        bool AlwaysResolveUserInput { get; }

        /// <summary>
        /// Gets or sets whether to return only results that match exactly the user input (case-insensitive).
        /// </summary>
        bool FilterExactMatchOnly { get; }

        /// <summary>
        /// Gets or sets whether to return the Microsoft Entra ID groups that the user is a member of.
        /// </summary>
        bool EnableAugmentation { get; }

        /// <summary>
        /// Gets or sets a string that will appear as a prefix of the text of each result, in the people picker.
        /// </summary>
        string EntityDisplayTextPrefix { get; }

        /// <summary>
        /// Gets or sets the timeout in milliseconds before an operation to Microsoft Entra ID is canceled.
        /// </summary>
        int Timeout { get; }

        /// <summary>
        /// This property is not used by EntraCP and is available to developers for their own needs
        /// </summary>
        string CustomData { get; }
        #endregion

        #region EntraID specific settings
        /// <summary>
        /// Gets the list of Entra ID tenants registered
        /// </summary>
        List<EntraIDTenant> EntraIDTenants { get; }

        /// <summary>
        /// Gets the proxy address used by EntraCP to connect to your Entra ID tenant
        /// </summary>
        string ProxyAddress { get; }

        /// <summary>
        /// Gets if only security-enabled groups should be returned - https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0
        /// </summary>
        bool FilterSecurityEnabledGroupsOnly { get; }

        /// <summary>
        /// Gets a list of Entra groups ID (max 18 values), separated by a comma. Users must be members of at least 1, to be searchable. Leave empty to not apply any filtering.
        /// </summary>
        string RestrictSearchableUsersByGroups { get; }

        /// <summary>
        /// Gets the lifetime in minutes of the cache which stores data from Entra ID which may be time-consuming to retrieve with each request
        /// </summary>
        int TenantDataCacheLifetimeInMinutes { get; }
        #endregion
    }

    public class EntraIDProviderSettings : IEntraIDProviderSettings
    {
        #region Base settings
        public long Version { get; set; }
        public ClaimTypeConfigCollection ClaimTypes { get; set; }
        public bool AlwaysResolveUserInput { get; set; } = false;
        public bool FilterExactMatchOnly { get; set; } = false;
        public bool EnableAugmentation { get; set; } = true;
        public string EntityDisplayTextPrefix { get; set; }
        public int Timeout { get; set; } = ClaimsProviderConstants.DEFAULT_TIMEOUT;
        public string CustomData { get; set; }
        #endregion

        #region EntraID specific settings
        public List<EntraIDTenant> EntraIDTenants { get; set; } = new List<EntraIDTenant>();
        public string ProxyAddress { get; set; }
        public bool FilterSecurityEnabledGroupsOnly { get; set; } = false;
        public string RestrictSearchableUsersByGroups { get; set; }
        public int TenantDataCacheLifetimeInMinutes { get; set; } = ClaimsProviderConstants.DefaultTenantDataCacheLifetimeInMinutes;
        #endregion

        public EntraIDProviderSettings() { }

        public static EntraIDProviderSettings GetDefaultSettings(string claimsProviderName)
        {
            EntraIDProviderSettings entityProviderSettings = new EntraIDProviderSettings
            {
                ClaimTypes = EntraIDProviderSettings.ReturnDefaultClaimTypesConfig(claimsProviderName),
            };
            return entityProviderSettings;
        }

        /// <summary>
        /// Returns the default claim types configuration list, based on the identity claim type set in the TrustedLoginProvider associated with <paramref name="claimProviderName"/>
        /// </summary>
        /// <returns></returns>
        public static ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig(string claimsProviderName)
        {
            if (String.IsNullOrWhiteSpace(claimsProviderName))
            {
                throw new ArgumentNullException(nameof(claimsProviderName));
            }

            SPTrustedLoginProvider spTrust = Utils.GetSPTrustAssociatedWithClaimsProvider(claimsProviderName);
            if (spTrust == null)
            {
                Logger.Log($"No SPTrustedLoginProvider associated with claims provider '{claimsProviderName}' was found.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                return null;
            }

            ClaimTypeConfigCollection newCTConfigCollection = new ClaimTypeConfigCollection(spTrust)
            {
                // Identity claim type. "Name" (http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name) is a reserved claim type in SharePoint that cannot be used in the SPTrust.
                //new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.UserPrincipalName, ClaimType = WIF4_5.ClaimTypes.Upn},
                new IdentityClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.UserPrincipalName, ClaimType = spTrust.IdentityClaimTypeInformation.MappedClaimType},

                // Additional properties to find user and create entity with the identity claim type (UseMainClaimTypeOfDirectoryObject=true)
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.DisplayName, UseMainClaimTypeOfDirectoryObject = true, EntityDataKey = PeopleEditorEntityDataKeys.DisplayName},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.GivenName, UseMainClaimTypeOfDirectoryObject = true}, //Yvan
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.Surname, UseMainClaimTypeOfDirectoryObject = true},   //Duhamel
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.Mail, EntityDataKey = PeopleEditorEntityDataKeys.Email, UseMainClaimTypeOfDirectoryObject = true},

                // Additional properties to populate metadata of entity created: no claim type set, EntityDataKey is set and UseMainClaimTypeOfDirectoryObject = false (default value)
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.MobilePhone, EntityDataKey = PeopleEditorEntityDataKeys.MobilePhone},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.JobTitle, EntityDataKey = PeopleEditorEntityDataKeys.JobTitle},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.Department, EntityDataKey = PeopleEditorEntityDataKeys.Department},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.OfficeLocation, EntityDataKey = PeopleEditorEntityDataKeys.Location},

                // Group
                new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, EntityProperty = DirectoryObjectProperty.Id, ClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType, EntityPropertyToUseAsDisplayText = DirectoryObjectProperty.DisplayName},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, EntityProperty = DirectoryObjectProperty.DisplayName, UseMainClaimTypeOfDirectoryObject = true, EntityDataKey = PeopleEditorEntityDataKeys.DisplayName},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, EntityProperty = DirectoryObjectProperty.Mail, EntityDataKey = PeopleEditorEntityDataKeys.Email},
            };
            return newCTConfigCollection;
        }
    }

    public class EntraIDProviderConfiguration : SPPersistedObject, IEntraIDProviderSettings
    {
        public string LocalAssemblyVersion => ClaimsProviderConstants.ClaimsProviderVersion;
        /// <summary>
        /// Gets the settings, based on the configuration stored in this persisted object
        /// </summary>
        public IEntraIDProviderSettings Settings
        {
            get
            {
                if (_Settings == null)
                {
                    _Settings = GenerateSettingsFromCurrentConfiguration();
                }
                return _Settings;
            }
        }
        private IEntraIDProviderSettings _Settings;

        #region "Base settings implemented from IEntraIDEntityProviderSettings"

        public ClaimTypeConfigCollection ClaimTypes
        {
            get
            {
                if (_ClaimTypes == null)
                {
                    _ClaimTypes = new ClaimTypeConfigCollection(ref this._ClaimTypesCollection, this.SPTrust);
                }
                return _ClaimTypes;
            }
            private set
            {
                _ClaimTypes = value;
                _ClaimTypesCollection = value == null ? null : value.innerCol;
            }
        }
        [Persisted]
        private Collection<ClaimTypeConfig> _ClaimTypesCollection;
        private ClaimTypeConfigCollection _ClaimTypes;

        public bool AlwaysResolveUserInput
        {
            get => _AlwaysResolveUserInput;
            private set => _AlwaysResolveUserInput = value;
        }
        [Persisted]
        private bool _AlwaysResolveUserInput;

        public bool FilterExactMatchOnly
        {
            get => _FilterExactMatchOnly;
            private set => _FilterExactMatchOnly = value;
        }
        [Persisted]
        private bool _FilterExactMatchOnly;

        public bool EnableAugmentation
        {
            get => _EnableAugmentation;
            private set => _EnableAugmentation = value;
        }
        [Persisted]
        private bool _EnableAugmentation = true;

        public string EntityDisplayTextPrefix
        {
            get => _EntityDisplayTextPrefix;
            private set => _EntityDisplayTextPrefix = value;
        }
        [Persisted]
        private string _EntityDisplayTextPrefix;

        public int Timeout
        {
            get
            {
                return _Timeout;
            }
            private set => _Timeout = value;
        }
        [Persisted]
        private int _Timeout = ClaimsProviderConstants.DEFAULT_TIMEOUT;

        public string CustomData
        {
            get => _CustomData;
            private set => _CustomData = value;
        }
        [Persisted]
        private string _CustomData;
        #endregion


        #region "EntraID settings implemented from IEntraIDEntityProviderSettings"
        /// <summary>
        /// Gets or sets the list of Entra ID tenants registered
        /// </summary>
        public List<EntraIDTenant> EntraIDTenants
        {
            get => _EntraIDTenants;
            private set => _EntraIDTenants = value;
        }
        [Persisted]
        private List<EntraIDTenant> _EntraIDTenants = new List<EntraIDTenant>();

        /// <summary>
        /// Gets or sets the proxy address used by EntraCP to connect to your Entra ID tenant
        /// </summary>
        public string ProxyAddress
        {
            get => _ProxyAddress;
            private set => _ProxyAddress = value;
        }
        [Persisted]
        private string _ProxyAddress;

        /// <summary>
        /// Gets or sets if only security-enabled groups should be returned - https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0
        /// </summary>
        public bool FilterSecurityEnabledGroupsOnly
        {
            get => _FilterSecurityEnabledGroupsOnly;
            private set => _FilterSecurityEnabledGroupsOnly = value;
        }
        [Persisted]
        private bool _FilterSecurityEnabledGroupsOnly = false;

        /// <summary>
        /// Gets or sets a list of Entra groups ID (max 18 values), separated by a comma. Users must be members of at least 1, to be searchable. Leave empty to not apply any filtering.
        /// </summary>
        public string RestrictSearchableUsersByGroups
        {
            get => _RestrictSearchableUsersByGroups;
            set => _RestrictSearchableUsersByGroups = value;
        }
        [Persisted]
        private string _RestrictSearchableUsersByGroups;

        /// <summary>
        /// Gets or sets the lifetime in minutes of the cache which stores data from Entra ID which may be time-consuming to retrieve with each request
        /// </summary>
        public int TenantDataCacheLifetimeInMinutes
        {
            get => _TenantDataCacheLifetimeInMinutes;
            set => _TenantDataCacheLifetimeInMinutes = value;
        }
        [Persisted]
        private int _TenantDataCacheLifetimeInMinutes = ClaimsProviderConstants.DefaultTenantDataCacheLifetimeInMinutes;
        #endregion

        #region "Other properties"
        /// <summary>
        /// Gets or sets the name of the claims provider using this settings
        /// </summary>
        public string ClaimsProviderName
        {
            get => _ClaimsProviderName;
            set => _ClaimsProviderName = value;
        }
        [Persisted]
        private string _ClaimsProviderName;

        [Persisted]
        private string ClaimsProviderVersion;

        private SPTrustedLoginProvider _SPTrust;
        protected SPTrustedLoginProvider SPTrust
        {
            get
            {
                if (this._SPTrust == null)
                {
                    this._SPTrust = Utils.GetSPTrustAssociatedWithClaimsProvider(this.ClaimsProviderName);
                }
                return this._SPTrust;
            }
        }
        #endregion

        public EntraIDProviderConfiguration() { }
        public EntraIDProviderConfiguration(string persistedObjectName, SPPersistedObject parent, string claimsProviderName) : base(persistedObjectName, parent)
        {
            this.ClaimsProviderName = claimsProviderName;
            this.Initialize();
        }

        private void Initialize()
        {
            this.InitializeDefaultSettings();
        }

        public virtual bool InitializeDefaultSettings()
        {
            this.ClaimTypes = ReturnDefaultClaimTypesConfig();
            return true;
        }

        /// <summary>
        /// Returns a TSettings from the properties of the current persisted object
        /// </summary>
        /// <returns></returns>
        protected virtual IEntraIDProviderSettings GenerateSettingsFromCurrentConfiguration()
        {
            IEntraIDProviderSettings entityProviderSettings = new EntraIDProviderSettings()
            {
                AlwaysResolveUserInput = this.AlwaysResolveUserInput,
                ClaimTypes = this.ClaimTypes,
                CustomData = this.CustomData,
                EnableAugmentation = this.EnableAugmentation,
                EntityDisplayTextPrefix = this.EntityDisplayTextPrefix,
                FilterExactMatchOnly = this.FilterExactMatchOnly,
                Timeout = this.Timeout,
                Version = this.Version,

                // Properties specific to type IEntraSettings
                EntraIDTenants = this.EntraIDTenants,
                ProxyAddress = this.ProxyAddress,
                FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly,
                RestrictSearchableUsersByGroups = this.RestrictSearchableUsersByGroups,
                TenantDataCacheLifetimeInMinutes = this.TenantDataCacheLifetimeInMinutes,
            };
            return (IEntraIDProviderSettings)entityProviderSettings;
        }

        /// <summary>
        /// Updates tenant credentials with a new client secret, and optionnally a new client id
        /// </summary>
        /// <param name="tenantName">Name of the tenant to update</param>
        /// <param name="newClientSecret">New client secret</param>
        /// <param name="newClientId">New client id, or empty if it does not change</param>
        /// <returns>True if credentials were successfully updated</returns>
        /// <exception cref="ArgumentNullException"></exception>
        public bool UpdateTenantCredentials(string tenantName, string newClientSecret, string newClientId = "")
        {
            if (String.IsNullOrWhiteSpace(tenantName))
            {
                throw new ArgumentNullException(nameof(tenantName));
            }

            if (String.IsNullOrWhiteSpace(newClientSecret))
            {
                throw new ArgumentNullException(nameof(newClientSecret));
            }

            EntraIDTenant tenant = this.EntraIDTenants.FirstOrDefault(x => x.Name.Equals(tenantName, StringComparison.InvariantCultureIgnoreCase));
            if (tenant == null) { return false; }
            string clientId = String.IsNullOrWhiteSpace(newClientId) ? tenant.ClientId : newClientId;
            return tenant.SetCredentials(clientId, newClientSecret);
        }

        /// <summary>
        /// Updates tenant credentials with a new client certificate, and optionnally a new client id
        /// </summary>
        /// <param name="tenantName">Name of the tenant to update</param>
        /// <param name="newClientCertificatePfxFilePath">File path to the new client certificate</param>
        /// <param name="newClientCertificatePfxPassword">Optional password of the client certificate</param>
        /// <param name="newClientId">New client id, or empty if it does not change</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public bool UpdateTenantCredentials(string tenantName, string newClientCertificatePfxFilePath, string newClientCertificatePfxPassword = "", string newClientId = "")
        {
            if (String.IsNullOrWhiteSpace(tenantName))
            {
                throw new ArgumentNullException(nameof(tenantName));
            }

            EntraIDTenant tenant = this.EntraIDTenants.FirstOrDefault(x => x.Name.Equals(tenantName, StringComparison.InvariantCultureIgnoreCase));
            if (tenant == null) { return false; }

            string clientId = String.IsNullOrWhiteSpace(newClientId) ? tenant.ClientId : newClientId;
            return tenant.SetCredentials(clientId, newClientCertificatePfxFilePath, newClientCertificatePfxPassword);
        }

        /// <summary>
        /// If it is valid, commits the current settings to the SharePoint settings database
        /// </summary>
        public override void Update()
        {
            this.ValidateConfiguration();
            base.Update();
            Logger.Log($"Successfully updated configuration '{this.Name}' with Id {this.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// If it is valid, commits the current settings to the SharePoint settings database
        /// </summary>
        /// <param name="ensure">If true, the call will not throw if the object already exists.</param>
        public override void Update(bool ensure)
        {
            this.ValidateConfiguration();
            base.Update(ensure);
            Logger.Log($"Successfully updated configuration '{this.Name}' with Id {this.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Ensures that the current configuration is valid and can be safely persisted in the configuration database
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public virtual void ValidateConfiguration()
        {
            // In case ClaimTypes collection was modified, test if it is still valid
            if (this.ClaimTypes == null)
            {
                throw new InvalidOperationException($"Configuration is not valid because collection {nameof(ClaimTypes)} is null");
            }
            try
            {
                ClaimTypeConfigCollection testUpdateCollection = new ClaimTypeConfigCollection(this.SPTrust);
                foreach (ClaimTypeConfig curCTConfig in this.ClaimTypes)
                {
                    testUpdateCollection.Add(curCTConfig, false);
                }
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException($"Some changes made to collection {nameof(ClaimTypes)} are invalid and cannot be committed to configuration database. Inspect inner exception for more details about the error.", ex);
            }

            // Ensure identity claim type is present and valid
            ClaimTypeConfig identityClaimTypeConfig = this.ClaimTypes.GetIdentifierConfiguration(DirectoryObjectType.User);
            if (identityClaimTypeConfig == null)
            {
                throw new InvalidOperationException($"The configuration is invalid because the identity claim type configuration is missing in the collection {nameof(ClaimTypes)}, so changes cannot be committed to the configuration database.");
            }
            else if (identityClaimTypeConfig is IdentityClaimTypeConfig == false)
            {
                throw new InvalidOperationException($"The configuration is invalid because the identity claim type configuration is invalid in the collection {nameof(ClaimTypes)}, so changes cannot be committed to the configuration database.");
            }

            foreach (EntraIDTenant tenant in this.EntraIDTenants)
            {
                if (tenant == null)
                {
                    throw new InvalidOperationException($"Configuration is not valid because a tenant is null in list {nameof(EntraIDTenants)}");
                }

                if (String.IsNullOrWhiteSpace(tenant.Name))
                {
                    throw new InvalidOperationException($"Configuration is not valid because a tenant has its property {nameof(tenant.Name)} not set in list {nameof(EntraIDTenants)}");
                }

                if (String.IsNullOrWhiteSpace(tenant.ClientId))
                {
                    throw new InvalidOperationException($"Configuration is not valid because tenant \"{tenant.Name}\" has its property {nameof(tenant.ClientId)} not set in list {nameof(EntraIDTenants)}");
                }

                if (String.IsNullOrWhiteSpace(tenant.ClientSecret) && tenant.ClientCertificateWithPrivateKey == null)
                {
                    throw new InvalidOperationException($"Configuration is not valid because tenant \"{tenant.Name}\" has both properties {nameof(tenant.ClientSecret)} and {nameof(tenant.ClientCertificateWithPrivateKey)} not set in list {nameof(EntraIDTenants)}, while one must be set");
                }

                if (!String.IsNullOrWhiteSpace(tenant.ClientSecret) && tenant.ClientCertificateWithPrivateKey != null)
                {
                    throw new InvalidOperationException($"Configuration is not valid because tenant \"{tenant.Name}\" has both properties {nameof(tenant.ClientSecret)} and {nameof(tenant.ClientCertificateWithPrivateKey)} set in list {nameof(EntraIDTenants)}, while only one must be set");
                }
            }

            if (this.TenantDataCacheLifetimeInMinutes < 1)
            {
                throw new InvalidOperationException($"The configuration is invalid because property {nameof(TenantDataCacheLifetimeInMinutes)} is set to 0 or a negative value. Minimum value is 1");
            }

            if (this.RestrictSearchableUsersByGroups != null)
            {
                string[] groupsId = this.RestrictSearchableUsersByGroups.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                // Batch request size limit in Graph is 20: https://learn.microsoft.com/en-us/graph/json-batching#batch-size-limitations
                // So max is 18 + 1 to get users + 1 to get groups
                if (groupsId.Length > 18)
                {
                    throw new InvalidOperationException($"The configuration is invalid because property {nameof(RestrictSearchableUsersByGroups)} exceeds the limit of 18 groups, which would generate a batch request too big for Graph. More information in https://learn.microsoft.com/en-us/graph/json-batching#batch-size-limitations");
                }
                Guid testGuidResult = Guid.Empty;
                foreach (string groupId in groupsId)
                {
                    if (!Guid.TryParse(groupId, out testGuidResult))
                    {
                        throw new InvalidOperationException($"The configuration is invalid because property {nameof(RestrictSearchableUsersByGroups)} is not set correctly. It should be a csv list of group IDs");
                    }
                }
            }
        }

        /// <summary>
        /// Removes the current persisted object from the SharePoint configuration database
        /// </summary>
        public override void Delete()
        {
            base.Delete();
            Logger.Log($"Successfully deleted configuration '{this.Name}' with Id {this.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Override this method to allow more users to update the object. True specifies that more users can update the object; otherwise, false. The default value is false.
        /// </summary>
        /// <returns></returns>
        protected override bool HasAdditionalUpdateAccess()
        {
            return false;
        }

        /// <summary>
        /// Applies the settings passed in parameter to the local configuration
        /// </summary>
        /// <param name="settings">Settings to apply to local configuration</param>
        /// <param name="commitChangesInDatabase">Commit the updates in the local configuration to the configuration database</param>
        public virtual void ApplySettings(IEntraIDProviderSettings settings, bool commitChangesInDatabase)
        {
            if (settings == null)
            {
                return;
            }

            if (settings.ClaimTypes == null)
            {
                this.ClaimTypes = null;
            }
            else
            {
                this.ClaimTypes = new ClaimTypeConfigCollection(this.SPTrust);
                foreach (ClaimTypeConfig claimTypeConfig in settings.ClaimTypes)
                {
                    this.ClaimTypes.Add(claimTypeConfig.CopyConfiguration(), false);
                }
            }
            this.AlwaysResolveUserInput = settings.AlwaysResolveUserInput;
            this.FilterExactMatchOnly = settings.FilterExactMatchOnly;
            this.EnableAugmentation = settings.EnableAugmentation;
            this.EntityDisplayTextPrefix = settings.EntityDisplayTextPrefix;
            this.Timeout = settings.Timeout;
            this.CustomData = settings.CustomData;

            this.EntraIDTenants = settings.EntraIDTenants;
            this.FilterSecurityEnabledGroupsOnly = settings.FilterSecurityEnabledGroupsOnly;
            this.ProxyAddress = settings.ProxyAddress;
            this.RestrictSearchableUsersByGroups = settings.RestrictSearchableUsersByGroups;
            this.TenantDataCacheLifetimeInMinutes = settings.TenantDataCacheLifetimeInMinutes;

            if (commitChangesInDatabase)
            {
                this.Update();
            }
        }

        public virtual IEntraIDProviderSettings GetDefaultSettings()
        {
            return EntraIDProviderSettings.GetDefaultSettings(this.ClaimsProviderName);
        }

        /// <summary>
        /// Generate and return default configuration
        /// </summary>
        /// <returns></returns>
        public static EntraIDProviderConfiguration ReturnDefaultConfiguration(string claimsProviderName)
        {
            EntraIDProviderConfiguration defaultConfig = new EntraIDProviderConfiguration();
            defaultConfig.ClaimsProviderName = claimsProviderName;
            defaultConfig.ClaimTypes = EntraIDProviderSettings.ReturnDefaultClaimTypesConfig(claimsProviderName);
            return defaultConfig;
        }

        public virtual ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig()
        {
            return EntraIDProviderSettings.ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
        }

        public void ResetClaimTypesList()
        {
            ClaimTypes.Clear();
            ClaimTypes = ReturnDefaultClaimTypesConfig();
            Logger.Log($"Claim types list of configuration '{Name}' was successfully reset to default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Returns the global configuration, stored as a persisted object in the SharePoint configuration database
        /// </summary>
        /// <param name="configurationId">The ID of the configuration</param>
        /// <param name="initializeLocalSettings">Set to true to initialize the property <see cref="Settings"/></param>
        /// <returns></returns>
        public static EntraIDProviderConfiguration GetGlobalConfiguration(Guid configurationId, bool initializeLocalSettings = false)
        {
            SPFarm parent = SPFarm.Local;
            try
            {
                //IEntraIDProviderSettings settings = (IEntraIDProviderSettings)parent.GetObject(configurationName, parent.Id, typeof(EntityProviderConfiguration));
                //Conf<TSettings> settings = (Conf<TSettings>)parent.GetObject(configurationName, parent.Id, T);
                //Conf<TSettings> settings = (Conf<TSettings>)parent.GetObject(configurationName, parent.Id, typeof(Conf<TSettings>));
                EntraIDProviderConfiguration configuration = (EntraIDProviderConfiguration)parent.GetObject(configurationId);
                //if (configuration != null && initializeLocalSettings == true)
                //{
                //    configuration.RefreshSettingsIfNeeded();
                //}
                return configuration;
            }
            catch (Exception ex)
            {
                Logger.LogException(String.Empty, $"while retrieving configuration ID '{configurationId}'", TraceCategory.Configuration, ex);
            }
            return null;
        }

        public static void DeleteGlobalConfiguration(Guid configurationId)
        {
            EntraIDProviderConfiguration configuration = GetGlobalConfiguration(configurationId);
            if (configuration == null)
            {
                Logger.Log($"Configuration ID '{configurationId}' was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
                return;
            }
            configuration.Delete();
            Logger.Log($"Configuration ID '{configurationId}' was successfully deleted from configuration database", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Creates a configuration. This will delete any existing configuration which may already exist
        /// </summary>
        /// <param name="configurationID">ID of the new configuration</param>
        /// <param name="configurationName">Name of the new configuration</param>
        /// <param name="claimsProviderName">Clais provider associated with this new configuration</param>
        /// <param name="T">Type of the new configuration</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static EntraIDProviderConfiguration CreateGlobalConfiguration(Guid configurationID, string configurationName, string claimsProviderName)
        {
            if (String.IsNullOrWhiteSpace(claimsProviderName))
            {
                throw new ArgumentNullException(nameof(claimsProviderName));
            }

            if (Utils.GetSPTrustAssociatedWithClaimsProvider(claimsProviderName) == null)
            {
                return null;
            }

            // Ensure it doesn't already exists and delete it if so
            EntraIDProviderConfiguration existingConfig = GetGlobalConfiguration(configurationID);
            if (existingConfig != null)
            {
                DeleteGlobalConfiguration(configurationID);
            }

            Logger.Log($"Creating configuration '{configurationName}' with Id {configurationID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);
            //ConstructorInfo ctorWithParameters = T.GetConstructor(new[] { typeof(string), typeof(SPFarm), typeof(string) });
            //EntraIDProviderConfiguration globalConfiguration = (EntraIDProviderConfiguration)ctorWithParameters.Invoke(new object[] { configurationName, SPFarm.Local, claimsProviderName });
            //TSettings defaultSettings = globalConfiguration.GetDefaultSettings();
            EntraIDProviderConfiguration globalConfiguration = new EntraIDProviderConfiguration(configurationName, SPFarm.Local, claimsProviderName);
            IEntraIDProviderSettings defaultSettings = globalConfiguration.GetDefaultSettings();
            globalConfiguration.ApplySettings(defaultSettings, false);
            globalConfiguration.Id = configurationID;
            globalConfiguration.Update(true);
            Logger.Log($"Created configuration '{configurationName}' with Id {globalConfiguration.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            return globalConfiguration;
        }
    }
}
