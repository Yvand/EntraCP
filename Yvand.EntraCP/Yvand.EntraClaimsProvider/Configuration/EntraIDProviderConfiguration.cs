﻿using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

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
        /// Gets or sets the timeout before giving up the query to Microsoft Entra ID.
        /// </summary>
        int Timeout { get; }

        /// <summary>
        /// This property is not used by EntraCP and is available to developers for their own needs
        /// </summary>
        string CustomData { get; }
        #endregion

        #region EntraID specific settings
        /// <summary>
        /// Gets the list of Azure tenants to use to get entities
        /// </summary>
        List<EntraIDTenant> EntraIDTenants { get; }

        /// <summary>
        /// Gets the proxy address used by AzureCP to connect to Azure AD
        /// </summary>
        string ProxyAddress { get; }

        /// <summary>
        /// Gets if only security-enabled groups should be returned
        /// </summary>
        bool FilterSecurityEnabledGroupsOnly { get; }
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
                    _ClaimTypes = new ClaimTypeConfigCollection(ref this._ClaimTypesCollection);
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
#if DEBUG
                return _Timeout * 100;
#endif
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
        public List<EntraIDTenant> EntraIDTenants
        {
            get => _EntraIDTenants;
            private set => _EntraIDTenants = value;
        }
        [Persisted]
        private List<EntraIDTenant> _EntraIDTenants = new List<EntraIDTenant>();

        public string ProxyAddress
        {
            get => _ProxyAddress;
            private set => _ProxyAddress = value;
        }
        [Persisted]
        private string _ProxyAddress;

        /// <summary>
        /// Set if only AAD groups with securityEnabled = true should be returned - https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0
        /// </summary>
        public bool FilterSecurityEnabledGroupsOnly
        {
            get => _FilterSecurityEnabledGroupsOnly;
            private set => _FilterSecurityEnabledGroupsOnly = value;
        }
        [Persisted]
        private bool _FilterSecurityEnabledGroupsOnly = false;
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
            };
            return (IEntraIDProviderSettings)entityProviderSettings;
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
        /// Applies the settings passed in parameter to the current settings
        /// </summary>
        /// <param name="settings"></param>
        public virtual void ApplySettings(IEntraIDProviderSettings settings, bool commitIfValid)
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

            if (commitIfValid)
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
