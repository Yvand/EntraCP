using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;

namespace Yvand.ClaimsProviders.Config
{
    public interface IEntityProviderSettings
    {
        long Version { get; }
        string Name { get; }
        string ClaimsProviderName { get; }
        ClaimTypeConfigCollection ClaimTypes { get; }
        bool AlwaysResolveUserInput { get; }
        bool FilterExactMatchOnly { get; }
        bool EnableAugmentation { get; }
        string EntityDisplayTextPrefix { get; }
        int Timeout { get; }
        string CustomData { get; }

        //string OriginalIssuerName { get; }
        //SPTrustedLoginProvider SPTrust { get; }
        List<ClaimTypeConfig> RuntimeClaimTypesList { get; }
        IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; }
        IdentityClaimTypeConfig IdentityClaimTypeConfig { get; }
        ClaimTypeConfig MainGroupClaimTypeConfig { get; }
    }

    public class EntityProviderSettings : IEntityProviderSettings
    {
        public long Version { get; set; }

        public string Name { get; set; }

        public string ClaimsProviderName { get; set; }

        public ClaimTypeConfigCollection ClaimTypes { get; set; }

        public bool AlwaysResolveUserInput { get; set; }

        public bool FilterExactMatchOnly { get; set; }

        public bool EnableAugmentation { get; set; }

        public string EntityDisplayTextPrefix { get; set; }

        public int Timeout { get; set; }

        public string CustomData { get; set; }

        public List<ClaimTypeConfig> RuntimeClaimTypesList { get; }

        public IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; }

        public IdentityClaimTypeConfig IdentityClaimTypeConfig { get; }

        public ClaimTypeConfig MainGroupClaimTypeConfig { get; }

        public EntityProviderSettings() { }

        public EntityProviderSettings(List<ClaimTypeConfig> runtimeClaimTypesList, IEnumerable<ClaimTypeConfig> runtimeMetadataConfig, IdentityClaimTypeConfig identityClaimTypeConfig, ClaimTypeConfig mainGroupClaimTypeConfig)
        {
            RuntimeClaimTypesList = runtimeClaimTypesList;
            RuntimeMetadataConfig = runtimeMetadataConfig;
            IdentityClaimTypeConfig = identityClaimTypeConfig;
            MainGroupClaimTypeConfig = mainGroupClaimTypeConfig;
        }
    }

    public class EntityProviderConfig<TConfiguration> : SPPersistedObject
        where TConfiguration : IEntityProviderSettings
    {
        /// <summary>
        /// Gets the local configuration, which is a copy of the global configuration stored in a persisted object
        /// </summary>
        public TConfiguration LocalConfiguration { get; private set; }

        /// <summary>
        /// Gets or sets the current version of the local configuration
        /// </summary>
        protected long LocalConfigurationVersion { get; private set; } = 0;

        #region "Internal runtime settings"
        /// <summary>
        /// Configuration of claim types and their mapping with LDAP attribute/class
        /// </summary>
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
            set
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
            set => _AlwaysResolveUserInput = value;
        }
        [Persisted]
        private bool _AlwaysResolveUserInput;
        public bool FilterExactMatchOnly
        {
            get => _FilterExactMatchOnly;
            set => _FilterExactMatchOnly = value;
        }
        [Persisted]
        private bool _FilterExactMatchOnly;
        public bool EnableAugmentation
        {
            get => _EnableAugmentation;
            set => _EnableAugmentation = value;
        }
        [Persisted]
        private bool _EnableAugmentation = true;
        public string EntityDisplayTextPrefix
        {
            get => _EntityDisplayTextPrefix;
            set => _EntityDisplayTextPrefix = value;
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
            set => _Timeout = value;
        }
        [Persisted]
        private int _Timeout = ClaimsProviderConstants.DEFAULT_TIMEOUT;

        ///// <summary>
        ///// Name of the SPTrustedLoginProvider where AzureCP is enabled
        ///// </summary>
        //[Persisted]
        //public string SPTrustName;
        [Persisted]
        private string _ClaimsProviderName;
        public string ClaimsProviderName
        {
            get => _ClaimsProviderName;
            set => _ClaimsProviderName = value;
        }

        [Persisted]
        private string ClaimsProviderVersion;

        /// <summary>
        /// This property is not used by AzureCP and is available to developers for their own needs
        /// </summary>
        public string CustomData
        {
            get => _CustomData;
            set => _CustomData = value;
        }
        [Persisted]
        private string _CustomData;
        #endregion


        #region "Public runtime settings"
        private SPTrustedLoginProvider _SPTrust;
        /// <summary>
        /// Gets the SharePoint trust that has its property ClaimProviderName set to the current claims provider
        /// </summary>
        public SPTrustedLoginProvider SPTrust
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

        /// <summary>
        /// Gets the issuer formatted to be like the property SPClaim.OriginalIssuer: "TrustedProvider:TrustedProviderName"
        /// </summary>
        public string OriginalIssuerName => this.SPTrust != null ? SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, this.SPTrust.Name) : String.Empty;
        #endregion

        #region "Internal runtime settings"
        protected List<ClaimTypeConfig> RuntimeClaimTypesList { get; private set; }
        protected IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; private set; }
        protected IdentityClaimTypeConfig IdentityClaimTypeConfig { get; private set; }
        protected ClaimTypeConfig MainGroupClaimTypeConfig { get; private set; }
        #endregion

        public EntityProviderConfig() { }
        public EntityProviderConfig(string persistedObjectName, SPPersistedObject parent, string claimsProviderName) : base(persistedObjectName, parent)
        {
            this.ClaimsProviderName = claimsProviderName;
            this.Initialize();
        }

        private void Initialize()
        {
            this.InitializeDefaultSettings();
            //this.InitializeInternalRuntimeSettings();
        }

        public virtual bool InitializeDefaultSettings()
        {
            this.ClaimTypes = ReturnDefaultClaimTypesConfig();
            return true;
        }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        protected virtual bool InitializeInternalRuntimeSettings()
        {
            if (this.ClaimTypes?.Count <= 0)
            {
                Logger.Log($"[{this.ClaimsProviderName}] Cannot continue because configuration '{this.Name}' has 0 claim configured.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                return false;
            }

            bool identityClaimTypeFound = false;
            bool groupClaimTypeFound = false;
            List<ClaimTypeConfig> claimTypesSetInTrust = new List<ClaimTypeConfig>();
            // Parse the ClaimTypeInformation collection set in the SPTrustedLoginProvider
            foreach (SPTrustedClaimTypeInformation claimTypeInformation in this.SPTrust.ClaimTypeInformation)
            {
                // Search if current claim type in trust exists in ClaimTypeConfigCollection
                ClaimTypeConfig claimTypeConfig = this.ClaimTypes.FirstOrDefault(x =>
                    String.Equals(x.ClaimType, claimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                    !x.UseMainClaimTypeOfDirectoryObject &&
                    x.EntityProperty != DirectoryObjectProperty.NotSet);

                if (claimTypeConfig == null)
                {
                    continue;
                }
                ClaimTypeConfig localClaimTypeConfig = claimTypeConfig.CopyConfiguration();
                localClaimTypeConfig.ClaimTypeDisplayName = claimTypeInformation.DisplayName;
                claimTypesSetInTrust.Add(localClaimTypeConfig);
                if (String.Equals(this.SPTrust.IdentityClaimTypeInformation.MappedClaimType, localClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    // Identity claim type found, set IdentityClaimTypeConfig property
                    identityClaimTypeFound = true;
                    this.IdentityClaimTypeConfig = IdentityClaimTypeConfig.ConvertClaimTypeConfig(localClaimTypeConfig);
                }
                else if (!groupClaimTypeFound && localClaimTypeConfig.EntityType == DirectoryObjectType.Group)
                {
                    groupClaimTypeFound = true;
                    this.MainGroupClaimTypeConfig = localClaimTypeConfig;
                }
            }

            if (!identityClaimTypeFound)
            {
                Logger.Log($"[{this.ClaimsProviderName}] Cannot continue because identity claim type '{this.SPTrust.IdentityClaimTypeInformation.MappedClaimType}' set in the SPTrustedIdentityTokenIssuer '{SPTrust.Name}' is missing in the ClaimTypeConfig list.", TraceSeverity.Unexpected, EventSeverity.ErrorCritical, TraceCategory.Core);
                return false;
            }

            // Check if there are additional properties to use in queries (UseMainClaimTypeOfDirectoryObject set to true)
            List<ClaimTypeConfig> additionalClaimTypeConfigList = new List<ClaimTypeConfig>();
            foreach (ClaimTypeConfig claimTypeConfig in this.ClaimTypes.Where(x => x.UseMainClaimTypeOfDirectoryObject))
            {
                ClaimTypeConfig localClaimTypeConfig = claimTypeConfig.CopyConfiguration();
                if (localClaimTypeConfig.EntityType == DirectoryObjectType.User)
                {
                    localClaimTypeConfig.ClaimType = this.IdentityClaimTypeConfig.ClaimType;
                    localClaimTypeConfig.EntityPropertyToUseAsDisplayText = this.IdentityClaimTypeConfig.EntityPropertyToUseAsDisplayText;
                }
                else
                {
                    // If not a user, it must be a group
                    if (this.MainGroupClaimTypeConfig == null)
                    {
                        continue;
                    }
                    localClaimTypeConfig.ClaimType = this.MainGroupClaimTypeConfig.ClaimType;
                    localClaimTypeConfig.EntityPropertyToUseAsDisplayText = this.MainGroupClaimTypeConfig.EntityPropertyToUseAsDisplayText;
                    localClaimTypeConfig.ClaimTypeDisplayName = this.MainGroupClaimTypeConfig.ClaimTypeDisplayName;
                }
                additionalClaimTypeConfigList.Add(localClaimTypeConfig);
            }

            this.RuntimeClaimTypesList = new List<ClaimTypeConfig>(claimTypesSetInTrust.Count + additionalClaimTypeConfigList.Count);
            this.RuntimeClaimTypesList.AddRange(claimTypesSetInTrust);
            this.RuntimeClaimTypesList.AddRange(additionalClaimTypeConfigList);

            // Get all PickerEntity metadata with a DirectoryObjectProperty set
            this.RuntimeMetadataConfig = this.ClaimTypes.Where(x =>
                !String.IsNullOrEmpty(x.EntityDataKey) &&
                x.EntityProperty != DirectoryObjectProperty.NotSet);

            return true;
        }

        /// <summary>
        /// Ensures that property LocalConfiguration is valid and up to date
        /// </summary>
        /// <param name="configurationName"></param>
        /// <returns>return true if local configuration is valid and up to date</returns>
        public TConfiguration RefreshLocalConfigurationIfNeeded()
        {
            Guid configurationId = this.Id;
            EntityProviderConfig<TConfiguration> globalConfiguration = GetGlobalConfiguration(configurationId);

            if (globalConfiguration == null)
            {
                Logger.Log($"[{ClaimsProviderName}] Cannot continue because configuration '{configurationId}' was not found in configuration database, visit AzureCP admin pages in central administration to create it.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                this.LocalConfiguration = default(TConfiguration);
                return default(TConfiguration);
            }

            if (this.LocalConfigurationVersion == globalConfiguration.Version)
            {
                Logger.Log($"[{ClaimsProviderName}] Configuration '{configurationId}' is up to date with version {this.LocalConfigurationVersion}.",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Core);
                return this.LocalConfiguration;
            }

            Logger.Log($"[{ClaimsProviderName}] Configuration '{globalConfiguration.Name}' has new version {globalConfiguration.Version}, refreshing local copy",
                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);

            globalConfiguration.ClaimsProviderName = this.ClaimsProviderName;
            bool success = globalConfiguration.InitializeInternalRuntimeSettings();
            if (!success)
            {
                return default;
            }
            this.IdentityClaimTypeConfig = globalConfiguration.IdentityClaimTypeConfig;
            this.MainGroupClaimTypeConfig = globalConfiguration.MainGroupClaimTypeConfig;
            this.RuntimeClaimTypesList = globalConfiguration.RuntimeClaimTypesList;
            this.MainGroupClaimTypeConfig = globalConfiguration.MainGroupClaimTypeConfig;
            this.LocalConfiguration = (TConfiguration)globalConfiguration.GenerateLocalConfiguration();
#if !DEBUGx
            this.LocalConfigurationVersion = globalConfiguration.Version;
#endif

            if (this.LocalConfiguration.ClaimTypes == null || this.LocalConfiguration.ClaimTypes.Count == 0)
            {
                Logger.Log($"[{ClaimsProviderName}] Configuration '{this.LocalConfiguration.Name}' was found but collection ClaimTypes is empty. Visit AzureCP admin pages in central administration to create it.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
            }
            return this.LocalConfiguration;
        }

        public override void Update()
        {
            this.ValidateConfiguration();
            base.Update();
            Logger.Log($"Successfully updated configuration '{this.Name}' with Id {this.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        public override void Update(bool ensure)
        {
            this.ValidateConfiguration();
            // If parameter ensure is true, the call will not throw if the object already exists.
            base.Update(ensure);
            Logger.Log($"Successfully updated configuration '{this.Name}' with Id {this.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Ensures that the current configuration is valid and can be safely saved and used
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public virtual void ValidateConfiguration()
        {
            // In case ClaimTypes collection was modified, test if it is still valid before committed changes to database
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
                throw new InvalidOperationException("Some changes made to list ClaimTypes are invalid and cannot be committed to configuration database. Inspect inner exception for more details about the error.", ex);
            }
        }

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

        // This method fires 3 times in a raw just when the configurationis updated, and anyway it bypassws the logic to update only if needed (and safely in regards to thread safety)
        //protected override void OnDeserialization()
        //{
        //    base.OnDeserialization();
        //    this.InitializeInternalRuntimeSettings();
        //}

        /// <summary>
        /// Returns a read-only configuration, copied from the current configuration.
        /// </summary>
        /// <returns></returns>
        protected virtual TConfiguration GenerateLocalConfiguration()
        {
            IEntityProviderSettings entityProviderSettings = new EntityProviderSettings(
                this.RuntimeClaimTypesList,
                this.RuntimeMetadataConfig,
                this.IdentityClaimTypeConfig,
                this.MainGroupClaimTypeConfig)
            {
                ClaimsProviderName = this.ClaimsProviderName,
                AlwaysResolveUserInput = this.AlwaysResolveUserInput,
                ClaimTypes = this.ClaimTypes,
                CustomData = this.CustomData,
                EnableAugmentation = this.EnableAugmentation,
                EntityDisplayTextPrefix = this.EntityDisplayTextPrefix,
                FilterExactMatchOnly = this.FilterExactMatchOnly,
                Name = this.Name,
                Timeout = this.Timeout,
                Version = this.Version,
            };
            return (TConfiguration)entityProviderSettings;
        }

        public virtual void ApplyConfiguration(TConfiguration configuration)
        {
            this.ClaimsProviderName = configuration.ClaimsProviderName;
            this.ClaimTypes = new ClaimTypeConfigCollection(this.SPTrust);
            foreach (ClaimTypeConfig claimTypeConfig in configuration.ClaimTypes)
            {
                this.ClaimTypes.Add(claimTypeConfig.CopyConfiguration(), false);
            }
            this.AlwaysResolveUserInput = configuration.AlwaysResolveUserInput;
            this.FilterExactMatchOnly = configuration.FilterExactMatchOnly;
            this.EnableAugmentation = configuration.EnableAugmentation;
            this.EntityDisplayTextPrefix = configuration.EntityDisplayTextPrefix;
            this.Timeout = configuration.Timeout;
            this.CustomData = configuration.CustomData;
        }

        //public virtual void ResetCurrentConfiguration()
        //{
        //    throw new NotImplementedException();
        //}

        public virtual ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Returns the global configuration, stored as a persisted object in the SharePoint configuration database
        /// </summary>
        /// <param name="configurationName">The name of the configuration</param>
        /// <param name="initializeLocalConfiguration">Set to true to initialize the runtime settings</param>
        /// <returns></returns>
        public static EntityProviderConfig<TConfiguration> GetGlobalConfiguration(Guid configurationId, bool initializeLocalConfiguration = false)
        {
            SPFarm parent = SPFarm.Local;
            try
            {
                //IEntityProviderSettings configuration = (IEntityProviderSettings)parent.GetObject(configurationName, parent.Id, typeof(EntityProviderConfiguration));
                //Conf<TConfiguration> configuration = (Conf<TConfiguration>)parent.GetObject(configurationName, parent.Id, T);
                //Conf<TConfiguration> configuration = (Conf<TConfiguration>)parent.GetObject(configurationName, parent.Id, typeof(Conf<TConfiguration>));
                EntityProviderConfig<TConfiguration> configuration = (EntityProviderConfig<TConfiguration>)parent.GetObject(configurationId);
                if (configuration != null && initializeLocalConfiguration == true)
                {
                    configuration.RefreshLocalConfigurationIfNeeded();
                }
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
            EntityProviderConfig<TConfiguration> configuration = (EntityProviderConfig<TConfiguration>)GetGlobalConfiguration(configurationId);
            if (configuration == null)
            {
                Logger.Log($"Configuration ID '{configurationId}' was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
                return;
            }
            configuration.Delete();
            Logger.Log($"Configuration ID '{configurationId}' was successfully deleted from configuration database", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        public static EntityProviderConfig<TConfiguration> CreateGlobalConfiguration(Guid configurationID, string configurationName, string claimsProviderName, Type T)
        {
            if (String.IsNullOrWhiteSpace(claimsProviderName))
            {
                throw new ArgumentNullException(nameof(claimsProviderName));
            }

            // Ensure it doesn't already exists and delete it if so
            EntityProviderConfig<TConfiguration> existingConfig = GetGlobalConfiguration(configurationID);
            if (existingConfig != null)
            {
                DeleteGlobalConfiguration(configurationID);
            }

            Logger.Log($"Creating configuration '{configurationName}' with Id {configurationID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);

            ConstructorInfo ctorWithParameters = T.GetConstructor(new[] { typeof(string), typeof(SPFarm), typeof(string) });
            EntityProviderConfig<TConfiguration> config = (EntityProviderConfig<TConfiguration>)ctorWithParameters.Invoke(new object[] { configurationName, SPFarm.Local, claimsProviderName });

            config.Id = configurationID;
            // If parameter ensure is true, the call will not throw if the object already exists.
            config.Update(true);
            Logger.Log($"Created configuration '{configurationName}' with Id {config.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            return config;
        }
    }
}
