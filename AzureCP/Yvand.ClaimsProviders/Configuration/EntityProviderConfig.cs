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

        public List<ClaimTypeConfig> RuntimeClaimTypesList { get; set; }

        public IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; set; }

        public IdentityClaimTypeConfig IdentityClaimTypeConfig { get; set; }

        public ClaimTypeConfig MainGroupClaimTypeConfig { get; set; }
    }

    public class EntityProviderConfig<TConfiguration> : SPPersistedObject
        where TConfiguration : IEntityProviderSettings//, IEntityProviderSettings // SPPersistedObject//, IPersistedEntityProviderSettings//, IRuntimeEntityProviderSettings
    {
        /// <summary>
        /// Gets or sets the local configuration, which is a copy of the global configuration stored in a persisted object
        /// </summary>
        protected TConfiguration LocalConfiguration { get; set; }

        /// <summary>
        /// Gets or sets the current version of the local configuration
        /// </summary>
        protected long LocalConfigurationVersion { get; set; } = 0;

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

        private SPTrustedLoginProvider _SPTrust;
        public SPTrustedLoginProvider SPTrust
        {
            get
            {
                if (this._SPTrust == null)
                {
                    //_SPTrust = SPSecurityTokenServiceManager.Local.TrustedLoginProviders.GetProviderByName(SPTrustName);
                    this._SPTrust = Utils.GetSPTrustAssociatedWithClaimsProvider(this.ClaimsProviderName);
                }
                return this._SPTrust;
            }
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

        ///// <summary>
        ///// Limit number of results returned to SharePoint during a search
        ///// </summary>
        //public int MaxSearchResultsCount
        //{
        //    get => _MaxSearchResultsCount;
        //    set => _MaxSearchResultsCount = value;
        //}
        //[Persisted]
        //private int _MaxSearchResultsCount = 30; // SharePoint sets maxCount to 30 in method FillSearch

        public bool RuntimeSettingsInitialized { get; private set; }

        // Runtime settings
        /// <summary>
        /// Returned issuer formatted like the property SPClaim.OriginalIssuer: "TrustedProvider:TrustedProviderName"
        /// </summary>
        public string OriginalIssuerName => this.SPTrust != null ? SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, this.SPTrust.Name) : String.Empty;
        protected List<ClaimTypeConfig> RuntimeClaimTypesList { get; private set; }
        protected IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; private set; }
        public IdentityClaimTypeConfig IdentityClaimTypeConfig { get; private set; }
        public ClaimTypeConfig MainGroupClaimTypeConfig { get; private set; }

        public EntityProviderConfig() { }
        public EntityProviderConfig(string persistedObjectName, SPPersistedObject parent, string claimsProviderName) : base(persistedObjectName, parent)
        {
            this.ClaimsProviderName = claimsProviderName;
            //this.InitializeDefaultSettings();
            //this.InitializeRuntimeSettings();
            this.Initialize();
        }

        public EntityProviderConfig(string claimsProviderName)
        {
            this.ClaimsProviderName = claimsProviderName;
            //this.InitializeDefaultSettings();
            //this.InitializeRuntimeSettings();
            this.Initialize();
        }

        private void Initialize()
        {
            //if (!this.IsInitialized)
            //{
            this.InitializeDefaultSettings();
            this.InitializeRuntimeSettings();
            //}
            //return this.IsInitialized;
        }

        protected virtual bool InitializeDefaultSettings()
        {
            this.ClaimTypes = ReturnDefaultClaimTypesConfig();
            return true;
        }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        public virtual bool InitializeRuntimeSettings()
        {
            if (this.RuntimeSettingsInitialized)
            {
                return true;
            }

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

            this.RuntimeSettingsInitialized = true;
            return true;
        }

        /// <summary>
        /// Ensure that property LocalConfiguration is valid and up to date
        /// </summary>
        /// <param name="configurationName"></param>
        /// <returns>return true if local configuration is valid and up to date</returns>
        public TConfiguration RefreshLocalConfigurationIfNeeded()
        {
            Guid configurationId = this.Id;
            EntityProviderConfig<TConfiguration> globalConfiguration = GetGlobalConfiguration(configurationId, typeof(TConfiguration));

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

            this.InitializeRuntimeSettings();
            this.LocalConfiguration = (TConfiguration)globalConfiguration.CopyConfiguration();
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
            this.RuntimeSettingsInitialized = false;
            Logger.Log($"Successfully updated configuration '{this.Name}' with Id {this.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        public override void Update(bool ensure)
        {
            this.ValidateConfiguration();
            // If parameter ensure is true, the call will not throw if the object already exists.
            base.Update(ensure);
            this.RuntimeSettingsInitialized = false;
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

        protected override void OnPostCreate()
        {
            base.OnPostCreate();
            //this.InitializeRuntimeSettings();
        }

        // This method fires 3 times in a raw just when the configurationis updated, and anyway it bypassws the logic to update only if needed (and safely in regards to thread safety)
        //protected override void OnDeserialization()
        //{
        //    base.OnDeserialization();
        //    this.InitializeRuntimeSettings();
        //}

        /// <summary>
        /// Returns a copy of the current object. This copy does not have any member of the base SharePoint base class set
        /// </summary>
        /// <returns></returns>
        public virtual TConfiguration CopyConfiguration()
        {
            // Cannot use reflection here to copy object because of the calls to methods CopyConfiguration() on some properties
            //EntityProviderConfiguration copy = new EntityProviderConfiguration(this.ClaimsProviderName);
            // Use default constructor to bypass initialization, which is useless since properties will be manually set here
            //EntityProviderConfiguration copy = new EntityProviderConfiguration();
            IEntityProviderSettings entityProviderSettings = new EntityProviderSettings
            {
                ClaimsProviderName = this.ClaimsProviderName,
                AlwaysResolveUserInput = this.AlwaysResolveUserInput,
                ClaimTypes = this.ClaimTypes,
                CustomData = this.CustomData,
                EnableAugmentation = this.EnableAugmentation,
                EntityDisplayTextPrefix = this.EntityDisplayTextPrefix,
                FilterExactMatchOnly = this.FilterExactMatchOnly,
                IdentityClaimTypeConfig = this.IdentityClaimTypeConfig,
                MainGroupClaimTypeConfig = this.MainGroupClaimTypeConfig,
                Name = this.Name,
                RuntimeClaimTypesList = this.RuntimeClaimTypesList,
                RuntimeMetadataConfig = this.RuntimeMetadataConfig,
                Timeout = this.Timeout,
                Version = this.Version,
            };
            return (TConfiguration)entityProviderSettings;
            //EntityProviderConfiguration<TConfiguration> copy = new EntityProviderConfiguration<TConfiguration>();
            //copy.ClaimsProviderName = this.ClaimsProviderName;
            //copy.RuntimeClaimTypesList = this.RuntimeClaimTypesList;
            //copy.RuntimeMetadataConfig = this.RuntimeMetadataConfig;
            //copy.IdentityClaimTypeConfig = this.IdentityClaimTypeConfig;
            //copy.MainGroupClaimTypeConfig = this.MainGroupClaimTypeConfig;
            ////copy = (EntityProviderConfiguration)Utils.CopyPersistedFields(typeof(EntityProviderConfiguration), this, copy);
            //copy.LocalConfiguration = (TConfiguration)Utils.CopyPersistedFields(typeof(TConfiguration), this, copy);
            ////copy.ClaimTypes = new ClaimTypeConfigCollection(this.ClaimTypes.SPTrust);
            ////foreach (ClaimTypeConfig currentObject in this.ClaimTypes)
            ////{
            ////    copy.ClaimTypes.Add(currentObject.CopyConfiguration(), false);
            ////}
            ////copy.InitializeRuntimeSettings();
            //return (TConfiguration)copy.LocalConfiguration;





            //copy.ClaimsProviderName = this.ClaimsProviderName;
            //copy.ClaimTypes = new ClaimTypeConfigCollection(this.ClaimTypes.SPTrust);
            //foreach (ClaimTypeConfig currentObject in this.ClaimTypes)
            //{
            //    copy.ClaimTypes.Add(currentObject.CopyConfiguration(), false);
            //}
            //copy.AlwaysResolveUserInput = this.AlwaysResolveUserInput;
            //copy.FilterExactMatchOnly = this.FilterExactMatchOnly;
            //copy.EnableAugmentation = this.EnableAugmentation;
            //copy.EntityDisplayTextPrefix = this.EntityDisplayTextPrefix;
            //copy.Timeout = this.Timeout;
            //copy.CustomData = this.CustomData;
            ////copy.MaxSearchResultsCount = this.MaxSearchResultsCount;            

            //copy.InitializeRuntimeSettings();
            //return copy;
        }

        public virtual void ApplyConfiguration(TConfiguration configuration)
        {
            this.ClaimsProviderName = configuration.ClaimsProviderName;
            foreach (ClaimTypeConfig claimTypeConfig in this.ClaimTypes)
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

        //IEntityProviderSettings IEntityProviderSettings.CopyConfiguration()
        //{
        //    throw new NotImplementedException();
        //}

        /// <summary>
        /// Returns the global configuration, stored as a persisted object in the SharePoint configuration database
        /// </summary>
        /// <param name="configurationName">The name of the configuration</param>
        /// <param name="initializeRuntimeSettings">Set to true to initialize the runtime settings</param>
        /// <returns></returns>
        public static EntityProviderConfig<TConfiguration> GetGlobalConfiguration(Guid configurationId, Type T, bool initializeRuntimeSettings = false)
        {
            SPFarm parent = SPFarm.Local;
            try
            {
                //IEntityProviderSettings configuration = (IEntityProviderSettings)parent.GetObject(configurationName, parent.Id, typeof(EntityProviderConfiguration));
                //Conf<TConfiguration> configuration = (Conf<TConfiguration>)parent.GetObject(configurationName, parent.Id, T);
                //Conf<TConfiguration> configuration = (Conf<TConfiguration>)parent.GetObject(configurationName, parent.Id, typeof(Conf<TConfiguration>));
                EntityProviderConfig<TConfiguration> configuration = (EntityProviderConfig<TConfiguration>)parent.GetObject(configurationId);
                if (configuration != null && initializeRuntimeSettings == true)
                {
                    configuration.InitializeRuntimeSettings();
                }
                return configuration;
            }
            catch (Exception ex)
            {
                Logger.LogException(String.Empty, $"while retrieving configuration ID '{configurationId}'", TraceCategory.Configuration, ex);
            }
            return null;
        }

        public static void DeleteGlobalConfiguration(Guid configurationId, Type T)
        {
            EntityProviderConfig<TConfiguration> configuration = (EntityProviderConfig<TConfiguration>)GetGlobalConfiguration(configurationId, T);
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
            EntityProviderConfig<TConfiguration> existingConfig = GetGlobalConfiguration(configurationID, T);
            if (existingConfig != null)
            {
                DeleteGlobalConfiguration(configurationID, T);
            }

            Logger.Log($"Creating configuration '{configurationName}' with Id {configurationID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);

            // Calling constructor as below is not possible and generate Compiler Error CS0304, so use reflection to call the desired constructor instead
            //TConfiguration config = new TConfiguration(persistedObjectName, SPFarm.Local, claimsProviderName);
            ConstructorInfo ctorWithParameters = T.GetConstructor(new[] { typeof(string), typeof(SPFarm), typeof(string) });
            EntityProviderConfig<TConfiguration> config = (EntityProviderConfig<TConfiguration>)ctorWithParameters.Invoke(new object[] { configurationName, SPFarm.Local, claimsProviderName });

            config.Id = configurationID;//new Guid(configurationID);
            // If parameter ensure is true, the call will not throw if the object already exists.
            config.Update(true);
            Logger.Log($"Created configuration '{configurationName}' with Id {config.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            return config;
        }
    }
}
