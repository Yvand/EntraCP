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
        ClaimTypeConfigCollection ClaimTypes { get; }
        bool AlwaysResolveUserInput { get; }
        bool FilterExactMatchOnly { get; }
        bool EnableAugmentation { get; }
        string EntityDisplayTextPrefix { get; }
        int Timeout { get; }
        string CustomData { get; }

        // Copy of the internal runtime settings, which external classes can only access through an object implementing this interface
        List<ClaimTypeConfig> RuntimeClaimTypesList { get; }
        IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; }
        IdentityClaimTypeConfig IdentityClaimTypeConfig { get; }
        ClaimTypeConfig MainGroupClaimTypeConfig { get; }
    }

    public class EntityProviderSettings : IEntityProviderSettings
    {
        public long Version { get; set; }

        public string Name { get; set; }

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

    public class EntityProviderConfig<TSettings> : SPPersistedObject
        where TSettings : IEntityProviderSettings
    {
        /// <summary>
        /// Gets the local settings, based on the global settings stored in a persisted object
        /// </summary>
        public TSettings LocalSettings { get; private set; }

        /// <summary>
        /// Gets the current version of the local settings
        /// </summary>
        protected long LocalSettingsVersion { get; private set; } = 0;

        #region "Public runtime settings"
        /// <summary>
        /// gets or sets the claim types and their mapping with a DirectoryObject property
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

        /// <summary>
        /// Gets or sets whether to skip Azure AD lookup and consider any input as valid.
        /// This can be useful to keep people picker working even if connectivity with the Azure tenant is lost.
        /// </summary>
        public bool AlwaysResolveUserInput
        {
            get => _AlwaysResolveUserInput;
            set => _AlwaysResolveUserInput = value;
        }
        [Persisted]
        private bool _AlwaysResolveUserInput;

        /// <summary>
        /// Gets or sets whether to return only results that match exactly the user input (case-insensitive).
        /// </summary>
        public bool FilterExactMatchOnly
        {
            get => _FilterExactMatchOnly;
            set => _FilterExactMatchOnly = value;
        }
        [Persisted]
        private bool _FilterExactMatchOnly;

        /// <summary>
        /// Gets or sets whether to return the Azure AD groups that the user is a member of.
        /// </summary>
        public bool EnableAugmentation
        {
            get => _EnableAugmentation;
            set => _EnableAugmentation = value;
        }
        [Persisted]
        private bool _EnableAugmentation = true;

        /// <summary>
        /// Gets or sets a string that will appear as a prefix of the text of each result, in the people picker.
        /// </summary>
        public string EntityDisplayTextPrefix
        {
            get => _EntityDisplayTextPrefix;
            set => _EntityDisplayTextPrefix = value;
        }
        [Persisted]
        private string _EntityDisplayTextPrefix;

        /// <summary>
        /// Gets or sets the timeout before giving up the query to Azure AD.
        /// </summary>
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

        #region "Public runtime properties"
        private SPTrustedLoginProvider _SPTrust;
        /// <summary>
        /// Gets the SharePoint trust that has its property ClaimProviderName equals to <see cref="ClaimsProviderName"/>
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
        }

        public virtual bool InitializeDefaultSettings()
        {
            this.ClaimTypes = ReturnDefaultClaimTypesConfig();
            return true;
        }

        /// <summary>
        /// Sets the internal runtime settings properties
        /// </summary>
        /// <returns>True if successful, false if not</returns>
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
        /// Ensures that property <see cref="LocalSettings"/> is valid and up to date
        /// </summary>
        /// <returns>The property <see cref="LocalSettings"/> if is valid, null otherwise</returns>
        public TSettings RefreshLocalSettingsIfNeeded()
        {
            Guid configurationId = this.Id;
            EntityProviderConfig<TSettings> globalConfiguration = GetGlobalConfiguration(configurationId);

            if (globalConfiguration == null)
            {
                Logger.Log($"[{ClaimsProviderName}] Cannot continue because configuration '{configurationId}' was not found in configuration database, visit AzureCP admin pages in central administration to create it.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                this.LocalSettings = default(TSettings);
                return default(TSettings);
            }

            if (this.LocalSettingsVersion == globalConfiguration.Version)
            {
                Logger.Log($"[{ClaimsProviderName}] Configuration '{configurationId}' is up to date with version {this.LocalSettingsVersion}.",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Core);
                return this.LocalSettings;
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
            this.LocalSettings = (TSettings)globalConfiguration.GenerateLocalSettings();
#if !DEBUGx
            this.LocalSettingsVersion = globalConfiguration.Version;
#endif

            if (this.LocalSettings.ClaimTypes == null || this.LocalSettings.ClaimTypes.Count == 0)
            {
                Logger.Log($"[{ClaimsProviderName}] Configuration '{this.LocalSettings.Name}' was found but collection ClaimTypes is empty. Visit AzureCP admin pages in central administration to create it.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
            }
            return this.LocalSettings;
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
        /// Ensures that the current settings is valid and can be safely saved and used
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

        // This method fires 3 times in a raw just when the configurationis updated, and anyway it bypassws the logic to update only if needed (and safely in regards to thread safety)
        //protected override void OnDeserialization()
        //{
        //    base.OnDeserialization();
        //    this.InitializeInternalRuntimeSettings();
        //}

        /// <summary>
        /// Returns a read-only settings, copied from the current settings.
        /// </summary>
        /// <returns></returns>
        protected virtual TSettings GenerateLocalSettings()
        {
            IEntityProviderSettings entityProviderSettings = new EntityProviderSettings(
                this.RuntimeClaimTypesList,
                this.RuntimeMetadataConfig,
                this.IdentityClaimTypeConfig,
                this.MainGroupClaimTypeConfig)
            {
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
            return (TSettings)entityProviderSettings;
        }

        /// <summary>
        /// Applies the settings passed in parameter to the current settings
        /// </summary>
        /// <param name="settings"></param>
        public virtual void ApplySettings(TSettings settings, bool commitIfValid)
        {
            this.ClaimTypes = new ClaimTypeConfigCollection(this.SPTrust);
            foreach (ClaimTypeConfig claimTypeConfig in settings.ClaimTypes)
            {
                this.ClaimTypes.Add(claimTypeConfig.CopyConfiguration(), false);
            }
            this.AlwaysResolveUserInput = settings.AlwaysResolveUserInput;
            this.FilterExactMatchOnly = settings.FilterExactMatchOnly;
            this.EnableAugmentation = settings.EnableAugmentation;
            this.EntityDisplayTextPrefix = settings.EntityDisplayTextPrefix;
            this.Timeout = settings.Timeout;
            this.CustomData = settings.CustomData;

            if(commitIfValid)
            {
                this.Update();
            }
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
        /// <param name="configurationId">The ID of the configuration</param>
        /// <param name="initializeLocalSettings">Set to true to initialize the property <see cref="LocalSettings"/></param>
        /// <returns></returns>
        public static EntityProviderConfig<TSettings> GetGlobalConfiguration(Guid configurationId, bool initializeLocalSettings = false)
        {
            SPFarm parent = SPFarm.Local;
            try
            {
                //IEntityProviderSettings settings = (IEntityProviderSettings)parent.GetObject(configurationName, parent.Id, typeof(EntityProviderConfiguration));
                //Conf<TSettings> settings = (Conf<TSettings>)parent.GetObject(configurationName, parent.Id, T);
                //Conf<TSettings> settings = (Conf<TSettings>)parent.GetObject(configurationName, parent.Id, typeof(Conf<TSettings>));
                EntityProviderConfig<TSettings> configuration = (EntityProviderConfig<TSettings>)parent.GetObject(configurationId);
                if (configuration != null && initializeLocalSettings == true)
                {
                    configuration.RefreshLocalSettingsIfNeeded();
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
            EntityProviderConfig<TSettings> configuration = (EntityProviderConfig<TSettings>)GetGlobalConfiguration(configurationId);
            if (configuration == null)
            {
                Logger.Log($"Configuration ID '{configurationId}' was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
                return;
            }
            configuration.Delete();
            Logger.Log($"Configuration ID '{configurationId}' was successfully deleted from configuration database", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        public static EntityProviderConfig<TSettings> CreateGlobalConfiguration(Guid configurationID, string configurationName, string claimsProviderName, Type T)
        {
            if (String.IsNullOrWhiteSpace(claimsProviderName))
            {
                throw new ArgumentNullException(nameof(claimsProviderName));
            }

            // Ensure it doesn't already exists and delete it if so
            EntityProviderConfig<TSettings> existingConfig = GetGlobalConfiguration(configurationID);
            if (existingConfig != null)
            {
                DeleteGlobalConfiguration(configurationID);
            }

            Logger.Log($"Creating configuration '{configurationName}' with Id {configurationID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);

            ConstructorInfo ctorWithParameters = T.GetConstructor(new[] { typeof(string), typeof(SPFarm), typeof(string) });
            EntityProviderConfig<TSettings> config = (EntityProviderConfig<TSettings>)ctorWithParameters.Invoke(new object[] { configurationName, SPFarm.Local, claimsProviderName });

            config.Id = configurationID;
            // If parameter ensure is true, the call will not throw if the object already exists.
            config.Update(true);
            Logger.Log($"Created configuration '{configurationName}' with Id {config.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            return config;
        }
    }
}
