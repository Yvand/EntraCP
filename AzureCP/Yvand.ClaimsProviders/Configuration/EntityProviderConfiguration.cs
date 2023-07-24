﻿using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Yvand.ClaimsProviders.ClaimsProviderLogging;
using Yvand.ClaimsProviders.Configuration.AzureAD;

namespace Yvand.ClaimsProviders.Configuration
{
    /// <summary>
    /// Define base settings that are persisted in a persisted object
    /// </summary>
    public interface IPersistedEntityProviderSettings
    {
        ClaimTypeConfigCollection ClaimTypes { get; set; }
        bool AlwaysResolveUserInput { get; set; }
        bool FilterExactMatchOnly { get; set; }
        bool EnableAugmentation { get; set; }
        string EntityDisplayTextPrefix { get; set; }
        int Timeout { get; set; }
        string CustomData { get; set; }
        int MaxSearchResultsCount { get; set; }
    }

    ///// <summary>
    ///// Define base settings that are not persisted
    ///// </summary>
    //public interface IRuntimeEntityProviderSettings
    //{
    //    List<ClaimTypeConfig> ProcessedClaimTypesList { get; set; }
    //    IEnumerable<ClaimTypeConfig> MetadataConfig { get; set; }
    //    /// <summary>
    //    /// ClaimTypeConfig mapped to the identity claim in the SPTrustedIdentityTokenIssuer
    //    /// </summary>
    //    IdentityClaimTypeConfig IdentityClaimTypeConfig { get; set; }

    //    /// <summary>
    //    /// Group ClaimTypeConfig used to set the claim type for other group ClaimTypeConfig that have UseMainClaimTypeOfDirectoryObject set to true
    //    /// </summary>
    //    ClaimTypeConfig MainGroupClaimTypeConfig { get; set; }
    //}

    public class EntityProviderConfiguration : SPPersistedObject, IPersistedEntityProviderSettings//, IRuntimeEntityProviderSettings
    {
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

        /// <summary>
        /// Limit number of results returned to SharePoint during a search
        /// </summary>
        public int MaxSearchResultsCount
        {
            get => _MaxSearchResultsCount;
            set => _MaxSearchResultsCount = value;
        }
        [Persisted]
        private int _MaxSearchResultsCount = 30; // SharePoint sets maxCount to 30 in method FillSearch

        protected bool IsInitialized = false;

        // Runtime settings
        internal List<ClaimTypeConfig> RuntimeClaimTypesList { get; private set; }
        internal IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; private set; }
        internal IdentityClaimTypeConfig IdentityClaimTypeConfig { get; private set; }
        internal ClaimTypeConfig MainGroupClaimTypeConfig { get; private set; }

        public EntityProviderConfiguration() { }
        public EntityProviderConfiguration(string persistedObjectName, SPPersistedObject parent, string claimsProviderName) : base(persistedObjectName, parent)
        {
            this.ClaimsProviderName = claimsProviderName;
            //this.InitializeDefaultSettings();
            //this.InitializeRuntimeSettings();
            this.Initialize();
        }

        public EntityProviderConfiguration(string claimsProviderName)
        {
            this.ClaimsProviderName = claimsProviderName;
            //this.InitializeDefaultSettings();
            //this.InitializeRuntimeSettings();
            this.Initialize();
        }

        public virtual bool Initialize()
        {
            if(!this.IsInitialized)
            {
                this.InitializeDefaultSettings();
                this.InitializeRuntimeSettings();
                this.IsInitialized = true;
            }
            return this.IsInitialized;
        }

        protected virtual bool InitializeDefaultSettings()
        {
            this.ClaimTypes = ReturnDefaultClaimTypesConfig();
            return true;
        }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        protected virtual bool InitializeRuntimeSettings()
        {
            if (this.ClaimTypes?.Count <= 0)
            {
                ClaimsProviderLogging.Log($"[{this.ClaimsProviderName}] Cannot continue because configuration '{this.Name}' has 0 claim configured.",
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
                    x.DirectoryObjectProperty != AzureADObjectProperty.NotSet);

                if (claimTypeConfig == null)
                {
                    continue;
                }
                claimTypeConfig.ClaimTypeDisplayName = claimTypeInformation.DisplayName;
                claimTypesSetInTrust.Add(claimTypeConfig);
                if (String.Equals(this.SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    // Identity claim type found, set IdentityClaimTypeConfig property
                    identityClaimTypeFound = true;
                    this.IdentityClaimTypeConfig = IdentityClaimTypeConfig.ConvertClaimTypeConfig(claimTypeConfig);
                }
                else if (!groupClaimTypeFound && claimTypeConfig.EntityType == DirectoryObjectType.Group)
                {
                    groupClaimTypeFound = true;
                    this.MainGroupClaimTypeConfig = claimTypeConfig;
                }
            }

            if (!identityClaimTypeFound)
            {
                ClaimsProviderLogging.Log($"[{this.ClaimsProviderName}] Cannot continue because identity claim type '{this.SPTrust.IdentityClaimTypeInformation.MappedClaimType}' set in the SPTrustedIdentityTokenIssuer '{SPTrust.Name}' is missing in the ClaimTypeConfig list.", TraceSeverity.Unexpected, EventSeverity.ErrorCritical, TraceCategory.Core);
                return false;
            }

            // Check if there are additional properties to use in queries (UseMainClaimTypeOfDirectoryObject set to true)
            List<ClaimTypeConfig> additionalClaimTypeConfigList = new List<ClaimTypeConfig>();
            foreach (ClaimTypeConfig claimTypeConfig in this.ClaimTypes.Where(x => x.UseMainClaimTypeOfDirectoryObject))
            {
                if (claimTypeConfig.EntityType == DirectoryObjectType.User)
                {
                    claimTypeConfig.ClaimType = this.IdentityClaimTypeConfig.ClaimType;
                    claimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText = this.IdentityClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText;
                }
                else
                {
                    // If not a user, it must be a group
                    if (this.MainGroupClaimTypeConfig == null)
                    {
                        continue;
                    }
                    claimTypeConfig.ClaimType = this.MainGroupClaimTypeConfig.ClaimType;
                    claimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText = this.MainGroupClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText;
                    claimTypeConfig.ClaimTypeDisplayName = this.MainGroupClaimTypeConfig.ClaimTypeDisplayName;
                }
                additionalClaimTypeConfigList.Add(claimTypeConfig);
            }

            this.RuntimeClaimTypesList = new List<ClaimTypeConfig>(claimTypesSetInTrust.Count + additionalClaimTypeConfigList.Count);
            this.RuntimeClaimTypesList.AddRange(claimTypesSetInTrust);
            this.RuntimeClaimTypesList.AddRange(additionalClaimTypeConfigList);

            // Get all PickerEntity metadata with a DirectoryObjectProperty set
            this.RuntimeMetadataConfig = this.ClaimTypes.Where(x =>
                !String.IsNullOrEmpty(x.EntityDataKey) &&
                x.DirectoryObjectProperty != AzureADObjectProperty.NotSet);

            return true;
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
            this.InitializeRuntimeSettings();
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
        public virtual EntityProviderConfiguration CopyConfiguration()
        {
            // Cannot use reflection here to copy object because of the calls to methods CopyConfiguration() on some properties
            EntityProviderConfiguration copy = new EntityProviderConfiguration();
            copy.ClaimsProviderName = this.ClaimsProviderName;
            copy.ClaimTypes = new ClaimTypeConfigCollection();
            copy.ClaimTypes.SPTrust = this.ClaimTypes.SPTrust;
            foreach (ClaimTypeConfig currentObject in this.ClaimTypes)
            {
                copy.ClaimTypes.Add(currentObject.CopyConfiguration(), false);
            }
            copy.AlwaysResolveUserInput = this.AlwaysResolveUserInput;
            copy.FilterExactMatchOnly = this.FilterExactMatchOnly;
            copy.EnableAugmentation = this.EnableAugmentation;
            copy.EntityDisplayTextPrefix = this.EntityDisplayTextPrefix;
            copy.Timeout = this.Timeout;
            copy.CustomData = this.CustomData;
            copy.MaxSearchResultsCount = this.MaxSearchResultsCount;
            return copy;
        }

        public virtual void ResetCurrentConfiguration()
        {

        }

        public virtual ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig()
        {
            throw new NotImplementedException();
        }
    }
}
