using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;

namespace Yvand.ClaimsProviders.Config
{
    public interface IAADSettings : IEntityProviderSettings
    {
        List<AzureTenant> AzureTenants { get; }
        string ProxyAddress { get; }
        bool FilterSecurityEnabledGroupsOnly { get; }
    }

    public class AADEntityProviderSettings : IAADSettings
    {
        public List<AzureTenant> AzureTenants  { get; set; }

        public string ProxyAddress  { get; set; }

        public bool FilterSecurityEnabledGroupsOnly  { get; set; }

        public long Version  { get; set; }

        public string Name  { get; set; }

        public string ClaimsProviderName  { get; set; }

        public ClaimTypeConfigCollection ClaimTypes  { get; set; }

        public bool AlwaysResolveUserInput  { get; set; }

        public bool FilterExactMatchOnly  { get; set; }

        public bool EnableAugmentation  { get; set; }

        public string EntityDisplayTextPrefix  { get; set; }

        public int Timeout  { get; set; }

        public string CustomData  { get; set; }

        public List<ClaimTypeConfig> RuntimeClaimTypesList  { get; set; }

        public IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig  { get; set; }

        public IdentityClaimTypeConfig IdentityClaimTypeConfig  { get; set; }

        public ClaimTypeConfig MainGroupClaimTypeConfig  { get; set; }
    }

    public class AADConf<TConfiguration> : Conf<TConfiguration>
        where TConfiguration : IAADSettings
    {
        public List<AzureTenant> AzureTenants
        {
            get => _AzureTenants;
            set => _AzureTenants = value;
        }
        [Persisted]
        private List<AzureTenant> _AzureTenants;

        public string ProxyAddress
        {
            get => _ProxyAddress;
            set => _ProxyAddress = value;
        }
        [Persisted]
        private string _ProxyAddress;

        /// <summary>
        /// Set if only AAD groups with securityEnabled = true should be returned - https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0
        /// </summary>
        public bool FilterSecurityEnabledGroupsOnly
        {
            get => _FilterSecurityEnabledGroupsOnly;
            set => _FilterSecurityEnabledGroupsOnly = value;
        }
        [Persisted]
        private bool _FilterSecurityEnabledGroupsOnly = false;

        public AADConf() : base() { }
        public AADConf(string configurationName, SPPersistedObject parent, string claimsProviderName) : base(configurationName, parent, claimsProviderName)
        {
        }

        public AADConf(string claimsProviderName) : base(claimsProviderName)
        {
        }

        protected override bool InitializeDefaultSettings()
        {
            this.AzureTenants = new List<AzureTenant>();
            return base.InitializeDefaultSettings();
        }

        public override bool InitializeRuntimeSettings()
        {
            bool success = base.InitializeRuntimeSettings();
            foreach (var tenant in this.AzureTenants)
            {
                tenant.InitializeAuthentication(this.Timeout, this.ProxyAddress);
            }
            return success;
        }

        //public new AzureADEntityProviderConfiguration CopyConfiguration()
        public override TConfiguration CopyConfiguration()
        {
            //IAzureADEntityProviderSettings settings = base.CopyConfiguration();
            IAADSettings entityProviderSettings = new AADEntityProviderSettings
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

                AzureTenants = this.AzureTenants,
                ProxyAddress = this.ProxyAddress,
                FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly,
            };
            return (TConfiguration)entityProviderSettings;
            //// This is not possible to case an object to an inherited type from its base type: https://stackoverflow.com/questions/12565736/convert-base-class-to-derived-class
            ////EntityProviderConfiguration baseCopy = base.CopyConfiguration();
            ////AzureADEntityProviderConfiguration copy = (AzureADEntityProviderConfiguration)baseCopy;
            ////AzureADEntityProviderConfiguration copy = new AzureADEntityProviderConfiguration(this.ClaimsProviderName);
            //// Use default constructor to bypass initialization, which is useless since properties will be manually set here
            //AzureADEntityProviderConfiguration<TConfiguration> copy = new AzureADEntityProviderConfiguration<TConfiguration>();
            //copy.ClaimsProviderName = this.ClaimsProviderName;
            //copy = (AzureADEntityProviderConfiguration<TConfiguration>)Utils.CopyPersistedFields(typeof(EntityProviderConfiguration<TConfiguration>), this, copy);
            //copy = (AzureADEntityProviderConfiguration<TConfiguration>)Utils.CopyPersistedFields(typeof(AzureADEntityProviderConfiguration<TConfiguration>), this, copy);
            ////copy.ClaimTypes = new ClaimTypeConfigCollection(this.ClaimTypes.SPTrust);
            ////foreach (ClaimTypeConfig currentObject in this.ClaimTypes)
            ////{
            ////    copy.ClaimTypes.Add(currentObject.CopyConfiguration(), false);
            ////}
            ////copy.AzureTenants = new List<AzureTenant>();
            ////foreach(AzureTenant tenant in this.AzureTenants)
            ////{
            ////    copy.AzureTenants.Add(tenant.CopyConfiguration());
            ////}
            //copy.InitializeRuntimeSettings();
            //return copy.LocalConfiguration;
        }

        public void ApplyConfiguration(AADConf<TConfiguration> configuration)
        {
            // This is not possible to case an object to an inherited type from its base type: https://stackoverflow.com/questions/12565736/convert-base-class-to-derived-class

            // Redo here the ApplyConfiguration done in base class
            this.ClaimsProviderName = configuration.ClaimsProviderName;
            this.ClaimTypes = new ClaimTypeConfigCollection(configuration.ClaimTypes.SPTrust);
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

            // Copy properties specific to type AzureADEntityProviderConfiguration
            this.AzureTenants = configuration.AzureTenants;
            this.FilterSecurityEnabledGroupsOnly = configuration.FilterSecurityEnabledGroupsOnly;
            this.ProxyAddress = configuration.ProxyAddress;
        }

        /// <summary>
        /// Generate and return default configuration
        /// </summary>
        /// <returns></returns>
        public static AADConf<TConfiguration> ReturnDefaultConfiguration(string claimsProviderName)
        {
            AADConf<TConfiguration> defaultConfig = new AADConf<TConfiguration>();
            defaultConfig.ClaimsProviderName = claimsProviderName;
            defaultConfig.AzureTenants = new List<AzureTenant>();
            defaultConfig.ClaimTypes = ReturnDefaultClaimTypesConfig(claimsProviderName);
            return defaultConfig;
        }

        public override ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig()
        {
            return AADConf<TConfiguration>.ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
        }

        /// <summary>
        /// Generate and return default claim types configuration list
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

        public void ResetClaimTypesList()
        {
            ClaimTypes.Clear();
            ClaimTypes = ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
            Logger.Log($"Claim types list of configuration '{Name}' was successfully reset to default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }
    }
}
