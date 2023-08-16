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

    public class AADEntityProviderSettings : EntityProviderSettings, IAADSettings
    {
        public List<AzureTenant> AzureTenants { get; set; }

        public string ProxyAddress { get; set; }

        public bool FilterSecurityEnabledGroupsOnly { get; set; }

        public AADEntityProviderSettings() : base() { }

        public AADEntityProviderSettings(List<ClaimTypeConfig> runtimeClaimTypesList, IEnumerable<ClaimTypeConfig> runtimeMetadataConfig, IdentityClaimTypeConfig identityClaimTypeConfig, ClaimTypeConfig mainGroupClaimTypeConfig)
            : base(runtimeClaimTypesList, runtimeMetadataConfig, identityClaimTypeConfig, mainGroupClaimTypeConfig)
        {
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
    }

    public class AADEntityProviderConfig<TConfiguration> : EntityProviderConfig<TConfiguration>
        where TConfiguration : IAADSettings
    {
        public List<AzureTenant> AzureTenants
        {
            get => _AzureTenants;
            set => _AzureTenants = value;
        }
        [Persisted]
        private List<AzureTenant> _AzureTenants = new List<AzureTenant>();

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

        public AADEntityProviderConfig() : base() { }
        public AADEntityProviderConfig(string configurationName, SPPersistedObject parent, string claimsProviderName) : base(configurationName, parent, claimsProviderName)
        {
        }

        public override bool InitializeDefaultSettings()
        {
            return base.InitializeDefaultSettings();
        }

        protected override bool InitializeInternalRuntimeSettings()
        {
            bool success = base.InitializeInternalRuntimeSettings();
            foreach (var tenant in this.AzureTenants)
            {
                tenant.InitializeAuthentication(this.Timeout, this.ProxyAddress);
            }

            return success;
        }

        protected override TConfiguration GenerateLocalSettings()
        {
            IAADSettings entityProviderSettings = new AADEntityProviderSettings(
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

                // Properties specific to type IAADSettings
                AzureTenants = this.AzureTenants,
                ProxyAddress = this.ProxyAddress,
                FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly,
            };
            return (TConfiguration)entityProviderSettings;

            //TSettings baseEntityProviderSettings = base.GenerateLocalSettings();
            //AADEntityProviderSettings entityProviderSettings = baseEntityProviderSettings as AADEntityProviderSettings;
            //entityProviderSettings.AzureTenants = this.AzureTenants;
            //entityProviderSettings.ProxyAddress = this.ProxyAddress;
            //entityProviderSettings.FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly;
            //return (TSettings)(IAADSettings)entityProviderSettings;
        }

        public override void ApplySettings(TConfiguration configuration, bool commitIfValid)
        {
            // Properties specific to type IAADSettings
            this.AzureTenants = configuration.AzureTenants;
            this.FilterSecurityEnabledGroupsOnly = configuration.FilterSecurityEnabledGroupsOnly;
            this.ProxyAddress = configuration.ProxyAddress;
            
            base.ApplySettings(configuration, commitIfValid);
        }

        /// <summary>
        /// Generate and return default configuration
        /// </summary>
        /// <returns></returns>
        public static AADEntityProviderConfig<TConfiguration> ReturnDefaultConfiguration(string claimsProviderName)
        {
            AADEntityProviderConfig<TConfiguration> defaultConfig = new AADEntityProviderConfig<TConfiguration>();
            defaultConfig.ClaimsProviderName = claimsProviderName;
            defaultConfig.ClaimTypes = AADEntityProviderSettings.ReturnDefaultClaimTypesConfig(claimsProviderName);
            return defaultConfig;
        }

        public override ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig()
        {
            return AADEntityProviderSettings.ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
        }

        public void ResetClaimTypesList()
        {
            ClaimTypes.Clear();
            ClaimTypes = AADEntityProviderSettings.ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
            Logger.Log($"Claim types list of configuration '{Name}' was successfully reset to default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }
    }
}
