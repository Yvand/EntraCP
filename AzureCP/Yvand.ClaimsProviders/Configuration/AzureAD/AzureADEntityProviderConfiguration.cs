using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;

namespace Yvand.ClaimsProviders.Configuration.AzureAD
{
    public class AzureADEntityProviderConfiguration : EntityProviderConfiguration
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

        public AzureADEntityProviderConfiguration() : base() { }
        public AzureADEntityProviderConfiguration(string configurationName, SPPersistedObject parent, string claimsProviderName) : base(configurationName, parent, claimsProviderName)
        {
        }

        public AzureADEntityProviderConfiguration(string claimsProviderName) : base(claimsProviderName)
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
        public override EntityProviderConfiguration CopyConfiguration()
        {
            // This is not possible to case an object to an inherited type from its base type: https://stackoverflow.com/questions/12565736/convert-base-class-to-derived-class
            //EntityProviderConfiguration baseCopy = base.CopyConfiguration();
            //AzureADEntityProviderConfiguration copy = (AzureADEntityProviderConfiguration)baseCopy;
            //AzureADEntityProviderConfiguration copy = new AzureADEntityProviderConfiguration(this.ClaimsProviderName);
            // Use default constructor to bypass initialization, which is useless since properties will be manually set here
            AzureADEntityProviderConfiguration copy = new AzureADEntityProviderConfiguration();
            copy.ClaimsProviderName = this.ClaimsProviderName;

            // Redo here the CopyConfiguration done in base class
            copy.ClaimTypes = new ClaimTypeConfigCollection(this.ClaimTypes.SPTrust);
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
            //copy.MaxSearchResultsCount = this.MaxSearchResultsCount;

            // Copy properties specific to type AzureADEntityProviderConfiguration
            copy.AzureTenants = this.AzureTenants;
            copy.FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly;
            copy.ProxyAddress = this.ProxyAddress;

            copy.InitializeRuntimeSettings();
            return copy;
        }

        /// <summary>
        /// Generate and return default configuration
        /// </summary>
        /// <returns></returns>
        public static AzureADEntityProviderConfiguration ReturnDefaultConfiguration(string claimsProviderName)
        {
            AzureADEntityProviderConfiguration defaultConfig = new AzureADEntityProviderConfiguration();
            defaultConfig.ClaimsProviderName = claimsProviderName;
            defaultConfig.AzureTenants = new List<AzureTenant>();
            defaultConfig.ClaimTypes = ReturnDefaultClaimTypesConfig(claimsProviderName);
            return defaultConfig;
        }

        public override ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig()
        {
            return AzureADEntityProviderConfiguration.ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
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
