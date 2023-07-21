using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Yvand.ClaimsProviders.ClaimsProviderLogging;

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

        public AzureADEntityProviderConfiguration(){}
        public AzureADEntityProviderConfiguration(string persistedObjectName, SPPersistedObject parent, string claimsProviderName) : base(persistedObjectName, parent, claimsProviderName)
        {
        }

        public AzureADEntityProviderConfiguration(string claimsProviderName) : base(claimsProviderName)
        {
        }

        public override bool InitializeRuntimeSettings()
        {
            bool success = base.InitializeRuntimeSettings();
            // Set properties AuthenticationProvider and GraphService
            foreach (var tenant in this.AzureTenants)
            {
                tenant.InitializeGraphForAppOnlyAuth(this.ClaimsProviderName, this.Timeout);
            }
            return success;
        }

        new public AzureADEntityProviderConfiguration CopyConfiguration()
        {
            EntityProviderConfiguration baseCopy = base.CopyConfiguration();
            AzureADEntityProviderConfiguration copy = (AzureADEntityProviderConfiguration)baseCopy;
            copy.AzureTenants = this.AzureTenants;
            copy.FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly;
            return copy;
        }

        /// <summary>
        /// Returns the configuration of AzureCP
        /// </summary>
        /// <returns></returns>
        public static AzureADEntityProviderConfiguration GetConfiguration()
        {
            return GetConfiguration(ClaimsProviderConstants.CONFIG_NAME);
        }

        /// <summary>
        /// Returns the configuration of AzureCP, but does not initialize the runtime settings
        /// </summary>
        /// <param name="persistedObjectName">Name of the configuration</param>
        /// <returns></returns>
        public static AzureADEntityProviderConfiguration GetConfiguration(string persistedObjectName)
        {
            SPPersistedObject parent = SPFarm.Local;
            try
            {
                AzureADEntityProviderConfiguration persistedObject = parent.GetChild<AzureADEntityProviderConfiguration>(persistedObjectName);
                if (persistedObject != null)
                {
                    //persistedObject.CheckAndCleanConfiguration(spTrustName);
                    return persistedObject;
                }
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(String.Empty, $"while retrieving configuration '{persistedObjectName}'", TraceCategory.Configuration, ex);
            }
            return null;
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

        /// <summary>
        /// Generate and return default claim types configuration list
        /// </summary>
        /// <returns></returns>
        public static ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig(string claimsProviderName)
        {
            if (String.IsNullOrWhiteSpace(claimsProviderName))
            {
                throw new ArgumentNullException("claimsProviderName cannot be null.");
            }

            SPTrustedLoginProvider spTrust = Utils.GetSPTrustAssociatedWithClaimsProvider(claimsProviderName);
            if (spTrust == null)
            {
                ClaimsProviderLogging.Log($"No SPTrustedLoginProvider associated with claims provider '{claimsProviderName}' was found.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                return null;
            }

            ClaimTypeConfigCollection newCTConfigCollection = new ClaimTypeConfigCollection()
            {
                // Identity claim type. "Name" (http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name) is a reserved claim type in SharePoint that cannot be used in the SPTrust.
                //new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.UserPrincipalName, ClaimType = WIF4_5.ClaimTypes.Upn},
                new IdentityClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.UserPrincipalName, ClaimType = spTrust.IdentityClaimTypeInformation.MappedClaimType},

                // Additional properties to find user and create entity with the identity claim type (UseMainClaimTypeOfDirectoryObject=true)
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.DisplayName, UseMainClaimTypeOfDirectoryObject = true, EntityDataKey = PeopleEditorEntityDataKeys.DisplayName},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.GivenName, UseMainClaimTypeOfDirectoryObject = true}, //Yvan
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.Surname, UseMainClaimTypeOfDirectoryObject = true},   //Duhamel
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.Mail, EntityDataKey = PeopleEditorEntityDataKeys.Email, UseMainClaimTypeOfDirectoryObject = true},

                // Additional properties to populate metadata of entity created: no claim type set, EntityDataKey is set and UseMainClaimTypeOfDirectoryObject = false (default value)
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.MobilePhone, EntityDataKey = PeopleEditorEntityDataKeys.MobilePhone},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.JobTitle, EntityDataKey = PeopleEditorEntityDataKeys.JobTitle},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.Department, EntityDataKey = PeopleEditorEntityDataKeys.Department},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation, EntityDataKey = PeopleEditorEntityDataKeys.Location},

                // Group
                new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, DirectoryObjectProperty = AzureADObjectProperty.Id, ClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType, DirectoryObjectPropertyToShowAsDisplayText = AzureADObjectProperty.DisplayName},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, DirectoryObjectProperty = AzureADObjectProperty.DisplayName, UseMainClaimTypeOfDirectoryObject = true, EntityDataKey = PeopleEditorEntityDataKeys.DisplayName},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, DirectoryObjectProperty = AzureADObjectProperty.Mail, EntityDataKey = PeopleEditorEntityDataKeys.Email},
            };
            newCTConfigCollection.SPTrust = spTrust;
            return newCTConfigCollection;
        }

        public void ResetClaimTypesList()
        {
            ClaimTypes.Clear();
            ClaimTypes = ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
            ClaimsProviderLogging.Log($"Claim types list of configuration '{Name}' was successfully reset to default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Create a persisted object with default configuration of AzureCP.
        /// </summary>
        /// <param name="persistedObjectID">GUID of the configuration, stored as a persisted object into SharePoint configuration database</param>
        /// <param name="persistedObjectName">Name of the configuration, stored as a persisted object into SharePoint configuration database</param>
        /// <param name="claimsProviderName">Name of the SPTrustedLoginProvider that claims provider is associated with</param>
        /// <returns></returns>
        public static AzureADEntityProviderConfiguration CreateConfiguration(string persistedObjectID, string persistedObjectName, string claimsProviderName)
        {
            if (String.IsNullOrEmpty(claimsProviderName))
            {
                throw new ArgumentNullException("spTrustName");
            }

            // Ensure it doesn't already exists and delete it if so
            AzureADEntityProviderConfiguration existingConfig = AzureADEntityProviderConfiguration.GetConfiguration(persistedObjectName);
            if (existingConfig != null)
            {
                DeleteConfiguration(persistedObjectName);
            }

            ClaimsProviderLogging.Log($"Creating configuration '{persistedObjectName}' with Id {persistedObjectID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);
            AzureADEntityProviderConfiguration config = new AzureADEntityProviderConfiguration(persistedObjectName, SPFarm.Local, claimsProviderName);
            //config.ResetCurrentConfiguration();
            config.Id = new Guid(persistedObjectID);
            config.Update();
            ClaimsProviderLogging.Log($"Created configuration '{persistedObjectName}' with Id {config.Id}", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
            return config;
        }

        /// <summary>
        /// Delete persisted object from configuration database
        /// </summary>
        /// <param name="persistedObjectName">Name of persisted object to delete</param>
        public static void DeleteConfiguration(string persistedObjectName)
        {
            AzureADEntityProviderConfiguration config = AzureADEntityProviderConfiguration.GetConfiguration(persistedObjectName);
            if (config == null)
            {
                ClaimsProviderLogging.Log($"Configuration '{persistedObjectName}' was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
                return;
            }
            config.Delete();
            ClaimsProviderLogging.Log($"Configuration '{persistedObjectName}' was successfully deleted from configuration database", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }
    }
}
