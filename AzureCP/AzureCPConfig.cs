using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using Microsoft.Web.Hosting.Administration;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Web;
using static azurecp.ClaimsProviderLogging;
using WIF4_5 = System.Security.Claims;

namespace azurecp
{
    public enum AzureCloudInstance
    {
        //
        // Summary:
        //     Value communicating that the AzureCloudInstance is not specified.
        None,
        //
        // Summary:
        //     Microsoft Azure public cloud. Maps to https://login.microsoftonline.com
        AzurePublic,
        //
        // Summary:
        //     Microsoft Chinese national cloud. Maps to https://login.chinacloudapi.cn
        AzureChina,
        //
        // Summary:
        //     Microsoft German national cloud ("Black Forest"). Maps to https://login.microsoftonline.de
        AzureGermany,
        //
        // Summary:
        //     US Government cloud. Maps to https://login.microsoftonline.us
        AzureUsGovernment
    }

    public interface IAzureCPConfiguration
    {
        List<AzureTenant> AzureTenants { get; set; }
        ClaimTypeConfigCollection ClaimTypes { get; set; }
        bool AlwaysResolveUserInput { get; set; }
        bool FilterExactMatchOnly { get; set; }
        bool EnableAugmentation { get; set; }
        string EntityDisplayTextPrefix { get; set; }
        bool EnableRetry { get; set; }
        int Timeout { get; set; }
        string CustomData { get; set; }
        int MaxSearchResultsCount { get; set; }
        bool FilterSecurityEnabledGroupsOnly { get; set; }
    }

    public static class ClaimsProviderConstants
    {
        public static string CONFIG_ID => "0E9F8FB6-B314-4CCC-866D-DEC0BE76C237";
        public static string CONFIG_NAME => "AzureCPConfig";
        public static string GraphServiceEndpointVersion => "v1.0";
        //public static string DefaultGraphServiceEndpoint => "https://graph.microsoft.com/";
        //public static string DefaultLoginServiceEndpoint => "https://login.microsoftonline.com/";
        /// <summary>
        /// List of Microsoft Graph service root endpoints based on National Cloud as described on https://docs.microsoft.com/en-us/graph/deployments
        /// </summary>
        public static List<KeyValuePair<AzureCloudInstance, Uri>> AzureCloudEndpoints = new List<KeyValuePair<AzureCloudInstance, Uri>>()
        {
            new KeyValuePair<AzureCloudInstance, Uri>(AzureCloudInstance.AzurePublic, new Uri("https://login.microsoftonline.com")),
            new KeyValuePair<AzureCloudInstance, Uri>(AzureCloudInstance.AzureChina, new Uri("https://login.chinacloudapi.cn")),
            new KeyValuePair<AzureCloudInstance, Uri>(AzureCloudInstance.AzureGermany, new Uri("https://login.microsoftonline.de")),
            new KeyValuePair<AzureCloudInstance, Uri>(AzureCloudInstance.AzureUsGovernment, new Uri("https://login.microsoftonline.us")),
            new KeyValuePair<AzureCloudInstance, Uri>(AzureCloudInstance.None, new Uri("https://login.microsoftonline.com")),
        };
        public static string GroupClaimEntityType { get; set; } = SPClaimEntityTypes.FormsRole;
        public static bool EnforceOnly1ClaimTypeForGroup => true;     // In AzureCP, only 1 claim type can be used to create group permissions
        public static string DefaultMainGroupClaimType => WIF4_5.ClaimTypes.Role;
        public static string PUBLICSITEURL => "https://azurecp.yvand.net/";
        public static string GUEST_USERTYPE => "Guest";
        public static string MEMBER_USERTYPE => "Member";
        private static object Sync_SetClaimsProviderVersion = new object();
        private static string _ClaimsProviderVersion;
        public static readonly string ClientCertificatePrivateKeyPassword = "YVANDwRrEHVHQ57ge?uda";
        public static string ClaimsProviderVersion
        {
            get
            {
                if (!String.IsNullOrEmpty(_ClaimsProviderVersion))
                {
                    return _ClaimsProviderVersion;
                }

                // Method FileVersionInfo.GetVersionInfo() may hang and block all LDAPCP threads, so it is read only 1 time
                lock (Sync_SetClaimsProviderVersion)
                {
                    if (!String.IsNullOrEmpty(_ClaimsProviderVersion))
                    {
                        return _ClaimsProviderVersion;
                    }

                    try
                    {
                        _ClaimsProviderVersion = FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(AzureCP)).Location).FileVersion;
                    }
                    // If assembly was removed from the GAC, CLR throws a FileNotFoundException
                    catch (System.IO.FileNotFoundException)
                    {
                        // Current process will never detect if assembly is added to the GAC later, which is fine
                        _ClaimsProviderVersion = " ";
                    }
                    return _ClaimsProviderVersion;
                }
            }
        }

#if DEBUG
        public static int DEFAULT_TIMEOUT => 10000;
#else
        public static int DEFAULT_TIMEOUT => 4000;    // 4 secs
#endif
    }

    public class AzureCPConfig : SPPersistedObject, IAzureCPConfiguration
    {
        public List<AzureTenant> AzureTenants
        {
            get => AzureTenantsPersisted;
            set => AzureTenantsPersisted = value;
        }
        [Persisted]
        private List<AzureTenant> AzureTenantsPersisted;

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
            get => AlwaysResolveUserInputPersisted;
            set => AlwaysResolveUserInputPersisted = value;
        }
        [Persisted]
        private bool AlwaysResolveUserInputPersisted;

        public bool FilterExactMatchOnly
        {
            get => FilterExactMatchOnlyPersisted;
            set => FilterExactMatchOnlyPersisted = value;
        }
        [Persisted]
        private bool FilterExactMatchOnlyPersisted;

        public bool EnableAugmentation
        {
            get => AugmentAADRolesPersisted;
            set => AugmentAADRolesPersisted = value;
        }
        [Persisted]
        private bool AugmentAADRolesPersisted = true;

        public string EntityDisplayTextPrefix
        {
            get => _EntityDisplayTextPrefix;
            set => _EntityDisplayTextPrefix = value;
        }
        [Persisted]
        private string _EntityDisplayTextPrefix;

        public bool EnableRetry
        {
            get => _EnableRetry;
            set => _EnableRetry = value;
        }
        [Persisted]
        private bool _EnableRetry = false;

        public int Timeout
        {
            get => _Timeout;
            set => _Timeout = value;
        }
        [Persisted]
        private int _Timeout = ClaimsProviderConstants.DEFAULT_TIMEOUT;

        /// <summary>
        /// Name of the SPTrustedLoginProvider where AzureCP is enabled
        /// </summary>
        [Persisted]
        public string SPTrustName;

        private SPTrustedLoginProvider _SPTrust;
        private SPTrustedLoginProvider SPTrust
        {
            get
            {
                if (_SPTrust == null)
                {
                    _SPTrust = SPSecurityTokenServiceManager.Local.TrustedLoginProviders.GetProviderByName(SPTrustName);
                }
                return _SPTrust;
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

        public AzureCPConfig(string persistedObjectName, SPPersistedObject parent, string spTrustName) : base(persistedObjectName, parent)
        {
            this.SPTrustName = spTrustName;
        }

        public AzureCPConfig() { }

        /// <summary>
        /// Override this method to allow more users to update the object. True specifies that more users can update the object; otherwise, false. The default value is false.
        /// </summary>
        /// <returns></returns>
        protected override bool HasAdditionalUpdateAccess()
        {
            return false;
        }

        /// <summary>
        /// Returns the configuration of AzureCP
        /// </summary>
        /// <returns></returns>
        public static AzureCPConfig GetConfiguration()
        {
            return GetConfiguration(ClaimsProviderConstants.CONFIG_NAME, String.Empty);
        }

        /// <summary>
        /// Returns the configuration of AzureCP
        /// </summary>
        /// <param name="persistedObjectName"></param>
        /// <returns></returns>
        public static AzureCPConfig GetConfiguration(string persistedObjectName)
        {
            return GetConfiguration(persistedObjectName, String.Empty);
        }

        /// <summary>
        /// Returns the configuration of AzureCP
        /// </summary>
        /// <param name="persistedObjectName">Name of the configuration</param>
        /// <param name="spTrustName">Name of the SPTrustedLoginProvider using the claims provider</param>
        /// <returns></returns>
        public static AzureCPConfig GetConfiguration(string persistedObjectName, string spTrustName)
        {
            SPPersistedObject parent = SPFarm.Local;
            try
            {
                AzureCPConfig persistedObject = parent.GetChild<AzureCPConfig>(persistedObjectName);
                if (persistedObject != null)
                {
                    persistedObject.CheckAndCleanConfiguration(spTrustName);
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
        /// Commit changes to configuration database
        /// </summary>
        public override void Update()
        {
            // In case ClaimTypes collection was modified, test if it is still valid before committed changes to database
            try
            {
                ClaimTypeConfigCollection testUpdateCollection = new ClaimTypeConfigCollection();
                testUpdateCollection.SPTrust = this.SPTrust;
                foreach (ClaimTypeConfig curCTConfig in this.ClaimTypes)
                {
                    testUpdateCollection.Add(curCTConfig, false);
                }
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException("Some changes made to list ClaimTypes are invalid and cannot be committed to configuration database. Inspect inner exception for more details about the error.", ex);
            }

            base.Update();
            ClaimsProviderLogging.Log($"Configuration '{base.DisplayName}' was updated successfully to version {base.Version} in configuration database.",
                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
        }

        public static AzureCPConfig ResetConfiguration(string persistedObjectName)
        {
            AzureCPConfig previousConfig = GetConfiguration(persistedObjectName, String.Empty);
            if (previousConfig == null) { return null; }
            Guid configId = previousConfig.Id;
            string spTrustName = previousConfig.SPTrustName;
            DeleteConfiguration(persistedObjectName);
            AzureCPConfig newConfig = CreateConfiguration(configId.ToString(), persistedObjectName, spTrustName);
            ClaimsProviderLogging.Log($"Configuration '{persistedObjectName}' was successfully reset to its default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            return newConfig;
        }

        /// <summary>
        /// Set properties of current configuration to their default values
        /// </summary>
        /// <returns></returns>
        public void ResetCurrentConfiguration()
        {
            AzureCPConfig defaultConfig = ReturnDefaultConfiguration(this.SPTrustName) as AzureCPConfig;
            ApplyConfiguration(defaultConfig);
            CheckAndCleanConfiguration(String.Empty);
        }

        /// <summary>
        /// Apply configuration in parameter to current object. It does not copy SharePoint base class members
        /// </summary>
        /// <param name="configToApply"></param>
        public void ApplyConfiguration(AzureCPConfig configToApply)
        {
            // Copy non-inherited public properties
            PropertyInfo[] propertiesToCopy = this.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
            foreach (PropertyInfo property in propertiesToCopy)
            {
                if (property.CanWrite)
                {
                    object value = property.GetValue(configToApply);
                    if (value != null)
                    {
                        property.SetValue(this, value);
                    }
                }
            }

            // Member SPTrustName is not exposed through a property, so it must be set explicitly
            this.SPTrustName = configToApply.SPTrustName;
        }

        /// <summary>
        /// Returns a copy of the current object. This copy does not have any member of the base SharePoint base class set
        /// </summary>
        /// <returns></returns>
        public AzureCPConfig CopyConfiguration()
        {
            // Cannot use reflection here to copy object because of the calls to methods CopyConfiguration() on some properties
            AzureCPConfig copy = new AzureCPConfig();
            copy.SPTrustName = this.SPTrustName;
            copy.AzureTenants = new List<AzureTenant>(this.AzureTenants);
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
            copy.EnableRetry = this.EnableRetry;
            copy.Timeout = this.Timeout;
            copy.CustomData = this.CustomData;
            copy.MaxSearchResultsCount = this.MaxSearchResultsCount;
            copy.FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly;
            return copy;
        }

        public void ResetClaimTypesList()
        {
            ClaimTypes.Clear();
            ClaimTypes = ReturnDefaultClaimTypesConfig(this.SPTrustName);
            ClaimsProviderLogging.Log($"Claim types list of configuration '{Name}' was successfully reset to default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// If AzureCP is associated with a SPTrustedLoginProvider, create its configuration with default settings and save it into configuration database. If it already exists, it will be replaced.
        /// </summary>
        /// <returns></returns>
        public static AzureCPConfig CreateDefaultConfiguration()
        {
            SPTrustedLoginProvider spTrust = AzureCP.GetSPTrustAssociatedWithCP(AzureCP._ProviderInternalName);
            if (spTrust == null)
            {
                return null;
            }
            else
            {
                return CreateConfiguration(ClaimsProviderConstants.CONFIG_ID, ClaimsProviderConstants.CONFIG_NAME, spTrust.Name);
            }
        }

        /// <summary>
        /// Create a persisted object with default configuration of AzureCP.
        /// </summary>
        /// <param name="persistedObjectID">GUID of the configuration, stored as a persisted object into SharePoint configuration database</param>
        /// <param name="persistedObjectName">Name of the configuration, stored as a persisted object into SharePoint configuration database</param>
        /// <param name="spTrustName">Name of the SPTrustedLoginProvider that claims provider is associated with</param>
        /// <returns></returns>
        public static AzureCPConfig CreateConfiguration(string persistedObjectID, string persistedObjectName, string spTrustName)
        {
            if (String.IsNullOrEmpty(spTrustName))
            {
                throw new ArgumentNullException("spTrustName");
            }

            // Ensure it doesn't already exists and delete it if so
            AzureCPConfig existingConfig = AzureCPConfig.GetConfiguration(persistedObjectName, String.Empty);
            if (existingConfig != null)
            {
                DeleteConfiguration(persistedObjectName);
            }

            ClaimsProviderLogging.Log($"Creating configuration '{persistedObjectName}' with Id {persistedObjectID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);
            AzureCPConfig PersistedObject = new AzureCPConfig(persistedObjectName, SPFarm.Local, spTrustName);
            PersistedObject.ResetCurrentConfiguration();
            PersistedObject.Id = new Guid(persistedObjectID);
            PersistedObject.Update();
            ClaimsProviderLogging.Log($"Created configuration '{persistedObjectName}' with Id {PersistedObject.Id}", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
            return PersistedObject;
        }

        /// <summary>
        /// Generate and return default configuration
        /// </summary>
        /// <returns></returns>
        public static IAzureCPConfiguration ReturnDefaultConfiguration(string spTrustName)
        {
            AzureCPConfig defaultConfig = new AzureCPConfig();
            defaultConfig.SPTrustName = spTrustName;
            defaultConfig.AzureTenants = new List<AzureTenant>();
            defaultConfig.ClaimTypes = ReturnDefaultClaimTypesConfig(spTrustName);
            return defaultConfig;
        }

        /// <summary>
        /// Generate and return default claim types configuration list
        /// </summary>
        /// <returns></returns>
        public static ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig(string spTrustName)
        {
            if (String.IsNullOrWhiteSpace(spTrustName))
            {
                throw new ArgumentNullException("spTrustName cannot be null.");
            }

            SPTrustedLoginProvider spTrust = SPSecurityTokenServiceManager.Local.TrustedLoginProviders.GetProviderByName(spTrustName);
            if (spTrust == null)
            {
                ClaimsProviderLogging.Log($"SPTrustedLoginProvider '{spTrustName}' was not found ", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
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

        /// <summary>
        /// Delete persisted object from configuration database
        /// </summary>
        /// <param name="persistedObjectName">Name of persisted object to delete</param>
        public static void DeleteConfiguration(string persistedObjectName)
        {
            AzureCPConfig config = AzureCPConfig.GetConfiguration(persistedObjectName, String.Empty);
            if (config == null)
            {
                ClaimsProviderLogging.Log($"Configuration '{persistedObjectName}' was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
                return;
            }
            config.Delete();
            ClaimsProviderLogging.Log($"Configuration '{persistedObjectName}' was successfully deleted from configuration database", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Check if current configuration is compatible with current version of AzureCP, and fix it if not. If object comes from configuration database, changes are committed in configuration database
        /// </summary>
        /// <param name="spTrustName">Name of the SPTrust if it changed, null or empty string otherwise</param>
        /// <returns>Bollean indicates whether the configuration was updated in configuration database</returns>
        public bool CheckAndCleanConfiguration(string spTrustName)
        {
            // ClaimsProviderConstants.ClaimsProviderVersion can be null if assembly was removed from GAC
            if (String.IsNullOrEmpty(ClaimsProviderConstants.ClaimsProviderVersion))
            {
                return false;
            }

            bool configUpdated = false;

            if (!String.IsNullOrEmpty(spTrustName) && !String.Equals(this.SPTrustName, spTrustName, StringComparison.InvariantCultureIgnoreCase))
            {
                ClaimsProviderLogging.Log($"Updated property SPTrustName from \"{this.SPTrustName}\" to \"{spTrustName}\" in configuration \"{base.DisplayName}\".",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
                this.SPTrustName = spTrustName;
                configUpdated = true;
            }

            if (!String.Equals(this.ClaimsProviderVersion, ClaimsProviderConstants.ClaimsProviderVersion, StringComparison.InvariantCultureIgnoreCase))
            {
                // Detect if current assembly has a version different than AzureCPConfig.ClaimsProviderVersion. If so, config needs a sanity check
                ClaimsProviderLogging.Log($"Updated property ClaimsProviderVersion from \"{this.ClaimsProviderVersion}\" to \"{ClaimsProviderConstants.ClaimsProviderVersion}\" in configuration \"{base.DisplayName}\".",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
                this.ClaimsProviderVersion = ClaimsProviderConstants.ClaimsProviderVersion;
                configUpdated = true;
            }
            else if (!String.IsNullOrEmpty(this.SPTrustName))
            {
                // ClaimTypeConfigCollection.SPTrust is not persisted so it should always be set explicitely
                // Done in "else if" to not set this.ClaimTypes.SPTrust if we are not sure that this.ClaimTypes is in a good state
                this.ClaimTypes.SPTrust = this.SPTrust;
            }

            // Either claims provider was associated to a new SPTrustedLoginProvider
            // Or version of the current assembly changed (upgrade)
            // So let's do a sanity check of the configuration
            if (configUpdated)
            {
                try
                {
                    // If AzureCP was updated from a version < v12, this.ClaimTypes.Count will throw a NullReferenceException
                    int testClaimTypeCollection = this.ClaimTypes.Count;
                }
                catch (NullReferenceException)
                {
                    this.ClaimTypes = ReturnDefaultClaimTypesConfig(this.SPTrustName);
                    configUpdated = true;
                }

                if (!String.IsNullOrEmpty(this.SPTrustName))
                {
                    // ClaimTypeConfigCollection.SPTrust is not persisted so it should always be set explicitely
                    this.ClaimTypes.SPTrust = this.SPTrust;
                }


                // Starting with v13, identity claim type is automatically detected and added when list is reset to default (so it should always be present)
                // And it has its own class IdentityClaimTypeConfig that must be present as is to work correctly with Guest accounts
                // Since this is fixed by resetting claim type config list, this also addresses the duplicate DirectoryObjectProperty per EntityType constraint
                if (this.SPTrust != null)
                {
                    ClaimTypeConfig identityCTConfig = this.ClaimTypes.FirstOrDefault(x => String.Equals(x.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase));
                    if (identityCTConfig == null || !(identityCTConfig is IdentityClaimTypeConfig))
                    {
                        this.ClaimTypes = ReturnDefaultClaimTypesConfig(this.SPTrustName);
                        ClaimsProviderLogging.Log($"Claim types configuration list was reset because the identity claim type was either not found or not configured correctly",
                           TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
                        configUpdated = true;
                    }
                }

                // Starting with v13, adding 2 times a ClaimTypeConfig with the same EntityType and same DirectoryObjectProperty throws an InvalidOperationException
                // But this was possible before, so list this.ClaimTypes must be checked to be sure we are not in this scenario, and cleaned if so
                foreach (DirectoryObjectType entityType in Enum.GetValues(typeof(DirectoryObjectType)))
                {
                    var duplicatedPropertiesList = this.ClaimTypes.Where(x => x.EntityType == entityType)   // Check 1 EntityType
                                                              .GroupBy(x => x.DirectoryObjectProperty)      // Group by DirectoryObjectProperty
                                                              .Select(x => new
                                                              {
                                                                  DirectoryObjectProperty = x.Key,
                                                                  ObjectCount = x.Count()                   // For each DirectoryObjectProperty, how many items found
                                                              })
                                                              .Where(x => x.ObjectCount > 1);               // Keep only DirectoryObjectProperty found more than 1 time (for a given EntityType)
                    foreach (var duplicatedProperty in duplicatedPropertiesList)
                    {
                        ClaimTypeConfig ctConfigToDelete = null;
                        if (SPTrust != null && entityType == DirectoryObjectType.User)
                        {
                            ctConfigToDelete = this.ClaimTypes.FirstOrDefault(x => x.DirectoryObjectProperty == duplicatedProperty.DirectoryObjectProperty && x.EntityType == entityType && !String.Equals(x.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase));
                        }
                        else
                        {
                            ctConfigToDelete = this.ClaimTypes.FirstOrDefault(x => x.DirectoryObjectProperty == duplicatedProperty.DirectoryObjectProperty && x.EntityType == entityType);
                        }

                        this.ClaimTypes.Remove(ctConfigToDelete);
                        configUpdated = true;
                        ClaimsProviderLogging.Log($"Removed claim type '{ctConfigToDelete.ClaimType}' from claim types configuration list because it duplicates property {ctConfigToDelete.DirectoryObjectProperty}",
                           TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
                    }
                }

                if (Version > 0)
                {
                    try
                    {
                        // SPContext may be null if code does not run in a SharePoint process (e.g. in unit tests process)
                        if (SPContext.Current != null) { SPContext.Current.Web.AllowUnsafeUpdates = true; }
                        this.Update();
                        ClaimsProviderLogging.Log($"Configuration '{this.Name}' was upgraded in configuration database and some settings were updated or reset to their default configuration",
                            TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
                    }
                    catch (Exception)
                    {
                        // It may fail if current user doesn't have permission to update the object in configuration database
                        ClaimsProviderLogging.Log($"Configuration '{this.Name}' was upgraded locally, but changes could not be applied in configuration database. Please visit admin pages in central administration to upgrade configuration globally.",
                            TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
                    }
                    finally
                    {
                        if (SPContext.Current != null) { SPContext.Current.Web.AllowUnsafeUpdates = false; }
                    }
                }
            }
            return configUpdated;
        }

        /// <summary>
        /// Return the Azure AD tenant in the current configuration based on its name.
        /// </summary>
        /// <param name="azureTenantName">Name of the tenant, for example TENANTNAME.onMicrosoft.com.</param>
        /// <returns>AzureTenant found in the current configuration.</returns>
        public AzureTenant GetAzureTenantByName(string azureTenantName)
        {
            AzureTenant match = null;
            foreach (AzureTenant tenant in this.AzureTenants)
            {
                if (String.Equals(tenant.Name, azureTenantName, StringComparison.InvariantCultureIgnoreCase))
                {
                    match = tenant;
                    break;
                }
            }
            return match;
        }
    }

    public class AzureTenant : SPAutoSerializingObject
    {
        public Guid Identifier
        {
            get => Id;
            set => Id = value;
        }
        [Persisted]
        private Guid Id = Guid.NewGuid();

        /// <summary>
        /// Name of the tenant, e.g. TENANTNAME.onMicrosoft.com
        /// </summary>
        public string Name
        {
            get => TenantName;
            set => TenantName = value;
        }
        [Persisted]
        private string TenantName;

        /// <summary>
        /// Application ID of the application created in Azure AD tenant to authorize AzureCP
        /// </summary>
        public string ApplicationId
        {
            get => ClientId;
            set => ClientId = value;
        }
        [Persisted]
        private string ClientId;

        /// <summary>
        /// Password of the application
        /// </summary>
        public string ApplicationSecret
        {
            get => ClientSecret;
            set => ClientSecret = value;
        }
        [Persisted]
        private string ClientSecret;

        /// <summary>
        /// Set to true to return only Member users from this tenant
        /// </summary>
        public bool ExcludeMembers
        {
            get => ExcludeMemberUsers;
            set => ExcludeMemberUsers = value;
        }
        [Persisted]
        private bool ExcludeMemberUsers = false;

        /// <summary>
        /// Set to true to return only Guest users from this tenant
        /// </summary>
        public bool ExcludeGuests
        {
            get => ExcludeGuestUsers;
            set => ExcludeGuestUsers = value;
        }
        [Persisted]
        private bool ExcludeGuestUsers = false;

        /// <summary>
        /// Client ID of AD Connect used in extension attribues
        /// </summary>
        [Persisted]
        private Guid ExtensionAttributesApplicationIdPersisted;

        public Guid ExtensionAttributesApplicationId
        {
            get => ExtensionAttributesApplicationIdPersisted;
            set => ExtensionAttributesApplicationIdPersisted = value;
        }

        public X509Certificate2 ClientCertificatePrivateKey
        {
            get
            {
                return m_ClientCertificatePrivateKey;
            }
            set
            {
                if (value == null) { return; }
                m_ClientCertificatePrivateKey = value;
                try
                {
                    // https://stackoverflow.com/questions/32354790/how-to-check-is-x509certificate2-exportable-or-not
                    m_ClientCertificatePrivateKeyRawData = value.Export(X509ContentType.Pfx, ClaimsProviderConstants.ClientCertificatePrivateKeyPassword);
                }
                catch (CryptographicException ex)
                {
                    // X509Certificate2.Export() is expected to fail if the private key is not exportable, which depends on the X509KeyStorageFlags used when creating the X509Certificate2 object
                    //ClaimsProviderLogging.LogException(AzureCP._ProviderInternalName, $"while setting the certificate for tenant '{this.Name}'. Is the private key of the certificate exportable?", TraceCategory.Core, ex);
                    //throw;  // The caller should be informed that the certificate could not be set
                }
            }
        }
        private X509Certificate2 m_ClientCertificatePrivateKey;
        [Persisted]
        private byte[] m_ClientCertificatePrivateKeyRawData;

        public string AuthenticationMode
        {
            get
            {
                return String.IsNullOrWhiteSpace(this.ClientSecret) ? "ClientCertificate" : "ClientSecret";
            }
        }

        public AzureCloudInstance CloudInstance
        {
            get => (AzureCloudInstance)Enum.Parse(typeof(AzureCloudInstance), m_CloudInstance);
            set => m_CloudInstance = value.ToString();
        }
        [Persisted]
        private string m_CloudInstance = AzureCloudInstance.AzurePublic.ToString();

        /// <summary>
        /// Instance of the IAuthenticationProvider class for this specific Azure AD tenant
        /// </summary>
        //private AADAppOnlyAuthenticationProvider AuthenticationProvider { get; set; }

        public GraphServiceClient GraphService { get; set; }

        public string UserFilter { get; set; }
        public string GroupFilter { get; set; }
        public string[] UserSelect { get; set; }
        public string[] GroupSelect { get; set; }

        public AzureTenant()
        {
        }

        protected override void OnDeserialization()
        {
            if (m_ClientCertificatePrivateKeyRawData != null)
            {
                try
                {
                    // EphemeralKeySet: Keep the private key in-memory, it won't be written to disk - https://www.pkisolutions.com/handling-x509keystorageflags-in-applications/
                    m_ClientCertificatePrivateKey = ImportPfxCertificateBlob(m_ClientCertificatePrivateKeyRawData, ClaimsProviderConstants.ClientCertificatePrivateKeyPassword, X509KeyStorageFlags.EphemeralKeySet);
                }
                catch (CryptographicException ex)
                {
                    ClaimsProviderLogging.LogException(AzureCP._ProviderInternalName, $"while deserializating the certificate for tenant '{this.Name}'.", TraceCategory.Core, ex);
                }
            }
        }

        /// <summary>
        /// Set properties AuthenticationProvider and GraphService
        /// </summary>
        public void InitializeGraphForAppOnlyAuth(string claimsProviderName, int timeout)
        {
            try
            {
                TokenCredential tokenCredential;
                TokenCredentialOptions tokenCredentialOptions = new TokenCredentialOptions();
                tokenCredentialOptions.AuthorityHost = ClaimsProviderConstants.AzureCloudEndpoints.SingleOrDefault(kvp => kvp.Key == this.CloudInstance).Value;

                if (!String.IsNullOrWhiteSpace(ClientSecret))
                {
                    //this.AuthenticationProvider = new AADAppOnlyAuthenticationProvider(this.CloudInstance, this.Name, this.ApplicationId, this.ApplicationSecret, claimsProviderName, timeout);
                    tokenCredential = new ClientSecretCredential(this.Name, this.ApplicationId, this.ApplicationSecret, tokenCredentialOptions);
                }
                else
                {
                    //this.AuthenticationProvider = new AADAppOnlyAuthenticationProvider(this.CloudInstance, this.Name, this.ApplicationId, this.ClientCertificatePrivateKey, claimsProviderName, timeout);
                    tokenCredential = new ClientCertificateCredential(this.Name, this.ApplicationId, this.ClientCertificatePrivateKey, tokenCredentialOptions);
                }
                this.GraphService = new GraphServiceClient(tokenCredential, new[] { "https://graph.microsoft.com/.default" });
                //UriBuilder graphUriBuilder = new UriBuilder(this.AuthenticationProvider.GraphServiceEndpoint);
                //graphUriBuilder.Path = $"/{ClaimsProviderConstants.GraphServiceEndpointVersion}";
                //this.GraphService = new GraphServiceClient(graphUriBuilder.ToString(), this.AuthenticationProvider);

            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(AzureCP._ProviderInternalName, $"while setting client context for tenant '{this.Name}'.", TraceCategory.Core, ex);
            }
        }

        /// <summary>
        /// Returns a copy of the current object. This copy does not have any member of the base SharePoint base class set
        /// </summary>
        /// <returns></returns>
        internal AzureTenant CopyConfiguration()
        {
            AzureTenant copy = new AzureTenant();
            // Copy non-inherited public properties
            PropertyInfo[] propertiesToCopy = this.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);
            foreach (PropertyInfo property in propertiesToCopy)
            {
                if (property.CanWrite)
                {
                    object value = property.GetValue(this);
                    if (value != null)
                    {
                        property.SetValue(copy, value);
                    }
                }
            }
            return copy;
        }

        /// <summary>
        /// Update the credentials used to connect to the Azure AD tenant
        /// </summary>
        /// <param name="newApplicationSecret">New application (client) secret</param>
        public void UpdateCredentials(string newApplicationSecret)
        {
            SetCredentials(this.ApplicationId, newApplicationSecret);
        }

        /// <summary>
        /// Set the credentials used to connect to the Azure AD tenant
        /// </summary>
        /// <param name="applicationId">Application (client) ID</param>
        /// <param name="applicationSecret">Application (client) secret</param>
        public void SetCredentials(string applicationId, string applicationSecret)
        {
            this.ApplicationId = applicationId;
            this.ApplicationSecret = applicationSecret;
            this.ClientCertificatePrivateKey = null;
        }

        /// <summary>
        /// Update the credentials used to connect to the Azure AD tenant
        /// </summary>
        /// <param name="newCertificate">New certificate with its private key</param>
        public void UpdateCredentials(X509Certificate2 newCertificate)
        {
            SetCredentials(this.ApplicationId, newCertificate);
        }

        /// <summary>
        /// Set the credentials used to connect to the Azure AD tenant
        /// </summary>
        /// <param name="applicationId">Application (client) secret</param>
        /// <param name="certificate">Certificate with its private key</param>
        public void SetCredentials(string applicationId, X509Certificate2 certificate)
        {
            this.ApplicationId = applicationId;
            this.ApplicationSecret = String.Empty;
            this.ClientCertificatePrivateKey = certificate;
        }

        /// <summary>
        /// Import the input blob certificate into a pfx X509Certificate2 object
        /// </summary>
        /// <param name="blob"></param>
        /// <param name="certificatePassword"></param>
        /// <param name="keyStorageFlags"></param>
        /// <returns></returns>
        public static X509Certificate2 ImportPfxCertificateBlob(byte[] blob, string certificatePassword, X509KeyStorageFlags keyStorageFlags)
        {
            if (X509Certificate2.GetCertContentType(blob) != X509ContentType.Pfx)
            {
                return null;
            }

            if (String.IsNullOrWhiteSpace(certificatePassword))
            {
                // If passwordless, import private key as documented in https://support.microsoft.com/en-us/topic/kb5025823-change-in-how-net-applications-import-x-509-certificates-bf81c936-af2b-446e-9f7a-016f4713b46b
                return new X509Certificate2(blob, (string)null, keyStorageFlags);
            }
            else
            {
                return new X509Certificate2(blob, certificatePassword, keyStorageFlags);
            }
        }
    }

    /// <summary>
    /// Contains information about current operation
    /// </summary>
    public class OperationContext
    {
        /// <summary>
        /// Indicates what kind of operation SharePoint is requesting
        /// </summary>
        public OperationType OperationType
        {
            get => _OperationType;
            set => _OperationType = value;
        }
        private OperationType _OperationType;

        /// <summary>
        /// Set only if request is a validation or an augmentation, to the incoming entity provided by SharePoint
        /// </summary>
        public SPClaim IncomingEntity
        {
            get => _IncomingEntity;
            set => _IncomingEntity = value;
        }
        private SPClaim _IncomingEntity;

        /// <summary>
        /// User submitting the query in the poeple picker, retrieved from HttpContext. Can be null
        /// </summary>
        public SPClaim UserInHttpContext
        {
            get => _UserInHttpContext;
            set => _UserInHttpContext = value;
        }
        private SPClaim _UserInHttpContext;

        /// <summary>
        /// Uri provided by SharePoint
        /// </summary>
        public Uri UriContext
        {
            get => _UriContext;
            set => _UriContext = value;
        }
        private Uri _UriContext;

        /// <summary>
        /// EntityTypes expected by SharePoint in the entities returned
        /// </summary>
        public DirectoryObjectType[] DirectoryObjectTypes
        {
            get => _DirectoryObjectTypes;
            set => _DirectoryObjectTypes = value;
        }
        private DirectoryObjectType[] _DirectoryObjectTypes;

        public string HierarchyNodeID
        {
            get => _HierarchyNodeID;
            set => _HierarchyNodeID = value;
        }
        private string _HierarchyNodeID;

        public int MaxCount
        {
            get => _MaxCount;
            set => _MaxCount = value;
        }
        private int _MaxCount;

        /// <summary>
        /// If request is a validation: contains the value of the SPClaim. If request is a search: contains the input
        /// </summary>
        public string Input
        {
            get => _Input;
            set => _Input = value;
        }
        private string _Input;

        public bool InputHasKeyword
        {
            get => _InputHasKeyword;
            set => _InputHasKeyword = value;
        }
        private bool _InputHasKeyword;

        /// <summary>
        /// Indicates if search operation should return only results that exactly match the Input
        /// </summary>
        public bool ExactSearch
        {
            get => _ExactSearch;
            set => _ExactSearch = value;
        }
        private bool _ExactSearch;

        /// <summary>
        /// Set only if request is a validation or an augmentation, to the ClaimTypeConfig that matches the ClaimType of the incoming entity
        /// </summary>
        public ClaimTypeConfig IncomingEntityClaimTypeConfig
        {
            get => _IncomingEntityClaimTypeConfig;
            set => _IncomingEntityClaimTypeConfig = value;
        }
        private ClaimTypeConfig _IncomingEntityClaimTypeConfig;

        /// <summary>
        /// Contains the relevant list of ClaimTypeConfig for every type of request. In case of validation or augmentation, it will contain only 1 item.
        /// </summary>
        public List<ClaimTypeConfig> CurrentClaimTypeConfigList
        {
            get => _CurrentClaimTypeConfigList;
            set => _CurrentClaimTypeConfigList = value;
        }
        private List<ClaimTypeConfig> _CurrentClaimTypeConfigList;

        public OperationContext(IAzureCPConfiguration currentConfiguration, OperationType currentRequestType, List<ClaimTypeConfig> processedClaimTypeConfigList, string input, SPClaim incomingEntity, Uri context, string[] entityTypes, string hierarchyNodeID, int maxCount)
        {
            this.OperationType = currentRequestType;
            this.Input = input;
            this.IncomingEntity = incomingEntity;
            this.UriContext = context;
            this.HierarchyNodeID = hierarchyNodeID;
            this.MaxCount = maxCount;

            if (entityTypes != null)
            {
                List<DirectoryObjectType> aadEntityTypes = new List<DirectoryObjectType>();
                if (entityTypes.Contains(SPClaimEntityTypes.User))
                {
                    aadEntityTypes.Add(DirectoryObjectType.User);
                }
                if (entityTypes.Contains(ClaimsProviderConstants.GroupClaimEntityType))
                {
                    aadEntityTypes.Add(DirectoryObjectType.Group);
                }
                this.DirectoryObjectTypes = aadEntityTypes.ToArray();
            }

            HttpContext httpctx = HttpContext.Current;
            if (httpctx != null)
            {
                WIF4_5.ClaimsPrincipal cp = httpctx.User as WIF4_5.ClaimsPrincipal;
                if (cp != null)
                {
                    if (SPClaimProviderManager.IsEncodedClaim(cp.Identity.Name))
                    {
                        this.UserInHttpContext = SPClaimProviderManager.Local.DecodeClaimFromFormsSuffix(cp.Identity.Name);
                    }
                    else
                    {
                        // This code is reached only when called from central administration: current user is always a Windows user
                        this.UserInHttpContext = SPClaimProviderManager.Local.ConvertIdentifierToClaim(cp.Identity.Name, SPIdentifierTypes.WindowsSamAccountName);
                    }
                }
            }

            if (currentRequestType == OperationType.Validation)
            {
                this.InitializeValidation(processedClaimTypeConfigList);
            }
            else if (currentRequestType == OperationType.Search)
            {
                this.InitializeSearch(processedClaimTypeConfigList, currentConfiguration.FilterExactMatchOnly);
            }
            else if (currentRequestType == OperationType.Augmentation)
            {
                this.InitializeAugmentation(processedClaimTypeConfigList);
            }
        }

        /// <summary>
        /// Validation is when SharePoint expects exactly 1 PickerEntity from the incoming SPClaim
        /// </summary>
        /// <param name="processedClaimTypeConfigList"></param>
        protected void InitializeValidation(List<ClaimTypeConfig> processedClaimTypeConfigList)
        {
            if (this.IncomingEntity == null) { throw new ArgumentNullException("IncomingEntity"); }
            this.IncomingEntityClaimTypeConfig = processedClaimTypeConfigList.FirstOrDefault(x =>
               String.Equals(x.ClaimType, this.IncomingEntity.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
               !x.UseMainClaimTypeOfDirectoryObject);

            if (this.IncomingEntityClaimTypeConfig == null)
            {
                ClaimsProviderLogging.Log($"[{AzureCP._ProviderInternalName}] Unable to validate entity \"{this.IncomingEntity.Value}\" because its claim type \"{this.IncomingEntity.ClaimType}\" was not found in the ClaimTypes list of current configuration.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                throw new InvalidOperationException($"[{AzureCP._ProviderInternalName}] Unable validate entity \"{this.IncomingEntity.Value}\" because its claim type \"{this.IncomingEntity.ClaimType}\" was not found in the ClaimTypes list of current configuration.");
            }

            // CurrentClaimTypeConfigList must also be set
            this.CurrentClaimTypeConfigList = new List<ClaimTypeConfig>(1);
            this.CurrentClaimTypeConfigList.Add(this.IncomingEntityClaimTypeConfig);
            this.ExactSearch = true;
            this.Input = this.IncomingEntity.Value;
        }

        /// <summary>
        /// Search is when SharePoint expects a list of any PickerEntity that match input provided
        /// </summary>
        /// <param name="processedClaimTypeConfigList"></param>
        protected void InitializeSearch(List<ClaimTypeConfig> processedClaimTypeConfigList, bool exactSearch)
        {
            this.ExactSearch = exactSearch;
            if (!String.IsNullOrEmpty(this.HierarchyNodeID))
            {
                // Restrict search to ClaimType currently selected in the hierarchy (may return multiple results if identity claim type)
                CurrentClaimTypeConfigList = processedClaimTypeConfigList.FindAll(x =>
                    String.Equals(x.ClaimType, this.HierarchyNodeID, StringComparison.InvariantCultureIgnoreCase) &&
                    this.DirectoryObjectTypes.Contains(x.EntityType));
            }
            else
            {
                // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                CurrentClaimTypeConfigList = processedClaimTypeConfigList.FindAll(x => this.DirectoryObjectTypes.Contains(x.EntityType));
            }
        }

        protected void InitializeAugmentation(List<ClaimTypeConfig> processedClaimTypeConfigList)
        {
            if (this.IncomingEntity == null) { throw new ArgumentNullException("IncomingEntity"); }
            this.IncomingEntityClaimTypeConfig = processedClaimTypeConfigList.FirstOrDefault(x =>
               String.Equals(x.ClaimType, this.IncomingEntity.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
               !x.UseMainClaimTypeOfDirectoryObject);

            if (this.IncomingEntityClaimTypeConfig == null)
            {
                ClaimsProviderLogging.Log($"[{AzureCP._ProviderInternalName}] Unable to augment entity \"{this.IncomingEntity.Value}\" because its claim type \"{this.IncomingEntity.ClaimType}\" was not found in the ClaimTypes list of current configuration.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                throw new InvalidOperationException($"[{AzureCP._ProviderInternalName}] Unable to augment entity \"{this.IncomingEntity.Value}\" because its claim type \"{this.IncomingEntity.ClaimType}\" was not found in the ClaimTypes list of current configuration.");
            }
        }
    }

    public enum AzureADObjectProperty
    {
        NotSet,
        AccountEnabled,
        Department,
        DisplayName,
        GivenName,
        Id,
        JobTitle,
        Mail,
        MobilePhone,
        OfficeLocation,
        Surname,
        UserPrincipalName,
        UserType,
        // https://github.com/Yvand/AzureCP/issues/77: Include all other String properties of class User - https://docs.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0#properties
        AgeGroup,
        City,
        CompanyName,
        ConsentProvidedForMinor,
        Country,
        EmployeeId,
        FaxNumber,
        LegalAgeGroupClassification,
        MailNickname,
        OnPremisesDistinguishedName,
        OnPremisesImmutableId,
        OnPremisesSecurityIdentifier,
        OnPremisesDomainName,
        OnPremisesSamAccountName,
        OnPremisesUserPrincipalName,
        PasswordPolicies,
        PostalCode,
        PreferredLanguage,
        State,
        StreetAddress,
        UsageLocation,
        AboutMe,
        MySite,
        PreferredName,
        ODataType,
        extensionAttribute1,
        extensionAttribute2,
        extensionAttribute3,
        extensionAttribute4,
        extensionAttribute5,
        extensionAttribute6,
        extensionAttribute7,
        extensionAttribute8,
        extensionAttribute9,
        extensionAttribute10,
        extensionAttribute11,
        extensionAttribute12,
        extensionAttribute13,
        extensionAttribute14,
        extensionAttribute15
    }

    public enum DirectoryObjectType
    {
        User,
        Group
    }

    public class AzureADUserTypeHelper
    {
        public const string GuestUserType = "Guest";
        public const string MemberUserType = "Member";
    }

    public enum OperationType
    {
        Search,
        Validation,
        Augmentation,
    }
}
