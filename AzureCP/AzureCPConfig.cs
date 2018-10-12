using Microsoft.Graph;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Web;
using static azurecp.ClaimsProviderLogging;
using WIF4_5 = System.Security.Claims;

namespace azurecp
{
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
    }

    public class ClaimsProviderConstants
    {
        public const string CONFIG_ID = "0E9F8FB6-B314-4CCC-866D-DEC0BE76C237";
        public const string CONFIG_NAME = "AzureCPConfig";
        public const string GraphAPIResource = "https://graph.microsoft.com/";
        public const string AuthorityUriTemplate = "https://login.windows.net/{0}";
        public const string ResourceUrl = "https://graph.windows.net";
        public const string SearchPatternEquals = "{0} eq '{1}'";
        public const string SearchPatternStartsWith = "startswith({0}, '{1}')";
        public const string IdentityConfigSearchPatternEquals = "({0} eq '{1}' and UserType eq '{2}')";
        public const string IdentityConfigSearchPatternStartsWith = "(startswith({0}, '{1}') and UserType eq '{2}')";
        public static string GroupClaimEntityType = SPClaimEntityTypes.FormsRole;
        public const bool EnforceOnly1ClaimTypeForGroup = true;     // In AzureCP, only 1 claim type can be used to create group permissions
        public const string DefaultMainGroupClaimType = WIF4_5.ClaimTypes.Role;
        public const string PUBLICSITEURL = "https://yvand.github.io/AzureCP/";
        public static string ClaimsProviderVersion
        {
            get
            {
                try
                {
                    return FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(AzureCP)).Location).FileVersion;
                }
                // If assembly was removed from the GAC, CLR will throw that a FileNotFoundException
                catch (System.IO.FileNotFoundException)
                {
                    return String.Empty;
                }
            }
        }

#if DEBUG
        public const int DEFAULT_TIMEOUT = 10000;
#else
        public const int DEFAULT_TIMEOUT = 4000;    // 4 secs
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
        /// Name of the SPTrustedLoginProvider where LDAPCP is enabled
        /// </summary>
        [Persisted]
        private string SPTrustName;

        private SPTrustedLoginProvider _SPTrust;
        private SPTrustedLoginProvider SPTrust
        {
            get
            {
                if (_SPTrust == null) _SPTrust = SPSecurityTokenServiceManager.Local.TrustedLoginProviders.GetProviderByName(SPTrustName);
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
            if (previousConfig == null) return null;
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

        public void ApplyConfiguration(AzureCPConfig configToApply)
        {
            this.SPTrustName = configToApply.SPTrustName;
            this.AzureTenants = configToApply.AzureTenants;
            this.ClaimTypes = configToApply.ClaimTypes;
            this.AlwaysResolveUserInput = configToApply.AlwaysResolveUserInput;
            this.FilterExactMatchOnly = configToApply.FilterExactMatchOnly;
            this.EnableAugmentation = configToApply.EnableAugmentation;
            this.EntityDisplayTextPrefix = configToApply.EntityDisplayTextPrefix;
            this.EnableRetry = configToApply.EnableRetry;
            this.Timeout = configToApply.Timeout;
            this.CustomData = configToApply.CustomData;
        }

        public AzureCPConfig CopyPersistedProperties()
        {
            AzureCPConfig copy = new AzureCPConfig();
            copy.SPTrustName = this.SPTrustName;
            copy.AzureTenants = new List<AzureTenant>(this.AzureTenants);
            copy.ClaimTypes = new ClaimTypeConfigCollection();
            copy.ClaimTypes.SPTrust = this.ClaimTypes.SPTrust;
            foreach (ClaimTypeConfig currentObject in this.ClaimTypes)
            {
                copy.ClaimTypes.Add(currentObject.CopyPersistedProperties(), false);
            }
            copy.AlwaysResolveUserInput = this.AlwaysResolveUserInput;
            copy.FilterExactMatchOnly = this.FilterExactMatchOnly;
            copy.EnableAugmentation = this.EnableAugmentation;
            copy.EntityDisplayTextPrefix = this.EntityDisplayTextPrefix;
            copy.EnableRetry = this.EnableRetry;
            copy.Timeout = this.Timeout;
            copy.CustomData = this.CustomData;
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
        /// Create a persisted object with default configuration of AzureCP.
        /// </summary>
        /// <param name="persistedObjectID"></param>
        /// <param name="persistedObjectName"></param>
        /// <returns></returns>
        public static AzureCPConfig CreateConfiguration(string persistedObjectID, string persistedObjectName, string spTrustName)
        {
            if (String.IsNullOrEmpty(spTrustName))
            {
                throw new ArgumentNullException("spTrust");
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
            if (String.IsNullOrWhiteSpace(spTrustName)) throw new ArgumentNullException("spTrustName cannot be null.");

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

                // Additional properties to populate metadata of entity created: no claim type set, EntityDataKey is set and UseMainClaimTypeOfDirectoryObject = false (default value)
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.Mail, EntityDataKey = PeopleEditorEntityDataKeys.Email},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.MobilePhone, EntityDataKey = PeopleEditorEntityDataKeys.MobilePhone},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.JobTitle, EntityDataKey = PeopleEditorEntityDataKeys.JobTitle},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.Department, EntityDataKey = PeopleEditorEntityDataKeys.Department},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation, EntityDataKey = PeopleEditorEntityDataKeys.Location},

                // Group
                new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, DirectoryObjectProperty = AzureADObjectProperty.Id, ClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType, DirectoryObjectPropertyToShowAsDisplayText = AzureADObjectProperty.DisplayName},
                new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, DirectoryObjectProperty = AzureADObjectProperty.DisplayName, UseMainClaimTypeOfDirectoryObject = true, EntityDataKey = PeopleEditorEntityDataKeys.DisplayName},
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
                return false;

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
                catch (NullReferenceException ex)
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
                            ctConfigToDelete = this.ClaimTypes.FirstOrDefault(x => x.DirectoryObjectProperty == duplicatedProperty.DirectoryObjectProperty && x.EntityType == entityType && !String.Equals(x.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase));
                        else
                            ctConfigToDelete = this.ClaimTypes.FirstOrDefault(x => x.DirectoryObjectProperty == duplicatedProperty.DirectoryObjectProperty && x.EntityType == entityType);

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
                        if (SPContext.Current != null) SPContext.Current.Web.AllowUnsafeUpdates = true;
                        this.Update();
                        ClaimsProviderLogging.Log($"Configuration '{this.Name}' was upgraded in configuration database and some settings were updated or reset to their default configuration",
                            TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
                    }
                    catch (Exception ex)
                    {
                        // It may fail if current user doesn't have permission to update the object in configuration database
                        ClaimsProviderLogging.Log($"Configuration '{this.Name}' was upgraded locally, but changes could not be applied in configuration database. Please visit admin pages in central administration to upgrade configuration globally.",
                            TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
                    }
                    finally
                    {
                        if (SPContext.Current != null) SPContext.Current.Web.AllowUnsafeUpdates = false;
                    }
                }
            }
            return configUpdated;
        }
    }


    public class AzureTenant : SPAutoSerializingObject
    {
        [Persisted]
        public Guid Id = Guid.NewGuid();

        /// <summary>
        /// Name of the tenant, e.g. TENANTNAME.onMicrosoft.com
        /// </summary>
        [Persisted]
        public string TenantName;

        /// <summary>
        /// Application ID of the application created in Azure AD tenant to authorize AzureCP
        /// </summary>
        [Persisted]
        public string ClientId;

        /// <summary>
        /// Password of the application
        /// </summary>
        [Persisted]
        public string ClientSecret;

        [Persisted]
        public bool MemberUserTypeOnly;

        /// <summary>
        /// Instance of the IAuthenticationProvider class for this specific Azure AD tenant
        /// </summary>
        private AADAppOnlyAuthenticationProvider AuthenticationProvider;

        public GraphServiceClient GraphService;

        public string UserFilter;
        public string GroupFilter;
        public string UserSelect;
        public string GroupSelect;

        public AzureTenant()
        {
        }

        /// <summary>
        /// Set properties AuthenticationProvider and GraphService
        /// </summary>
        public void SetAzureADContext(string claimsProviderName, int timeout)
        {
            try
            {
                this.AuthenticationProvider = new AADAppOnlyAuthenticationProvider(ClaimsProviderConstants.AuthorityUriTemplate, this.TenantName, this.ClientId, this.ClientSecret, claimsProviderName, timeout);
                this.GraphService = new GraphServiceClient(this.AuthenticationProvider);
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(AzureCP._ProviderInternalName, $"while setting client context for tenant '{this.TenantName}'.", TraceCategory.Core, ex);
            }
        }

        internal AzureTenant CopyPersistedProperties()
        {
            AzureTenant copy = new AzureTenant();
            copy.AuthenticationProvider = this.AuthenticationProvider;
            copy.ClientId = this.ClientId;
            copy.ClientSecret = this.ClientSecret;
            copy.GraphService = this.GraphService;
            copy.GroupFilter = this.GroupFilter;
            copy.GroupSelect = this.GroupSelect;
            copy.Id = this.Id;
            copy.MemberUserTypeOnly = this.MemberUserTypeOnly;
            copy.TenantName = this.TenantName;
            copy.UserFilter = this.UserFilter;
            copy.UserSelect = this.UserSelect;
            return copy;
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
        public OperationType OperationType;

        /// <summary>
        /// Set only if request is a validation or an augmentation, to the incoming entity provided by SharePoint
        /// </summary>
        public SPClaim IncomingEntity;

        /// <summary>
        /// User submitting the query in the poeple picker, retrieved from HttpContext. Can be null
        /// </summary>
        public SPClaim UserInHttpContext;

        /// <summary>
        /// Uri provided by SharePoint
        /// </summary>
        public Uri UriContext;

        /// <summary>
        /// EntityTypes expected by SharePoint in the entities returned
        /// </summary>
        public DirectoryObjectType[] DirectoryObjectTypes;
        public string HierarchyNodeID;
        public int MaxCount;

        /// <summary>
        /// If request is a validation: contains the value of the SPClaim. If request is a search: contains the input
        /// </summary>
        public string Input;
        public bool InputHasKeyword;

        /// <summary>
        /// Indicates if search operation should return only results that exactly match the Input
        /// </summary>
        public bool ExactSearch;

        /// <summary>
        /// Set only if request is a validation or an augmentation, to the ClaimTypeConfig that matches the ClaimType of the incoming entity
        /// </summary>
        public ClaimTypeConfig IncomingEntityClaimTypeConfig;

        /// <summary>
        /// Contains the relevant list of ClaimTypeConfig for every type of request. In case of validation or augmentation, it will contain only 1 item.
        /// </summary>
        public List<ClaimTypeConfig> CurrentClaimTypeConfigList;

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
                    aadEntityTypes.Add(DirectoryObjectType.User);
                if (entityTypes.Contains(ClaimsProviderConstants.GroupClaimEntityType))
                    aadEntityTypes.Add(DirectoryObjectType.Group);
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
            if (this.IncomingEntity == null) throw new ArgumentNullException("IncomingEntity");
            this.IncomingEntityClaimTypeConfig = processedClaimTypeConfigList.FirstOrDefault(x =>
               String.Equals(x.ClaimType, this.IncomingEntity.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
               !x.UseMainClaimTypeOfDirectoryObject);
            if (this.IncomingEntityClaimTypeConfig == null) return;

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
            if (this.IncomingEntity == null) throw new ArgumentNullException("IncomingEntity");
            this.IncomingEntityClaimTypeConfig = processedClaimTypeConfigList.FirstOrDefault(x =>
               String.Equals(x.ClaimType, this.IncomingEntity.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
               !x.UseMainClaimTypeOfDirectoryObject);
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
        UserType
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
