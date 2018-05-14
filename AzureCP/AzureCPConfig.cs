using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WIF4_5 = System.Security.Claims;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.Graph;
using System.Net.Http.Headers;
using static azurecp.ClaimsProviderLogging;
using System.Collections.ObjectModel;

namespace azurecp
{
    public interface IAzureCPConfiguration
    {
        List<AzureTenant> AzureTenants { get; set; }
        ClaimTypeConfigCollection ClaimTypes { get; set; }
        bool AlwaysResolveUserInput { get; set; }
        bool FilterExactMatchOnly { get; set; }
        bool EnableAugmentation { get; set; }
    }

    public class ClaimsProviderConstants
    {
        public const string AZURECPCONFIG_ID = "0E9F8FB6-B314-4CCC-866D-DEC0BE76C237";
        public const string AZURECPCONFIG_NAME = "AzureCPConfig";
        public const string AuthString = "https://login.windows.net/{0}";
        public const string ResourceUrl = "https://graph.windows.net";
        //public const int timeout
#if DEBUG
        public const int timeout = 500000;    // 1000 secs      1000000
#else
        public const int timeout = 10000;    // 10 secs
#endif

        public const string SearchPatternEquals = "{0} eq '{1}'";
        public const string SearchPatternStartsWith = "startswith({0}, '{1}')";
        public static string GroupClaimEntityType = SPClaimEntityTypes.FormsRole;
        public const bool EnforceOnly1ClaimTypeForGroup = true;
    }

    public class AzureCPConfig : SPPersistedObject, IAzureCPConfiguration
    {
        public List<AzureTenant> AzureTenants
        {
            get { return AzureTenantsPersisted; }
            set { AzureTenantsPersisted = value; }
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
                _ClaimTypesCollection = value.innerCol;
            }
        }
        [Persisted]
        private Collection<ClaimTypeConfig> _ClaimTypesCollection;

        private ClaimTypeConfigCollection _ClaimTypes;

        public bool AlwaysResolveUserInput
        {
            get { return AlwaysResolveUserInputPersisted; }
            set { AlwaysResolveUserInputPersisted = value; }
        }
        [Persisted]
        private bool AlwaysResolveUserInputPersisted;

        public bool FilterExactMatchOnly
        {
            get { return FilterExactMatchOnlyPersisted; }
            set { FilterExactMatchOnlyPersisted = value; }
        }
        [Persisted]
        private bool FilterExactMatchOnlyPersisted;

        public bool EnableAugmentation
        {
            get { return AugmentAADRolesPersisted; }
            set { AugmentAADRolesPersisted = value; }
        }
        [Persisted]
        private bool AugmentAADRolesPersisted = true;

        public AzureCPConfig(string persistedObjectName, SPPersistedObject parent) : base(persistedObjectName, parent)
        { }

        public AzureCPConfig() { }

        public AzureCPConfig(bool initializeConfiguration)
        {
            if (initializeConfiguration)
            {
                this.AzureTenants = new List<AzureTenant>();
                this.ClaimTypes = GetDefaultClaimTypesConfig();
            }
        }

        protected override bool HasAdditionalUpdateAccess()
        {
            return false;
        }

        public static AzureCPConfig GetConfiguration(string persistedObjectName)
        {
            SPPersistedObject parent = SPFarm.Local;
            try
            {
                AzureCPConfig persistedObject = parent.GetChild<AzureCPConfig>(persistedObjectName);
                return persistedObject;
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.Log($"Error while retrieving configuration '{persistedObjectName}': {ex.Message}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
            }
            return null;
        }

        /// <summary>
        /// Commit changes to configuration database
        /// </summary>
        public override void Update()
        {
            base.Update();
            ClaimsProviderLogging.Log($"Configuration '{base.DisplayName}' was updated successfully in configuration database.",
                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
        }

        public static AzureCPConfig ResetConfiguration(string persistedObjectName)
        {
            AzureCPConfig previousConfig = GetConfiguration(persistedObjectName);
            if (previousConfig == null) return null;
            Guid configId = previousConfig.Id;
            DeleteConfiguration(persistedObjectName);
            AzureCPConfig newConfig = CreateConfiguration(configId.ToString(), persistedObjectName);
            ClaimsProviderLogging.Log($"Configuration '{persistedObjectName}' was successfully reset to its default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            return newConfig;
        }

        public void ApplyConfiguration(AzureCPConfig configToApply)
        {
            this.AzureTenants = configToApply.AzureTenants;
            this.ClaimTypes = configToApply.ClaimTypes;
            this.AlwaysResolveUserInput = configToApply.AlwaysResolveUserInput;
            this.FilterExactMatchOnly = configToApply.FilterExactMatchOnly;
            this.EnableAugmentation = configToApply.EnableAugmentation;
        }

        public AzureCPConfig CopyCurrentObject()
        {
            //return this.Clone() as LDAPCPConfig;  // DOES NOT work
            AzureCPConfig copy = new AzureCPConfig(true);
            copy.AlwaysResolveUserInput = this.AlwaysResolveUserInput;
            copy.FilterExactMatchOnly = this.FilterExactMatchOnly;
            copy.EnableAugmentation = this.EnableAugmentation;
            copy.ClaimTypes = new ClaimTypeConfigCollection();
            foreach (ClaimTypeConfig currentObject in this.ClaimTypes)
            {
                copy.ClaimTypes.Add(currentObject.CopyCurrentObject());
            }
            copy.AzureTenants = new List<AzureTenant>(this.AzureTenants);
            return copy;
        }

        public void ResetClaimTypesList()
        {
            ClaimTypes.Clear();
            ClaimTypes = GetDefaultClaimTypesConfig();
            ClaimsProviderLogging.Log($"Claim types list of configuration '{Name}' was successfully reset to default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Create the persisted object that contains default configuration of AzureCP.
        /// It should be created only in central administration with application pool credentials
        /// because this is the only place where we are sure user has the permission to write in the config database
        /// </summary>
        public static AzureCPConfig CreateConfiguration(string persistedObjectID, string persistedObjectName)
        {
            // Ensure it doesn't already exists and delete it if so
            AzureCPConfig existingConfig = AzureCPConfig.GetConfiguration(persistedObjectName);
            if (existingConfig != null)
            {
                DeleteConfiguration(persistedObjectName);
            }

            ClaimsProviderLogging.Log($"Creating configuration '{persistedObjectName}' with Id {persistedObjectID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);
            AzureCPConfig PersistedObject = new AzureCPConfig(persistedObjectName, SPFarm.Local);
            PersistedObject.ResetCurrentConfiguration();
            PersistedObject.Id = new Guid(persistedObjectID);
            PersistedObject.AzureTenants = new List<AzureTenant>();
            PersistedObject.Update();
            ClaimsProviderLogging.Log($"Created configuration '{persistedObjectName}' with Id {PersistedObject.Id}", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
            return PersistedObject;
        }

        /// <summary>
        /// Set properties of current configuration to their default values
        /// </summary>
        /// <param name="persistedObjectName"></param>
        /// <returns></returns>
        public void ResetCurrentConfiguration()
        {
            AzureCPConfig defaultConfig = new AzureCPConfig(true);
            ApplyConfiguration(defaultConfig);
        }

        public static IAzureCPConfiguration GetDefaultConfiguration()
        {
            IAzureCPConfiguration defaultConfig = new AzureCPConfig(true);
            return defaultConfig;
        }

        public static ClaimTypeConfigCollection GetDefaultClaimTypesConfig()
        {
            return new ClaimTypeConfigCollection
            {
                // By default ACS issues those 3 claim types: ClaimTypes.Name ClaimTypes.GivenName and ClaimTypes.Surname.
                // But ClaimTypes.Name (http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name) is a reserved claim type in SharePoint that cannot be used in the SPTrust.

                // Alternatives claim types to ClaimTypes.Name that might be used as identity claim types:
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.UserPrincipalName, ClaimType=WIF4_5.ClaimTypes.Upn, DirectoryObjectType = AzureADObjectType.User},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.UserPrincipalName, ClaimType=WIF4_5.ClaimTypes.Email, DirectoryObjectType = AzureADObjectType.User},

                // Additional properties to find user
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.DisplayName, UseMainClaimTypeOfDirectoryObject=true, DirectoryObjectType = AzureADObjectType.User, EntityDataKey=PeopleEditorEntityDataKeys.DisplayName},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.GivenName, UseMainClaimTypeOfDirectoryObject=true, DirectoryObjectType = AzureADObjectType.User},//Yvan
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.Surname, UseMainClaimTypeOfDirectoryObject=true, DirectoryObjectType = AzureADObjectType.User},//Duhamel

                // Retrieve additional properties to populate metadata in SharePoint (no claim type and UseMainClaimTypeOfDirectoryObject = false)
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.Mail, DirectoryObjectType = AzureADObjectType.User, EntityDataKey=PeopleEditorEntityDataKeys.Email},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.MobilePhone, DirectoryObjectType = AzureADObjectType.User, EntityDataKey=PeopleEditorEntityDataKeys.MobilePhone},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.JobTitle, DirectoryObjectType = AzureADObjectType.User, EntityDataKey=PeopleEditorEntityDataKeys.JobTitle},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.Department, DirectoryObjectType = AzureADObjectType.User, EntityDataKey=PeopleEditorEntityDataKeys.Department},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.OfficeLocation, DirectoryObjectType = AzureADObjectType.User, EntityDataKey=PeopleEditorEntityDataKeys.Location},

                // Group
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.Id, ClaimType=WIF4_5.ClaimTypes.Role, DirectoryObjectType = AzureADObjectType.Group, DirectoryObjectPropertyToShowAsDisplayText=AzureADObjectProperty.DisplayName},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.DisplayName, DirectoryObjectType = AzureADObjectType.Group, UseMainClaimTypeOfDirectoryObject = true},
            };
        }

        public static void DeleteConfiguration(string persistedObjectName)
        {
            AzureCPConfig config = AzureCPConfig.GetConfiguration(persistedObjectName);
            if (config == null)
            {
                ClaimsProviderLogging.Log($"Configuration '{persistedObjectName}' was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
                return;
            }
            config.Delete();
            ClaimsProviderLogging.Log($"Configuration '{persistedObjectName}' was successfully deleted from configuration database", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
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

        [Persisted]
        public string AADInstance = "https://login.windows.net/{0}";

        /// <summary>
        /// Instance of the IAuthenticationProvider class for this specific Azure AD tenant
        /// </summary>
        private AADAppOnlyAuthenticationProvider AuthenticationProvider;

        public GraphServiceClient GraphService;

        public AzureTenant()
        {
        }

        /// <summary>
        /// Set properties AuthenticationProvider and GraphService
        /// </summary>
        public void SetAzureADContext()
        {
            try
            {
                this.AuthenticationProvider = new AADAppOnlyAuthenticationProvider(this.AADInstance, this.TenantName, this.ClientId, this.ClientSecret);
                this.GraphService = new GraphServiceClient(this.AuthenticationProvider);
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(AzureCP._ProviderInternalName, $"while setting client context for tenant '{this.TenantName}'.", TraceCategory.Core, ex);
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
        public AzureADObjectType[] DirectoryObjectTypes;
        public string HierarchyNodeID;
        public int MaxCount;

        /// <summary>
        /// If request is a validation: contains the value of the SPClaim. If request is a search: contains the input
        /// </summary>
        public string Input;
        public bool InputHasKeyword;

        /// <summary>
        /// Indicate if search operation should return only results that exactly match the Input
        /// </summary>
        public bool ExactSearch;

        /// <summary>
        /// Set only if request is a validation or an augmentation, to the ClaimTypeConfig that matches the ClaimType of the incoming entity
        /// </summary>
        public ClaimTypeConfig CurrentClaimTypeConfig;

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
                List<AzureADObjectType> aadEntityTypes = new List<AzureADObjectType>();
                if (entityTypes.Contains(SPClaimEntityTypes.User))
                    aadEntityTypes.Add(AzureADObjectType.User);
                if (entityTypes.Contains(ClaimsProviderConstants.GroupClaimEntityType))
                    aadEntityTypes.Add(AzureADObjectType.Group);
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
            this.CurrentClaimTypeConfig = processedClaimTypeConfigList.FirstOrDefault(x =>
               String.Equals(x.ClaimType, this.IncomingEntity.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
               !x.UseMainClaimTypeOfDirectoryObject);
            if (this.CurrentClaimTypeConfig == null) return;

            // ClaimTypeConfigList must also be set
            this.CurrentClaimTypeConfigList = new List<ClaimTypeConfig>(1);
            this.CurrentClaimTypeConfigList.Add(this.CurrentClaimTypeConfig);
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
                    this.DirectoryObjectTypes.Contains(x.DirectoryObjectType));
            }
            else
            {
                // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                CurrentClaimTypeConfigList = processedClaimTypeConfigList.FindAll(x => this.DirectoryObjectTypes.Contains(x.DirectoryObjectType));
            }
        }

        protected void InitializeAugmentation(List<ClaimTypeConfig> processedClaimTypeConfigList)
        {
            if (this.IncomingEntity == null) throw new ArgumentNullException("IncomingEntity");
            this.CurrentClaimTypeConfig = processedClaimTypeConfigList.FirstOrDefault(x =>
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

    public enum AzureADObjectType
    {
        User,
        Group
    }

    public class AzureADUserTypeHelper
    {
        public const string GuestUserType = "Guest";
        public const string MemberUserType = "Member";
        public const string PropertyNameContainingUserType = "UserType";
    }

    public enum OperationType
    {
        Search,
        Validation,
        Augmentation,
    }
}
