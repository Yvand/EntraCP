using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WIF = System.Security.Claims;
using System.Text.RegularExpressions;
using System.Web;
using Microsoft.Graph;
using System.Net.Http.Headers;
using static azurecp.AzureCPLogging;
using System.Collections.ObjectModel;
//using WIF3_5 = Microsoft.IdentityModel.Claims;

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
        public class GraphUserType
        {
            public const string Guest = "Guest";
            public const string Member = "Member";
        }

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
        public const string SearchPatternStartsWith = "startswith({0},'{1}')";
        public static string GroupClaimEntityType = SPClaimEntityTypes.FormsRole;
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
                //else
                //{
                //    _ClaimTypesCollection = _ClaimTypes.innerCol;
                //}
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
                AzureCPLogging.Log(String.Format("Error while retrieving SPPersistedObject {0}: {1}", ClaimsProviderConstants.AZURECPCONFIG_NAME, ex.Message), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
            }
            return null;
        }

        /// <summary>
        /// Commit changes in configuration database
        /// </summary>
        public override void Update()
        {
            base.Update();
            AzureCPLogging.Log($"PersistedObject {base.DisplayName} was updated successfully.",
                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
        }

        public static AzureCPConfig ResetConfiguration(string persistedObjectName)
        {
            AzureCPConfig previousConfig = GetConfiguration(persistedObjectName);
            if (previousConfig == null) return null;
            Guid configId = previousConfig.Id;
            DeleteAzureCPConfig(persistedObjectName);
            AzureCPConfig newConfig = CreatePersistedObject(configId.ToString(), persistedObjectName);
            AzureCPLogging.Log($"PersistedObject {persistedObjectName} was successfully reset to its default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            return newConfig;

            //AzureCPConfig persistedObject = GetConfiguration(persistedObjectName);
            //if (persistedObject != null)
            //{
            //    AzureCPConfig newPersistedObject = GetDefaultConfiguration(persistedObjectName);
            //    newPersistedObject.Update();

            //    AzureCPLogging.Log($"PersistedObject {persistedObjectName} was successfully reset to its default configuration",
            //        TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            //}
        }

        //public void ResetCurrentConfiguration()
        //{
        //    AzureCPConfig defaultConfiguration = SetDefaultConfiguration() as AzureCPConfig;
        //    this.ApplyConfiguration(defaultConfiguration);
        //    AzureCPLogging.Log($"PersistedObject {this.Name} was successfully reset to its default configuration",
        //        TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        //}

        public void ApplyConfiguration(AzureCPConfig configToApply)
        {
            this.AzureTenants = configToApply.AzureTenants;
            this.ClaimTypes = configToApply.ClaimTypes;
            this.AlwaysResolveUserInput = configToApply.AlwaysResolveUserInput;
            this.FilterExactMatchOnly = configToApply.FilterExactMatchOnly;
            this.EnableAugmentation = configToApply.EnableAugmentation;
        }

        public AzureCPConfig CloneInReadOnlyObject()
        {
            //return this.Clone() as LDAPCPConfig;
            AzureCPConfig readOnlyCopy = new AzureCPConfig(true);
            readOnlyCopy.AlwaysResolveUserInput = this.AlwaysResolveUserInput;
            readOnlyCopy.FilterExactMatchOnly = this.FilterExactMatchOnly;
            readOnlyCopy.EnableAugmentation = this.EnableAugmentation;
            readOnlyCopy.ClaimTypes = new ClaimTypeConfigCollection();
            foreach (ClaimTypeConfig currentObject in this.ClaimTypes)
            {
                readOnlyCopy.ClaimTypes.Add(currentObject.CopyPersistedProperties());
            }
            readOnlyCopy.AzureTenants = new List<AzureTenant>(this.AzureTenants);
            return readOnlyCopy;
        }

        public void ResetClaimTypesList()
        {
            ClaimTypes.Clear();
            ClaimTypes = GetDefaultClaimTypesConfig();
            AzureCPLogging.Log($"Claim types list of PersistedObject {Name} was successfully reset to default configuration",
                TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Create the persisted object that contains default configuration of AzureCP.
        /// It should be created only in central administration with application pool credentials
        /// because this is the only place where we are sure user has the permission to write in the config database
        /// </summary>
        public static AzureCPConfig CreatePersistedObject(string persistedObjectID, string persistedObjectName)
        {
            // Ensure it doesn't already exists and delete it if so
            AzureCPConfig existingConfig = AzureCPConfig.GetConfiguration(persistedObjectName);
            if (existingConfig != null)
            {
                DeleteAzureCPConfig(persistedObjectName);
            }

            AzureCPLogging.Log($"Creating persisted object {persistedObjectName} with Id {persistedObjectID}...", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
            AzureCPConfig PersistedObject = new AzureCPConfig(persistedObjectName, SPFarm.Local);
            PersistedObject.ResetCurrentConfiguration();
            PersistedObject.Id = new Guid(persistedObjectID);
            PersistedObject.AzureTenants = new List<AzureTenant>();
            PersistedObject.Update();
            AzureCPLogging.Log($"Created PersistedObject {PersistedObject.Name} with Id {PersistedObject.Id}", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
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
            //this.AlwaysResolveUserInput = defaultConfig.AlwaysResolveUserInput;
            //this.EnableAugmentation = defaultConfig.EnableAugmentation;
            //this.FilterExactMatchOnly = defaultConfig.FilterExactMatchOnly;

            //this.AzureTenants = new List<AzureTenant>();
            //this.ClaimTypes = GetDefaultClaimTypesConfig();
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
                //new AzureADObject{GraphProperty=GraphProperty.UserPrincipalName, ClaimType=WIF.ClaimTypes.Name, ClaimEntityType=SPClaimEntityTypes.User},//yvand@TENANTNAME.onmicrosoft.com

                // Alternatives claim types to ClaimTypes.Name that might be used as identity claim types:
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.UserPrincipalName, ClaimType=WIF.ClaimTypes.Upn, DirectoryObjectType = AzureADObjectType.User},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.UserPrincipalName, ClaimType=WIF.ClaimTypes.Email, DirectoryObjectType = AzureADObjectType.User},

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

                // Role
                //new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.Id, ClaimType=WIF.ClaimTypes.Role, DirectoryObjectType = AzureADObjectType.Group, DirectoryObjectPropertyToShowAsDisplayText=AzureADObjectProperty.DisplayName},
                new ClaimTypeConfig{DirectoryObjectProperty=AzureADObjectProperty.DisplayName, ClaimType=WIF.ClaimTypes.Role, DirectoryObjectType = AzureADObjectType.Group, DirectoryObjectPropertyToShowAsDisplayText=AzureADObjectProperty.DisplayName},
            };
        }

        public static void DeleteAzureCPConfig(string persistedObjectName)
        {
            AzureCPConfig config = AzureCPConfig.GetConfiguration(persistedObjectName);
            if (config == null)
            {
                AzureCPLogging.Log($"Persisted object {persistedObjectName} was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
                return;
            }
            config.Delete();
            AzureCPLogging.Log($"Persisted object {persistedObjectName} was successfully deleted from configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
        }
    }


    public class AzureTenant : SPAutoSerializingObject
    {
        [Persisted]
        public Guid Id = Guid.NewGuid();

        [Persisted]
        public string TenantName;

        [Persisted]
        public string ClientId;

        [Persisted]
        public string ClientSecret;

        [Persisted]
        public bool MemberUserTypeOnly;

        /// <summary>
        /// Access token used to connect to AAD. Should not be persisted or accessible outside of the assembly
        /// </summary>
        public string AccessToken = String.Empty;

        [Persisted]
        public string AADInstance = "https://login.windows.net/{0}";

        public AADAppOnlyAuthenticationProvider AuthenticationProvider;

        public GraphServiceClient GraphService;

        public AzureTenant()
        {
        }

        internal AzureTenant CopyPersistedProperties()
        {
            AzureTenant copy = new AzureTenant()
            {
                TenantName = this.TenantName,
                ClientId = this.ClientId,
                ClientSecret = this.ClientSecret,
                MemberUserTypeOnly = this.MemberUserTypeOnly,
                AADInstance = this.AADInstance,
                // This is done in SetAzureADContext
                //AuthenticationProvider = this.AuthenticationProvider,
                //GraphService = this.GraphService
            };
            return copy;
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
                AzureCPLogging.LogException(AzureCP._ProviderInternalName, $"while setting client context for tenant '{this.TenantName}'.", TraceCategory.Core, ex);
            }
        }
    }

    /// <summary>
    /// Contains information about current request
    /// </summary>
    public class RequestInformation
    {
        /// <summary>
        /// Current LDAPCP configuration
        /// </summary>
        //public IAzureCPConfiguration CurrentConfiguration;

        /// <summary>
        /// Indicates what kind of operation SharePoint is sending to LDAPCP
        /// </summary>
        public RequestType RequestType;

        /// <summary>
        /// SPClaim sent in parameter to LDAPCP. Can be null
        /// </summary>
        public SPClaim IncomingEntity;

        /// <summary>
        /// User submitting the query in the poeple picker, retrieved from HttpContext. Can be null
        /// </summary>
        public SPClaim UserInHttpContext;

        public Uri Context;
        //public string[] EntityTypes;
        public AzureADObjectType[] DirectoryObjectTypes;
        private string OriginalInput;
        public string HierarchyNodeID;
        public int MaxCount;

        public string Input;
        public bool InputHasKeyword;
        public bool ExactSearch;
        public ClaimTypeConfig IdentityClaimTypeConfig;
        public List<ClaimTypeConfig> ClaimTypeConfigList;

        public RequestInformation(IAzureCPConfiguration currentConfiguration, RequestType currentRequestType, List<ClaimTypeConfig> processedClaimTypeConfigList, string input, SPClaim incomingEntity, Uri context, AzureADObjectType[] directoryObjectTypes, string hierarchyNodeID, int maxCount)
        {
            //this.CurrentConfiguration = currentConfiguration;
            this.RequestType = currentRequestType;
            this.OriginalInput = input;
            this.IncomingEntity = incomingEntity;
            this.Context = context;
            //this.EntityTypes = entityTypes;
            this.DirectoryObjectTypes = directoryObjectTypes;
            this.HierarchyNodeID = hierarchyNodeID;
            this.MaxCount = maxCount;

            HttpContext httpctx = HttpContext.Current;
            if (httpctx != null)
            {
                WIF.ClaimsPrincipal cp = httpctx.User as WIF.ClaimsPrincipal;
                // cp is typically null in central administration
                if (cp != null) this.UserInHttpContext = SPClaimProviderManager.Local.DecodeClaimFromFormsSuffix(cp.Identity.Name);
            }

            if (currentRequestType == RequestType.Validation)
            {
                this.InitializeValidation(processedClaimTypeConfigList);
            }
            else if (currentRequestType == RequestType.Search)
            {
                this.InitializeSearch(processedClaimTypeConfigList, currentConfiguration.FilterExactMatchOnly);
            }
            else if (currentRequestType == RequestType.Augmentation)
            {
                this.InitializeAugmentation(processedClaimTypeConfigList);
            }
        }

        /// <summary>
        /// Validation is when SharePoint asks LDAPCP to return 1 PickerEntity from a given SPClaim
        /// </summary>
        /// <param name="processedClaimTypeConfigList"></param>
        protected void InitializeValidation(List<ClaimTypeConfig> processedClaimTypeConfigList)
        {
            if (this.IncomingEntity == null) throw new ArgumentNullException("claimToValidate");
            this.IdentityClaimTypeConfig = FindClaimsSetting(processedClaimTypeConfigList, this.IncomingEntity.ClaimType);
            if (this.IdentityClaimTypeConfig == null) return;
            //this.ClaimTypeConfigList = new List<ClaimTypeConfig>() { this.IdentityClaimTypeConfig };
            this.ClaimTypeConfigList = processedClaimTypeConfigList.Where(x =>
                String.Equals(x.ClaimType, this.IncomingEntity.ClaimType, StringComparison.InvariantCultureIgnoreCase)
                && !x.UseMainClaimTypeOfDirectoryObject).ToList();
            this.ExactSearch = true;
            this.Input = this.IncomingEntity.Value;
        }

        /// <summary>
        /// Search is when SharePoint asks LDAPCP to return all PickerEntity that match input provided
        /// </summary>
        /// <param name="processedClaimTypeConfigList"></param>
        protected void InitializeSearch(List<ClaimTypeConfig> processedClaimTypeConfigList, bool exactSearch)
        {
            this.ExactSearch = exactSearch;
            this.Input = this.OriginalInput;
            if (!String.IsNullOrEmpty(this.HierarchyNodeID))
            {
                // Restrict search to attributes currently selected in the hierarchy (may return multiple results if identity claim type)
                ClaimTypeConfigList = processedClaimTypeConfigList.FindAll(x =>
                    String.Equals(x.ClaimType, this.HierarchyNodeID, StringComparison.InvariantCultureIgnoreCase) &&
                    //this.EntityTypes.Contains(x.ClaimEntityType));
                    this.DirectoryObjectTypes.Contains(x.DirectoryObjectType));
            }
            else
            {
                // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                //ClaimTypeConfigList = processedClaimTypeConfigList.FindAll(x => this.EntityTypes.Contains(x.ClaimEntityType));
                ClaimTypeConfigList = processedClaimTypeConfigList.FindAll(x => this.DirectoryObjectTypes.Contains(x.DirectoryObjectType));
            }
        }

        protected void InitializeAugmentation(List<ClaimTypeConfig> processedClaimTypeConfigList)
        {
            if (this.IncomingEntity == null) throw new ArgumentNullException("claimToValidate");
            this.IdentityClaimTypeConfig = FindClaimsSetting(processedClaimTypeConfigList, this.IncomingEntity.ClaimType);
            if (this.IdentityClaimTypeConfig == null) return;
        }

        public static ClaimTypeConfig FindClaimsSetting(List<ClaimTypeConfig> processedClaimTypeConfigList, string claimType)
        {
            var claimsSettings = processedClaimTypeConfigList.FindAll(x =>
                String.Equals(x.ClaimType, claimType, StringComparison.InvariantCultureIgnoreCase)
                && !x.UseMainClaimTypeOfDirectoryObject);
            if (claimsSettings.Count != 1)
            {
                // Should always find only 1 attribute at this stage
                AzureCPLogging.Log(String.Format("[{0}] Found {1} attributes that match the claim type \"{2}\", but exactly 1 is expected. Verify that there is no duplicate claim type. Aborting operation.", AzureCP._ProviderInternalName, claimsSettings.Count.ToString(), claimType), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Claims_Picking);
                return null;
            }
            return claimsSettings.First();
        }
    }

    public enum AzureADObjectProperty
    {
        // Values are aligned with enum type Microsoft.Azure.ActiveDirectory.GraphClient.GraphProperty in Microsoft.Azure.ActiveDirectory.GraphClient.dll
        None = 0,
        AccountEnabled = 1,
        Id = 2,
        Department = 20,
        DisplayName = 28,
        GivenName = 32,
        JobTitle = 37,
        Mail = 41,
        MobilePhone = 47,
        OfficeLocation = 54,
        Surname = 83,
        UserPrincipalName = 93,
        UserType = 94
    }

    public enum AzureADObjectType
    {
        User,
        Group
    }

    //public class GraphPropertyQuery
    //{
    //    public GraphProperty PropertyName;
    //    //public string SearchQuery = "startswith({0},'{1}')";
    //    //public string ValidationQuery = "{0} eq '{1}'";
    //    public Type FieldType = typeof(String);

    //    public GraphPropertyQuery(GraphProperty PropertyName)
    //    {
    //        this.PropertyName = PropertyName;
    //    }
    //}

    public enum RequestType
    {
        Search,
        Validation,
        Augmentation,
    }
}
