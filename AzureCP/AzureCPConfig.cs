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
//using WIF3_5 = Microsoft.IdentityModel.Claims;

namespace azurecp
{
    public interface IAzureCPConfiguration
    {
        List<AzureTenant> AzureTenants { get; set; }
        List<AzureADObject> AzureADObjects { get; set; }
        bool AlwaysResolveUserInput { get; set; }
        bool FilterExactMatchOnly { get; set; }
        bool AugmentAADRoles { get; set; }
    }

    public class Constants
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

        public List<AzureADObject> AzureADObjects
        {
            get { return AzureADObjectsPersisted; }
            set { AzureADObjectsPersisted = value; }
        }
        [Persisted]
        private List<AzureADObject> AzureADObjectsPersisted;

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

        public bool AugmentAADRoles
        {
            get { return AugmentAADRolesPersisted; }
            set { AugmentAADRolesPersisted = value; }
        }
        [Persisted]
        private bool AugmentAADRolesPersisted = true;

        public AzureCPConfig(string persistedObjectName, SPPersistedObject parent) : base(persistedObjectName, parent)
        { }

        public AzureCPConfig()
        { }

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
                AzureCPLogging.Log(String.Format("Error while retrieving SPPersistedObject {0}: {1}", Constants.AZURECPCONFIG_NAME, ex.Message), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
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

        public static AzureCPConfig ResetPersistedObject(string persistedObjectName)
        {
            AzureCPConfig persistedObject = GetConfiguration(persistedObjectName);
            if (persistedObject != null)
            {
                AzureCPConfig newPersistedObject = GetDefaultConfiguration(persistedObjectName);
                newPersistedObject.Update();

                AzureCPLogging.Log($"Claims list of PersistedObject {persistedObjectName} was successfully reset to default relationship table",
                    TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            }
            return null;
        }

        public static void ResetClaimsList(string persistedObjectName)
        {
            AzureCPConfig persistedObject = GetConfiguration(persistedObjectName);
            if (persistedObject != null)
            {
                persistedObject.AzureADObjects.Clear();
                persistedObject.AzureADObjects = GetDefaultAADClaimTypeList();
                persistedObject.Update();

                AzureCPLogging.Log(
                    String.Format("Claims list of PersistedObject {0} was successfully reset to default relationship table", Constants.AZURECPCONFIG_NAME),
                    TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            }
            return;
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

            AzureCPLogging.Log($"Creating persisted object {persistedObjectName} with ID {persistedObjectID}...", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
            AzureCPConfig PersistedObject = new AzureCPConfig(persistedObjectName, SPFarm.Local);
            PersistedObject.Id = new Guid(persistedObjectID);
            PersistedObject.AzureTenants = new List<AzureTenant>();
            PersistedObject = GetDefaultConfiguration(persistedObjectName);
            PersistedObject.Update();
            AzureCPLogging.Log($"Created PersistedObject {PersistedObject.Name} with Id {PersistedObject.Id}",
                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
            return PersistedObject;
        }

        public static AzureCPConfig GetDefaultConfiguration(string persistedObjectName)
        {
            AzureCPConfig persistedObject = new AzureCPConfig(persistedObjectName, SPFarm.Local);
            persistedObject.AzureTenants = new List<AzureTenant>();
            persistedObject.AzureADObjects = GetDefaultAADClaimTypeList();
            return persistedObject;
        }

        public static List<AzureADObject> GetDefaultAADClaimTypeList()
        {
            return new List<AzureADObject>
            {
                // By default ACS issues those 3 claim types: ClaimTypes.Name ClaimTypes.GivenName and ClaimTypes.Surname.
                // But ClaimTypes.Name (http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name) is a reserved claim type in SharePoint that cannot be used in the SPTrust.
                //new AzureADObject{GraphProperty=GraphProperty.UserPrincipalName, ClaimType=WIF.ClaimTypes.Name, ClaimEntityType=SPClaimEntityTypes.User},//yvand@TENANTNAME.onmicrosoft.com

                // Alternatives claim types to ClaimTypes.Name that might be used as identity claim types:
                new AzureADObject{GraphProperty=GraphProperty.UserPrincipalName, ClaimType=WIF.ClaimTypes.Upn, ClaimEntityType=SPClaimEntityTypes.User},
                new AzureADObject{GraphProperty=GraphProperty.UserPrincipalName, ClaimType=WIF.ClaimTypes.Email, ClaimEntityType=SPClaimEntityTypes.User},

                // Additional properties to find user
                new AzureADObject{GraphProperty=GraphProperty.DisplayName, CreateAsIdentityClaim=true, ClaimEntityType=SPClaimEntityTypes.User, EntityDataKey=PeopleEditorEntityDataKeys.DisplayName},
                new AzureADObject{GraphProperty=GraphProperty.GivenName, CreateAsIdentityClaim=true, ClaimEntityType=SPClaimEntityTypes.User},//Yvan
                new AzureADObject{GraphProperty=GraphProperty.Surname, CreateAsIdentityClaim=true, ClaimEntityType=SPClaimEntityTypes.User},//Duhamel

                // Retrieve additional properties to populate metadata in SharePoint (no claim type and CreateAsIdentityClaim = false)
                new AzureADObject{GraphProperty=GraphProperty.Mail, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.Email},
                new AzureADObject{GraphProperty=GraphProperty.MobilePhone, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.MobilePhone},
                new AzureADObject{GraphProperty=GraphProperty.JobTitle, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.JobTitle},
                new AzureADObject{GraphProperty=GraphProperty.Department, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.Department},
                new AzureADObject{GraphProperty=GraphProperty.OfficeLocation, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.Location},

                // Role
                //new AzureADObject{GraphProperty=GraphProperty.DisplayName, ClaimType=WIF.ClaimTypes.Role, ClaimEntityType=SPClaimEntityTypes.FormsRole},
                new AzureADObject{GraphProperty=GraphProperty.Id, ClaimType=WIF.ClaimTypes.Role, ClaimEntityType=SPClaimEntityTypes.FormsRole, GraphPropertyToDisplay=GraphProperty.DisplayName},
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

    /// <summary>
    /// Stores configuration associated to a claim type, and its mapping with the Azure AD attribute (GraphProperty)
    /// </summary>
    public class AzureADObject : SPAutoSerializingObject
    {
        public string ClaimType
        {
            get { return ClaimTypePersisted; }
            set { ClaimTypePersisted = value; }
        }
        [Persisted]
        private string ClaimTypePersisted;

        /// <summary>
        /// Azure AD attribute mapped to the claim type
        /// </summary>
        public GraphProperty GraphProperty
        {
            get { return (GraphProperty)Enum.ToObject(typeof(GraphProperty), GraphPropertyPersisted); }
            set { GraphPropertyPersisted = (int)value; }
        }
        [Persisted]
        private int GraphPropertyPersisted;


        /// <summary>
        /// Microsoft.SharePoint.Administration.Claims.SPClaimEntityTypes
        /// </summary>
        public string ClaimEntityType
        {
            get { return ClaimEntityTypePersisted; }
            set { ClaimEntityTypePersisted = value; }
        }
        [Persisted]
        private string ClaimEntityTypePersisted = SPClaimEntityTypes.User;

        /// <summary>
        /// Can contain a member of class PeopleEditorEntityDataKey http://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.webcontrols.peopleeditorentitydatakeys_members(v=office.15).aspx
        /// to populate additional metadata in permission created
        /// </summary>
        public string EntityDataKey
        {
            get { return EntityDataKeyPersisted; }
            set { EntityDataKeyPersisted = value; }
        }
        [Persisted]
        private string EntityDataKeyPersisted;

        /// <summary>
        /// Every claim value type is String by default
        /// </summary>
        public string ClaimValueType
        {
            get { return ClaimValueTypePersisted; }
            set { ClaimValueTypePersisted = value; }
        }
        [Persisted]
        private string ClaimValueTypePersisted = WIF.ClaimValueTypes.String;

        /// <summary>
        /// If set to true, property ClaimType should not be set
        /// </summary>
        public bool CreateAsIdentityClaim
        {
            get { return CreateAsIdentityClaimPersisted; }
            set { CreateAsIdentityClaimPersisted = value; }
        }
        [Persisted]
        private bool CreateAsIdentityClaimPersisted = false;

        /// <summary>
        /// If set, its value can be used as a prefix in the people picker to create a permission without actually quyerying Azure AD
        /// </summary>
        public string PrefixToBypassLookup
        {
            get { return PrefixToBypassLookupPersisted; }
            set { PrefixToBypassLookupPersisted = value; }
        }
        [Persisted]
        private string PrefixToBypassLookupPersisted;

        public GraphProperty GraphPropertyToDisplay
        {
            get { return (GraphProperty)Enum.ToObject(typeof(GraphProperty), GraphPropertyToDisplayPersisted); }
            set { GraphPropertyToDisplayPersisted = (int)value; }
        }
        [Persisted]
        private int GraphPropertyToDisplayPersisted;

        /// <summary>
        /// Set to only return values that exactly match the input
        /// </summary>
        public bool FilterExactMatchOnly
        {
            get { return FilterExactMatchOnlyPersisted; }
            set { FilterExactMatchOnlyPersisted = value; }
        }
        [Persisted]
        private bool FilterExactMatchOnlyPersisted = false;

        /// <summary>
        /// This azureObject is not intended to be used or modified in your code
        /// </summary>
        public string ClaimTypeMappingName
        {
            get { return ClaimTypeMappingNamePersisted; }
            set { ClaimTypeMappingNamePersisted = value; }
        }
        [Persisted]
        private string ClaimTypeMappingNamePersisted;

        /// <summary>
        /// This azureObject is not intended to be used or modified in your code
        /// </summary>
        public string PeoplePickerAttributeHierarchyNodeId
        {
            get { return PeoplePickerAttributeHierarchyNodeIdPersisted; }
            set { PeoplePickerAttributeHierarchyNodeIdPersisted = value; }
        }
        [Persisted]
        private string PeoplePickerAttributeHierarchyNodeIdPersisted;

        internal AzureADObject CopyPersistedProperties()
        {
            AzureADObject copy = new AzureADObject()
            {
                ClaimTypePersisted = this.ClaimTypePersisted,
                GraphPropertyPersisted = this.GraphPropertyPersisted,
                ClaimEntityTypePersisted = this.ClaimEntityTypePersisted,
                EntityDataKeyPersisted = this.EntityDataKeyPersisted,
                ClaimValueTypePersisted = this.ClaimValueTypePersisted,
                CreateAsIdentityClaimPersisted = this.CreateAsIdentityClaimPersisted,
                PrefixToBypassLookupPersisted = this.PrefixToBypassLookupPersisted,
                GraphPropertyToDisplayPersisted = this.GraphPropertyToDisplayPersisted,
                FilterExactMatchOnlyPersisted = this.FilterExactMatchOnlyPersisted,
                ClaimTypeMappingNamePersisted = this.ClaimTypeMappingNamePersisted,
                PeoplePickerAttributeHierarchyNodeIdPersisted = this.PeoplePickerAttributeHierarchyNodeIdPersisted,
            };
            return copy;
        }
    }

    public class AzureTenant : SPAutoSerializingObject
    {
        [Persisted]
        public Guid Id = Guid.NewGuid();

        [Persisted]
        public string TenantName;

        [Persisted]
        public string TenantId;

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
                TenantId = this.TenantId,
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
        public string[] EntityTypes;
        private string OriginalInput;
        public string HierarchyNodeID;
        public int MaxCount;

        public string Input;
        public bool InputHasKeyword;
        public bool ExactSearch;
        public AzureADObject IdentityClaimTypeConfig;
        public List<AzureADObject> ClaimTypeConfigList;

        public RequestInformation(IAzureCPConfiguration currentConfiguration, RequestType currentRequestType, List<AzureADObject> processedClaimTypeConfigList, string input, SPClaim incomingEntity, Uri context, string[] entityTypes, string hierarchyNodeID, int maxCount)
        {
            //this.CurrentConfiguration = currentConfiguration;
            this.RequestType = currentRequestType;
            this.OriginalInput = input;
            this.IncomingEntity = incomingEntity;
            this.Context = context;
            this.EntityTypes = entityTypes;
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
        protected void InitializeValidation(List<AzureADObject> processedClaimTypeConfigList)
        {
            if (this.IncomingEntity == null) throw new ArgumentNullException("claimToValidate");
            this.IdentityClaimTypeConfig = FindClaimsSetting(processedClaimTypeConfigList, this.IncomingEntity.ClaimType);
            if (this.IdentityClaimTypeConfig == null) return;
            this.ClaimTypeConfigList = new List<AzureADObject>() { this.IdentityClaimTypeConfig };
            this.ExactSearch = true;
            this.Input = this.IncomingEntity.Value;
        }

        /// <summary>
        /// Search is when SharePoint asks LDAPCP to return all PickerEntity that match input provided
        /// </summary>
        /// <param name="processedClaimTypeConfigList"></param>
        protected void InitializeSearch(List<AzureADObject> processedClaimTypeConfigList, bool exactSearch)
        {
            this.ExactSearch = exactSearch;
            this.Input = this.OriginalInput;
            if (!String.IsNullOrEmpty(this.HierarchyNodeID))
            {
                // Restrict search to attributes currently selected in the hierarchy (may return multiple results if identity claim type)
                ClaimTypeConfigList = processedClaimTypeConfigList.FindAll(x =>
                    String.Equals(x.ClaimType, this.HierarchyNodeID, StringComparison.InvariantCultureIgnoreCase) &&
                    this.EntityTypes.Contains(x.ClaimEntityType));
            }
            else
            {
                // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                ClaimTypeConfigList = processedClaimTypeConfigList.FindAll(x => this.EntityTypes.Contains(x.ClaimEntityType));
            }
        }

        protected void InitializeAugmentation(List<AzureADObject> processedClaimTypeConfigList)
        {
            if (this.IncomingEntity == null) throw new ArgumentNullException("claimToValidate");
            this.IdentityClaimTypeConfig = FindClaimsSetting(processedClaimTypeConfigList, this.IncomingEntity.ClaimType);
            if (this.IdentityClaimTypeConfig == null) return;
        }

        public static AzureADObject FindClaimsSetting(List<AzureADObject> processedClaimTypeConfigList, string claimType)
        {
            var claimsSettings = processedClaimTypeConfigList.FindAll(x =>
                String.Equals(x.ClaimType, claimType, StringComparison.InvariantCultureIgnoreCase)
                && !x.CreateAsIdentityClaim);
            if (claimsSettings.Count != 1)
            {
                // Should always find only 1 attribute at this stage
                AzureCPLogging.Log(String.Format("[{0}] Found {1} attributes that match the claim type \"{2}\", but exactly 1 is expected. Verify that there is no duplicate claim type. Aborting operation.", AzureCP._ProviderInternalName, claimsSettings.Count.ToString(), claimType), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Claims_Picking);
                return null;
            }
            return claimsSettings.First();
        }
    }

    public enum GraphProperty
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
