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
        public const int timeout = 5000;    // 1000 secs      1000000
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

        public AzureCPConfig(SPPersistedObject parent)
            : base(Constants.AZURECPCONFIG_NAME, parent)
        {
        }

        public AzureCPConfig()
        {
        }

        protected override bool HasAdditionalUpdateAccess()
        {
            return false;
        }

        public static AzureCPConfig GetFromConfigDB()
        {
            SPPersistedObject parent = SPFarm.Local;
            try
            {
                AzureCPConfig persistedObject = parent.GetChild<AzureCPConfig>(Constants.AZURECPCONFIG_NAME);
                return persistedObject;
            }
            catch (Exception ex)
            {
                AzureCPLogging.Log(String.Format("Error while retrieving SPPersistedObject {0}: {1}", Constants.AZURECPCONFIG_NAME, ex.Message), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
            }
            return null;
        }

        public static AzureCPConfig ResetPersistedObject()
        {
            AzureCPConfig persistedObject = GetFromConfigDB();
            if (persistedObject != null)
            {
                AzureCPConfig newPersistedObject = GetDefaultSettings(persistedObject);
                newPersistedObject.Update();

                AzureCPLogging.Log(
                    String.Format("Claims list of PersistedObject {0} was successfully reset to default relationship table", Constants.AZURECPCONFIG_NAME),
                    TraceSeverity.High, EventSeverity.Information, AzureCPLogging.Categories.Core);
            }
            return null;
        }

        public static void ResetClaimsList()
        {
            AzureCPConfig persistedObject = GetFromConfigDB();
            if (persistedObject != null)
            {
                persistedObject.AzureADObjects.Clear();
                persistedObject.AzureADObjects = GetDefaultAADClaimTypeList();
                persistedObject.Update();

                AzureCPLogging.Log(
                    String.Format("Claims list of PersistedObject {0} was successfully reset to default relationship table", Constants.AZURECPCONFIG_NAME),
                    TraceSeverity.High, EventSeverity.Information, AzureCPLogging.Categories.Core);
            }
            return;
        }

        /// <summary>
        /// Create the persisted object that contains default configuration of AzureCP.
        /// It should be created only in central administration with application pool credentials
        /// because this is the only place where we are sure user has the permission to write in the config database
        /// </summary>
        public static AzureCPConfig CreatePersistedObject()
        {
            // Ensure it doesn't already exists and delete it if so
            AzureCPConfig existingConfig = AzureCPConfig.GetFromConfigDB();
            if (existingConfig != null)
            {
                DeleteAzureCPConfig();
            }

            AzureCPConfig PersistedObject = new AzureCPConfig(SPFarm.Local);
            PersistedObject.Id = new Guid(Constants.AZURECPCONFIG_ID);
            PersistedObject.AzureTenants = new List<AzureTenant>();
            PersistedObject = GetDefaultSettings(PersistedObject);
            PersistedObject.Update();
            AzureCPLogging.Log(
                String.Format("Created PersistedObject {0} with Id {1}", PersistedObject.Name, PersistedObject.Id),
                TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Core);

            return PersistedObject;
        }

        public static AzureCPConfig GetDefaultSettings(AzureCPConfig persistedObject)
        {
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
                new AzureADObject{GraphProperty=GraphProperty.Mobile, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.MobilePhone},
                new AzureADObject{GraphProperty=GraphProperty.JobTitle, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.JobTitle},
                new AzureADObject{GraphProperty=GraphProperty.Department, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.Department},
                new AzureADObject{GraphProperty=GraphProperty.PhysicalDeliveryOfficeName, ClaimEntityType="User", EntityDataKey=PeopleEditorEntityDataKeys.Location},

                // Role
                new AzureADObject{GraphProperty=GraphProperty.DisplayName, ClaimType=WIF.ClaimTypes.Role, ClaimEntityType=SPClaimEntityTypes.FormsRole},
            };
        }

        public static void DeleteAzureCPConfig()
        {
            AzureCPConfig azureCPConfig = AzureCPConfig.GetFromConfigDB();
            if (azureCPConfig != null) azureCPConfig.Delete();
        }
    }

    /// <summary>
    /// Defines an azureObject persisted in config database
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

        public GraphProperty GraphProperty
        {
            get { return (GraphProperty)Enum.ToObject(typeof(GraphProperty), GraphPropertyPersisted); }
            set { GraphPropertyPersisted = (int)value; }
        }
        [Persisted]
        private int GraphPropertyPersisted;


        /// <summary>
        /// Microsoft.SharePoint.Administration.Claims.SPClaimEntityTypes
        /// Class name in namespace Microsoft.Azure.ActiveDirectory.GraphClient that will be retrieved with reflection
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

        public string QueryPrefix
        {
            get { return QueryPrefixPersisted; }
            set { QueryPrefixPersisted = value; }
        }
        [Persisted]
        private string QueryPrefixPersisted = String.Empty;

        /// <summary>
        /// Every claim value type is a string by default
        /// </summary>
        public string ClaimValueType
        {
            get { return ClaimValueTypePersisted; }
            set { ClaimValueTypePersisted = value; }
        }
        [Persisted]
        private string ClaimValueTypePersisted = WIF.ClaimValueTypes.String;

        /// <summary>
        /// Set to true if the claim type should always be queried in LDAP even if it is not defined in the SP trust (typically displayName and cn attributes)
        /// </summary>
        public bool CreateAsIdentityClaim
        {
            get { return CreateAsIdentityClaimPersisted; }
            set { CreateAsIdentityClaimPersisted = value; }
        }
        [Persisted]
        private bool CreateAsIdentityClaimPersisted = false;

        /// <summary>
        /// Set this to tell LDAPCP to validate user input (and create the permission) without LDAP lookup if it contains this keyword at the beginning
        /// </summary>
        public string PrefixToBypassLookup
        {
            get { return PrefixToBypassLookupPersisted; }
            set { PrefixToBypassLookupPersisted = value; }
        }
        [Persisted]
        private string PrefixToBypassLookupPersisted;

        /// <summary>
        /// Set this property to customize display text of the permission with a specific LDAP azureObject (different than LDAPAttributeName, that is the actual value of the permission)
        /// </summary>
        public GraphProperty GraphPropertyToDisplay
        {
            get { return (GraphProperty)Enum.ToObject(typeof(GraphProperty), GraphPropertyToDisplayPersisted); }
            set { GraphPropertyToDisplayPersisted = (int)value; }
        }
        [Persisted]
        private int GraphPropertyToDisplayPersisted;

        /// <summary>
        /// Set to only return values that exactly match the user input
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
            };
            return copy;
        }
    }

    /// <summary>
    /// Contains information about current request
    /// </summary>
    public class RequestInformation
    {
        public static string RegexAccountFromFullAccountName { get { return ".*\\\\(.*)"; } }
        //public static string RegexDomainFromFullAccountName { get { return "(.*)\\\\(.*)"; } }
        public static string RegexDomainFromFullAccountName { get { return "(.*)\\\\.*"; } }
        public static string RegexFullDomainFromEmail { get { return ".*@(.*)"; } }

        /// <summary>
        /// Current LDAPCP configuration
        /// </summary>
        public IAzureCPConfiguration CurrentConfiguration;

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
        public AzureADObject Attribute;
        public List<AzureADObject> Attributes;

        public RequestInformation(IAzureCPConfiguration currentConfiguration, RequestType currentRequestType, List<AzureADObject> processedAttributes, string input, SPClaim incomingEntity, Uri context, string[] entityTypes, string hierarchyNodeID, int maxCount)
        {
            this.CurrentConfiguration = currentConfiguration;
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
                this.InitializeValidation(processedAttributes);
            }
            else if (currentRequestType == RequestType.Search)
            {
                this.InitializeSearch(processedAttributes);
            }
            else if (currentRequestType == RequestType.Augmentation)
            {
                this.InitializeAugmentation(processedAttributes);
            }
        }

        /// <summary>
        /// Validation is when SharePoint asks LDAPCP to return 1 PickerEntity from a given SPClaim
        /// </summary>
        /// <param name="ProcessedAttributes"></param>
        protected void InitializeValidation(List<AzureADObject> ProcessedAttributes)
        {
            if (this.IncomingEntity == null) throw new ArgumentNullException("claimToValidate");
            this.Attribute = FindAttribute(ProcessedAttributes, this.IncomingEntity.ClaimType);
            if (this.Attribute == null) return;
            this.Attributes = new List<AzureADObject>() { this.Attribute };
            this.ExactSearch = true;
            this.Input = this.IncomingEntity.Value;
        }

        /// <summary>
        /// Search is when SharePoint asks LDAPCP to return all PickerEntity that match input provided
        /// </summary>
        /// <param name="ProcessedAttributes"></param>
        protected void InitializeSearch(List<AzureADObject> ProcessedAttributes)
        {
            this.ExactSearch = this.CurrentConfiguration.FilterExactMatchOnly;
            this.Input = this.OriginalInput;
            if (!String.IsNullOrEmpty(this.HierarchyNodeID))
            {
                // Restrict search to attributes currently selected in the hierarchy (may return multiple results if identity claim type)
                Attributes = ProcessedAttributes.FindAll(x =>
                    String.Equals(x.ClaimType, this.HierarchyNodeID, StringComparison.InvariantCultureIgnoreCase) &&
                    this.EntityTypes.Contains(x.ClaimEntityType));
            }
            else
            {
                // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                Attributes = ProcessedAttributes.FindAll(x => this.EntityTypes.Contains(x.ClaimEntityType));
            }
        }

        protected void InitializeAugmentation(List<AzureADObject> ProcessedAttributes)
        {
            if (this.IncomingEntity == null) throw new ArgumentNullException("claimToValidate");
            this.Attribute = FindAttribute(ProcessedAttributes, this.IncomingEntity.ClaimType);
            if (this.Attribute == null) return;
        }

        public static AzureADObject FindAttribute(List<AzureADObject> processedAttributes, string claimType)
        {
            var Attributes = processedAttributes.FindAll(x =>
                String.Equals(x.ClaimType, claimType, StringComparison.InvariantCultureIgnoreCase)
                && !x.CreateAsIdentityClaim);
            if (Attributes.Count != 1)
            {
                // Should always find only 1 attribute at this stage
                AzureCPLogging.Log(String.Format("[{0}] Found {1} attributes that match the claim type \"{2}\", but exactly 1 is expected. Verify that there is no duplicate claim type. Aborting operation.", AzureCP._ProviderInternalName, Attributes.Count.ToString(), claimType), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Picking);
                return null;
            }
            return Attributes.First();
        }

        public static string GetAccountFromFullAccountName(string fullAccountName)
        {
            return Regex.Replace(fullAccountName, RegexAccountFromFullAccountName, "$1", RegexOptions.None);
        }

        /// <summary>
        /// Returns the string before the '\'
        /// </summary>
        /// <param name="fullAccountName">e.g. "mylds.local\ldsgroup1"</param>
        /// <returns>e.g. "mylds.local"</returns>
        public static string GetDomainFromFullAccountName(string fullAccountName)
        {
            return Regex.Replace(fullAccountName, RegexDomainFromFullAccountName, "$1", RegexOptions.None);
        }

        public static string GetFQDNFromEmail(string email)
        {
            return Regex.Replace(email, RegexFullDomainFromEmail, "$1", RegexOptions.None);
        }

        public static string GetFirstSubString(string value, string separator)
        {
            int stop = value.IndexOf(separator);
            return (stop > -1) ? value.Substring(0, stop) : string.Empty;
        }        
    }

    public enum GraphProperty
    {
        AccountEnabled,
        DisplayName,
        Mail,
        PreferredName,
        UserPrincipalName,
        UserType
    }

    public enum RequestType
    {
        Search,
        Validation,
        Augmentation,
    }
}
