﻿using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Web;
using Yvand.EntraClaimsProvider.Logging;
using Logger = Yvand.EntraClaimsProvider.Logging.Logger;
using WIF4_5 = System.Security.Claims;

namespace Yvand.EntraClaimsProvider.Configuration
{
    public static class ClaimsProviderConstants
    {
        public static string CONFIGURATION_ID => "1DB24D85-B10E-4D06-B378-8CFD05A129F3";
        public static string CONFIGURATION_NAME => "EntraCPConfig";
        public static List<AzureCloudProperties> AzureClouds => new List<AzureCloudProperties>{
            new AzureCloudProperties
            {
                Name = AzureCloudName.AzureGlobal,
                NameInGraphCore = GraphClientFactory.Global_Cloud,
                Authority = AzureAuthorityHosts.AzurePublicCloud.ToString(),
                GraphScope = "https://graph.microsoft.com/.default",
                GraphUrl = "https://graph.microsoft.com",
            },
            new AzureCloudProperties
            {
                Name = AzureCloudName.AzureUsGovernmentL4,
                NameInGraphCore = GraphClientFactory.USGOV_Cloud,
                Authority = AzureAuthorityHosts.AzureGovernment.ToString(),
                GraphScope = "https://graph.microsoft.us/.default",
                GraphUrl = "https://graph.microsoft.us",
            },
            new AzureCloudProperties
            {
                Name = AzureCloudName.AzureUsGovernmentL5,
                NameInGraphCore = GraphClientFactory.USGOV_DOD_Cloud,
                Authority = AzureAuthorityHosts.AzureGovernment.ToString(),
                GraphScope = "https://dod-graph.microsoft.us/.default",
                GraphUrl = "https://dod-graph.microsoft.us",
            },
            new AzureCloudProperties
            {
                Name = AzureCloudName.AzureChina,
                NameInGraphCore = GraphClientFactory.China_Cloud,
                Authority = AzureAuthorityHosts.AzureChina.ToString(),
                GraphScope = "https://microsoftgraph.chinacloudapi.cn/.default",
                GraphUrl = "https://microsoftgraph.chinacloudapi.cn",
            },
        };
        public static string GroupClaimEntityType { get; set; } = SPClaimEntityTypes.FormsRole;
        public static bool EnforceOnly1ClaimTypeForGroup => true;     // In EntraCP, only 1 claim type can be used to create group permissions
        public static string DefaultMainGroupClaimType => WIF4_5.ClaimTypes.Role;
        public static string PUBLICSITEURL => "https://entracp.yvand.net/";
        public static string GUEST_USERTYPE => "Guest";
        public static string MEMBER_USERTYPE => "Member";
        public static string ClientCertificatePrivateKeyPassword => "YVANDwRrEHVHQ57ge?uda";
        private static object Lock_SetClaimsProviderVersion = new object();
        private static string _ClaimsProviderVersion;
        public static string ClaimsProviderVersion
        {
            get
            {
                if (!String.IsNullOrEmpty(_ClaimsProviderVersion))
                {
                    return _ClaimsProviderVersion;
                }

                // Method FileVersionInfo.GetVersionInfo() may hang and block all LDAPCP threads, so it is read only 1 time
                lock (Lock_SetClaimsProviderVersion)
                {
                    if (!String.IsNullOrEmpty(_ClaimsProviderVersion))
                    {
                        return _ClaimsProviderVersion;
                    }

                    try
                    {
                        _ClaimsProviderVersion = FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(EntraCP)).Location).FileVersion;
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
        public static int DEFAULT_TIMEOUT = 10 * 1000;
#else
        public static int DEFAULT_TIMEOUT = 15 * 1000;
#endif

        public static readonly int DefaultTenantDataCacheLifetimeInMinutes = 15;
    }

    public enum AzureCloudName
    {
        AzureGlobal,
        AzureChina,
        AzureUsGovernmentL4,
        AzureUsGovernmentL5,
    }

    public class AzureCloudProperties
    {
        public AzureCloudName Name;
        public string NameInGraphCore;
        public string Authority;
        public string GraphScope;
        public string GraphUrl;
    }

    public enum DirectoryObjectProperty
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

    public enum OperationType
    {
        Search,
        Validation,
        Augmentation,
    }

    /// <summary>
    /// Contains information about current operation
    /// </summary>
    public class OperationContext
    {
        /// <summary>
        /// Indicates what kind of operation SharePoint is requesting
        /// </summary>
        public OperationType OperationType { get; private set; }

        /// <summary>
        /// Set only if request is a validation or an augmentation, to the incoming entity provided by SharePoint
        /// </summary>
        public SPClaim IncomingEntity { get; private set; }

        /// <summary>
        /// User submitting the query in the poeple picker, retrieved from HttpContext. Can be null
        /// </summary>
        public SPClaim UserInHttpContext { get; private set; }

        /// <summary>
        /// Uri provided by SharePoint
        /// </summary>
        public Uri UriContext { get; private set; }

        /// <summary>
        /// EntityTypes expected by SharePoint in the entities returned
        /// </summary>
        public DirectoryObjectType[] DirectoryObjectTypes { get; private set; }

        public string HierarchyNodeID { get; private set; }
        public int MaxCount { get; }

        /// <summary>
        /// If request is a validation: contains the value of the SPClaim. If request is a search: contains the input
        /// </summary>
        public string Input { get; private set; }

        public bool InputHasKeyword { get; private set; }

        /// <summary>
        /// Indicates if search operation should return only results that exactly match the Input
        /// </summary>
        public bool ExactSearch { get; private set; }

        /// <summary>
        /// Contains the relevant list of ClaimTypeConfig for every type of request. In case of validation or augmentation, it will contain only 1 item.
        /// </summary>
        public List<ClaimTypeConfig> CurrentClaimTypeConfigList { get; private set; }

        /// <summary>
        /// List of the Azure tenants registered, in an object that is unique to this thread, so none of its changes will be shared with other threads
        /// </summary>
        public List<EntraIDTenant> AzureTenantsCopy { get; private set; }

        public OperationContext(ClaimsProviderSettings settings, OperationType currentRequestType, string input, SPClaim incomingEntity, Uri context, string[] entityTypes, string hierarchyNodeID, int maxCount)
        {
            this.OperationType = currentRequestType;
            this.Input = input;
            this.IncomingEntity = incomingEntity;
            this.UriContext = context;
            this.HierarchyNodeID = hierarchyNodeID;
            this.MaxCount = maxCount;

            // settings.EntraIDTenants must be cloned locally to ensure its properties ($select / $filter) won't be updated by multiple threads
            this.AzureTenantsCopy = new List<EntraIDTenant>(settings.EntraIDTenants.Count);
            foreach (EntraIDTenant tenant in settings.EntraIDTenants)
            {
                EntraIDTenant copy = new EntraIDTenant();
                Utils.CopyPublicProperties(typeof(EntraIDTenant), tenant, copy);
                AzureTenantsCopy.Add(copy);
            }

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
                if (cp != null && cp.Identity != null)
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
                this.InitializeValidation(settings.RuntimeClaimTypesList);
            }
            else if (currentRequestType == OperationType.Search)
            {
                this.InitializeSearch(settings.RuntimeClaimTypesList, settings.FilterExactMatchOnly);
            }
            else if (currentRequestType == OperationType.Augmentation)
            {
                this.InitializeAugmentation(settings.RuntimeClaimTypesList);
            }
        }

        /// <summary>
        /// Validation is when SharePoint expects exactly 1 PickerEntity from the incoming SPClaim
        /// </summary>
        /// <param name="runtimeClaimTypesList"></param>
        protected void InitializeValidation(List<ClaimTypeConfig> runtimeClaimTypesList)
        {
            if (this.IncomingEntity == null) { throw new ArgumentNullException(nameof(this.IncomingEntity)); }

            // FirstOrDefault returns null if no result, while First throws an exception
            ClaimTypeConfig incomingEntityClaimTypeConfig = runtimeClaimTypesList.FirstOrDefault(x =>
               String.Equals(x.ClaimType, this.IncomingEntity.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
               !x.UseMainClaimTypeOfDirectoryObject);
            if (incomingEntityClaimTypeConfig == null)
            {
                Logger.Log($"[{EntraCP.ClaimsProviderName}] Unable to validate entity \"{this.IncomingEntity.Value}\" because its claim type \"{this.IncomingEntity.ClaimType}\" was not found in the ClaimTypes list of current configuration.", TraceSeverity.Unexpected, TraceCategory.Configuration);
                throw new InvalidOperationException($"[{EntraCP.ClaimsProviderName}] Unable validate entity \"{this.IncomingEntity.Value}\" because its claim type \"{this.IncomingEntity.ClaimType}\" was not found in the ClaimTypes list of current configuration.");
            }
            this.CurrentClaimTypeConfigList = new List<ClaimTypeConfig>(1)
            {
                incomingEntityClaimTypeConfig,
            };
            this.ExactSearch = true;
        }

        /// <summary>
        /// Search is when SharePoint expects a list of any PickerEntity that match input provided
        /// </summary>
        /// <param name="runtimeClaimTypesList"></param>
        protected void InitializeSearch(List<ClaimTypeConfig> runtimeClaimTypesList, bool exactSearch)
        {
            this.ExactSearch = exactSearch;
            if (!String.IsNullOrEmpty(this.HierarchyNodeID))
            {
                // Restrict search to ClaimType currently selected in the hierarchy (may return multiple results if identity claim type)
                CurrentClaimTypeConfigList = runtimeClaimTypesList.FindAll(x =>
                    String.Equals(x.ClaimType, this.HierarchyNodeID, StringComparison.InvariantCultureIgnoreCase) &&
                    this.DirectoryObjectTypes.Contains(x.EntityType));
            }
            else
            {
                // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                CurrentClaimTypeConfigList = runtimeClaimTypesList.FindAll(x => this.DirectoryObjectTypes.Contains(x.EntityType));
            }
        }

        protected void InitializeAugmentation(List<ClaimTypeConfig> runtimeClaimTypesList)
        {
        }
    }
}
