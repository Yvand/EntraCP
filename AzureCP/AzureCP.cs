using Microsoft.Graph;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using Nito.AsyncEx;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using WIF = System.Security.Claims;

/*
 * DO NOT directly edit AzureCP class. It is designed to be inherited to customize it as desired.
 * Please download "AzureCP for Developers.zip" on https://github.com/Yvand/AzureCP to find examples and guidance.
 * */

namespace azurecp
{
    /// <summary>
    /// Provides search and resolution against Azure Active Directory
    /// Visit https://github.com/Yvand/AzureCP for documentation and updates.
    /// Please report any bug to https://github.com/Yvand/AzureCP.
    /// Author: Yvan Duhamel
    /// </summary>
    public class AzureCP : SPClaimProvider
    {
        public const string _ProviderInternalName = "AzureCP";
        public virtual string ProviderInternalName { get { return "AzureCP"; } }

        private object Sync_Init = new object();
        private ReaderWriterLockSlim Lock_Config = new ReaderWriterLockSlim();
        private long AzureCPConfigVersion = 0;

        /// <summary>
        /// Async lock to use AAD client context in only 1 thread at a time
        /// </summary>
        private AsyncLock AADContextLock = new AsyncLock();

        /// <summary>
        /// Contains configuration currently used by claims provider
        /// </summary>
        public IAzureCPConfiguration CurrentConfiguration;

        /// <summary>
        /// SPTrust associated with the claims provider
        /// </summary>
        protected SPTrustedLoginProvider SPTrust;

        /// <summary>
        /// object mapped to the identity claim in the SPTrustedIdentityTokenIssuer
        /// </summary>
        AzureADObject IdentityClaimTypeConfig;

        /// <summary>
        /// Processed list to use. It is guarranted to never contain an empty ClaimType
        /// </summary>
        public List<AzureADObject> ProcessedClaimTypeConfigList;
        public List<AzureADObject> UserMetadataClaimTypeConfigList;
        protected virtual string PickerEntityDisplayText { get { return "({0}) {1}"; } }
        protected virtual string PickerEntityOnMouseOver { get { return "{0}={1}"; } }

        protected string IssuerName
        {
            get
            {
                // The advantage of using the SPTrustedLoginProvider name for the issuer name is that it makes possible and easy to replace current claims provider with another one.
                // The other claims provider would simply have to use SPTrustedLoginProvider name too
                return SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, SPTrust.Name);
            }
        }

        public AzureCP(string displayName) : base(displayName)
        {
        }

        /// <summary>
        /// Initializes claim provider. This method is reserved for internal use and is not intended to be called from external code or changed
        /// </summary>
        public bool Initialize(Uri context, string[] entityTypes)
        {
            // Ensures thread safety to initialize class variables
            lock (Sync_Init)
            {
                // 1ST PART: GET CONFIGURATION OBJECT
                AzureCPConfig globalConfiguration = null;
                bool refreshConfig = false;
                bool success = true;
                bool initializeFromPersistedObject = true;
                try
                {
                    if (SPTrust == null)
                    {
                        SPTrust = GetSPTrustAssociatedWithCP(ProviderInternalName);
                        if (SPTrust == null) return false;
                    }
                    if (!CheckIfShouldProcessInput(context)) return false;

                    // Should not try to get PersistedObject if not OOB AzureCP since with current design it works correctly only for OOB AzureCP
                    if (String.Equals(ProviderInternalName, AzureCP._ProviderInternalName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        globalConfiguration = AzureCPConfig.GetFromConfigDB();
                        if (globalConfiguration == null)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject not found. Visit AzureCP admin pages in central administration to create it.", ProviderInternalName),
                                TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            // Cannot continue since it's not inherited and no persisted object exists
                            success = false;
                        }
                        else if (globalConfiguration.AzureADObjects == null || globalConfiguration.AzureADObjects.Count == 0)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject was found but there are no AzureADObject set. Visit AzureCP admin pages in central administration to create it.", ProviderInternalName),
                                TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            // Cannot continue 
                            success = false;
                        }
                        else if (globalConfiguration.AzureTenants == null || globalConfiguration.AzureTenants.Count == 0)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject was found but there are no Azure tenant set. Visit AzureCP admin pages in central administration to add one.", ProviderInternalName),
                                TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            // Cannot continue 
                            success = false;
                        }
                        else
                        {
                            // Persisted object is found and seems valid
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject found, version: {1}, previous version: {2}", ProviderInternalName, globalConfiguration.Version.ToString(), this.AzureCPConfigVersion.ToString()),
                                TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);
                            if (this.AzureCPConfigVersion != globalConfiguration.Version)
                            {
                                refreshConfig = true;
                                this.AzureCPConfigVersion = globalConfiguration.Version;
                                AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject changed, refreshing configuration", ProviderInternalName),
                                    TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Core);
                            }
                        }
                    }
                    else
                    {
                        // AzureCP class inherited, refresh config
                        // Configuration will be retrieved in SetCustomSettings method
                        initializeFromPersistedObject = false;
                        refreshConfig = true;
                        AzureCPLogging.Log(String.Format("[{0}] AzureCP class inherited", ProviderInternalName),
                            TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Core);
                    }
                }
                catch (Exception ex)
                {
                    success = false;
                    AzureCPLogging.LogException(ProviderInternalName, "in Initialize", AzureCPLogging.Categories.Core, ex);
                }
                finally
                { }

                if (!success) return success;
                if (!refreshConfig) return success;

                // 2ND PART: APPLY CONFIGURATION
                // Configuration needs to be refreshed, lock current thread in write mode
                Lock_Config.EnterWriteLock();
                try
                {
                    AzureCPLogging.Log(String.Format("[{0}] Refreshing configuration", ProviderInternalName),
                        TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Core);

                    // Create local persisted object that will never be saved in config DB, it's just a local copy
                    this.CurrentConfiguration = new AzureCPConfig();
                    if (initializeFromPersistedObject)
                    {
                        // All settings come from persisted object
                        this.CurrentConfiguration.AlwaysResolveUserInput = globalConfiguration.AlwaysResolveUserInput;
                        this.CurrentConfiguration.FilterExactMatchOnly = globalConfiguration.FilterExactMatchOnly;
                        this.CurrentConfiguration.AugmentAADRoles = globalConfiguration.AugmentAADRoles;

                        // Retrieve AzureADObjects
                        // A copy of collection AzureADObjects must be created because SetActualAADObjectCollection() may change it and it should be made in a copy totally independant from the persisted object
                        this.CurrentConfiguration.AzureADObjects = new List<AzureADObject>();
                        foreach (AzureADObject currentObject in globalConfiguration.AzureADObjects)
                        {
                            // Create a new AzureADObject
                            this.CurrentConfiguration.AzureADObjects.Add(currentObject.CopyPersistedProperties());
                        }

                        // Retrieve AzureTenants
                        // Create a copy of the collection to work in an copy separated from persisted object
                        this.CurrentConfiguration.AzureTenants = new List<AzureTenant>();
                        foreach (AzureTenant currentObject in globalConfiguration.AzureTenants)
                        {
                            // Create a copy from persisted object
                            this.CurrentConfiguration.AzureTenants.Add(currentObject.CopyPersistedProperties());
                        }
                    }
                    else
                    {
                        // All settings come from overriden SetCustomConfiguration method
                        SetCustomConfiguration(context, entityTypes);

                        // Ensure we get what we expect
                        if (this.CurrentConfiguration.AzureADObjects == null || this.CurrentConfiguration.AzureADObjects.Count == 0)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureADObjects was not set. Override method SetCustomConfiguration to set it.", ProviderInternalName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            return false;
                        }

                        if (this.CurrentConfiguration.AzureTenants == null || this.CurrentConfiguration.AzureTenants.Count == 0)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureTenants was not set. Override method SetCustomConfiguration to set it.", ProviderInternalName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            return false;
                        }
                    }

                    // Set properties AuthenticationProvider and GraphService
                    foreach (var coco in this.CurrentConfiguration.AzureTenants)
                    {
                        coco.SetAzureADContext();
                    }
                    success = this.InitializeClaimTypeConfigList(this.CurrentConfiguration.AzureADObjects);
                }
                catch (Exception ex)
                {
                    success = false;
                    AzureCPLogging.LogException(ProviderInternalName, "in Initialize, while refreshing configuration", AzureCPLogging.Categories.Core, ex);
                }
                finally
                {
                    Lock_Config.ExitWriteLock();
                }
                return success;
            }
        }

        /// <summary>
        /// Initializes claim provider. This method is reserved for internal use and is not intended to be called from external code or changed
        /// </summary>
        /// <param name="nonProcessedClaimTypeConfigList"></param>
        /// <returns></returns>
        private bool InitializeClaimTypeConfigList(List<AzureADObject> nonProcessedClaimTypeConfigList)
        {
            bool success = true;
            try
            {
                bool identityClaimTypeFound = false;
                // Get attributes defined in trust based on their claim type (unique way to map them)
                List<AzureADObject> claimTypesSetInTrust = new List<AzureADObject>();
                // There is a bug in the SharePoint API: SPTrustedLoginProvider.ClaimTypes should retrieve SPTrustedClaimTypeInformation.MappedClaimType, but it returns SPTrustedClaimTypeInformation.InputClaimType instead, so we cannot rely on it
                //foreach (var attr in _AttributesDefinitionList.Where(x => AssociatedSPTrustedLoginProvider.ClaimTypes.Contains(x.claimType)))
                //{
                //    attributesDefinedInTrust.Add(attr);
                //}
                foreach (SPTrustedClaimTypeInformation claimTypeInformation in SPTrust.ClaimTypeInformation)
                {
                    // Search if current claim type in trust exists in AzureADObjects
                    // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                    List<AzureADObject> mappedClaimTypeConfigList = nonProcessedClaimTypeConfigList.FindAll(x =>
                        String.Equals(x.ClaimType, claimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        !x.CreateAsIdentityClaim &&
                        x.GraphProperty != GraphProperty.None);
                    AzureADObject claimTypeSettings;
                    if (mappedClaimTypeConfigList.Count == 1)
                    {
                        claimTypeSettings = mappedClaimTypeConfigList.First();
                        claimTypesSetInTrust.Add(claimTypeSettings);

                        if (String.Equals(SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimTypeSettings.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                        {
                            // Identity claim type found, set IdentityAzureADObject property
                            identityClaimTypeFound = true;
                            IdentityClaimTypeConfig = claimTypeSettings;
                        }
                    }
                }

                // Check if identity claim is there. Should always check property SPTrustedClaimTypeInformation.MappedClaimType: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.administration.claims.sptrustedclaimtypeinformation.mappedclaimtype.aspx
                if (!identityClaimTypeFound)
                {
                    AzureCPLogging.Log(String.Format("[{0}] Impossible to continue because identity claim type \"{1}\" set in the SPTrustedIdentityTokenIssuer \"{2}\" is missing in AzureADObjects.", ProviderInternalName, SPTrust.IdentityClaimTypeInformation.MappedClaimType, SPTrust.Name), TraceSeverity.Unexpected, EventSeverity.ErrorCritical, AzureCPLogging.Categories.Core);
                    return false;
                }

                // This check is to find if there is a duplicate of the identity claim type that uses the same GraphProperty
                //AzureADObject objectToDelete = claimTypesSetInTrust.Find(x =>
                //    !String.Equals(x.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                //    !x.CreateAsIdentityClaim &&
                //    x.GraphProperty == GraphProperty.UserPrincipalName);
                //if (objectToDelete != null) claimTypesSetInTrust.Remove(objectToDelete);

                // Check if there are objects that should be always queried (CreateAsIdentityClaim) to add in the list
                List<AzureADObject> additionalClaimTypeConfigList = new List<AzureADObject>();
                foreach (AzureADObject claimTypeConfig in nonProcessedClaimTypeConfigList.Where(x => x.CreateAsIdentityClaim))// && !claimTypesSetInTrust.Contains(x, new LDAPPropertiesComparer())))
                {
                    // Check if identity claim type is already using same GraphProperty, and ignore current object if so
                    if (IdentityClaimTypeConfig.GraphProperty == claimTypeConfig.GraphProperty) continue;

                    // Normally ClaimType should be null if CreateAsIdentityClaim is set to true, but we check here it and handle this scenario
                    if (!String.IsNullOrEmpty(claimTypeConfig.ClaimType))
                    {
                        if (String.Equals(SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimTypeConfig.ClaimType))
                        {
                            // Not a big deal since it's set with identity claim type, so no inconsistent behavior to expect, just record an information
                            AzureCPLogging.Log(String.Format("[{0}] Object with GraphProperty {1} is set with CreateAsIdentityClaim to true and ClaimType {2}. Remove ClaimType property as it is useless.", ProviderInternalName, claimTypeConfig.GraphProperty, claimTypeConfig.ClaimType), TraceSeverity.Monitorable, EventSeverity.Information, AzureCPLogging.Categories.Core);
                        }
                        else if (claimTypesSetInTrust.Count(x => String.Equals(x.ClaimType, claimTypeConfig.ClaimType)) > 0)
                        {
                            // Same claim type already exists with CreateAsIdentityClaim == false. 
                            // Current object is a bad one and shouldn't be added. Don't add it but continue to build objects list
                            AzureCPLogging.Log(String.Format("[{0}] Claim type {1} is defined twice with CreateAsIdentityClaim set to true and false, which is invalid. Remove entry with CreateAsIdentityClaim set to true.", ProviderInternalName, claimTypeConfig.ClaimType), TraceSeverity.Monitorable, EventSeverity.Information, AzureCPLogging.Categories.Core);
                            continue;
                        }
                    }

                    claimTypeConfig.ClaimType = SPTrust.IdentityClaimTypeInformation.MappedClaimType;    // Give those objects the identity claim type
                    claimTypeConfig.ClaimEntityType = SPClaimEntityTypes.User;
                    claimTypeConfig.GraphPropertyToDisplay = IdentityClaimTypeConfig.GraphPropertyToDisplay; // Must be set otherwise display text of permissions will be inconsistent
                    additionalClaimTypeConfigList.Add(claimTypeConfig);
                }

                ProcessedClaimTypeConfigList = new List<AzureADObject>(claimTypesSetInTrust.Count + additionalClaimTypeConfigList.Count);
                ProcessedClaimTypeConfigList.AddRange(claimTypesSetInTrust);
                ProcessedClaimTypeConfigList.AddRange(additionalClaimTypeConfigList);

                // Parse objects to configure some settings
                // An object can have ClaimType set to null if only used to populate metadata of permission created
                foreach (var attr in ProcessedClaimTypeConfigList.Where(x => x.ClaimType != null))
                {
                    var trustedClaim = SPTrust.GetClaimTypeInformationFromMappedClaimType(attr.ClaimType);
                    // It should never be null
                    if (trustedClaim == null) continue;
                    attr.ClaimTypeMappingName = trustedClaim.DisplayName;
                }

                // Any metadata for a user with GraphProperty actually set is valid
                this.UserMetadataClaimTypeConfigList = nonProcessedClaimTypeConfigList.FindAll(x =>
                    !String.IsNullOrEmpty(x.EntityDataKey) &&
                    x.GraphProperty != GraphProperty.None &&
                    x.ClaimEntityType == SPClaimEntityTypes.User);
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in InitializeClaimTypeConfigList", AzureCPLogging.Categories.Core, ex);
                success = false;
            }
            return success;
        }

        /// <summary>
        /// Override this method to customize configuration of AzureCP
        /// </summary> 
        /// <param name="context">The context, as a URI</param>
        /// <param name="entityTypes">The EntityType entity types set to scope the search to</param>
        protected virtual void SetCustomConfiguration(Uri context, string[] entityTypes)
        {
        }

        /// <summary>
        /// Check if AzureCP should process input (and show results) based on current URL (context)
        /// </summary>
        /// <param name="context">The context, as a URI</param>
        /// <returns></returns>
        protected virtual bool CheckIfShouldProcessInput(Uri context)
        {
            if (context == null) return true;
            var webApp = SPWebApplication.Lookup(context);
            if (webApp == null) return false;
            if (webApp.IsAdministrationWebApplication) return true;

            // Not central admin web app, enable AzureCP only if current web app uses it
            // It is not possible to exclude zones where AzureCP is not used because:
            // Consider following scenario: default zone is WinClaims, intranet zone is Federated:
            // In intranet zone, when creating permission, AzureCP will be called 2 times. The 2nd time (in FillResolve (SPClaim)), the context will always be the URL of the default zone
            foreach (var zone in Enum.GetValues(typeof(SPUrlZone)))
            {
                SPIisSettings iisSettings = webApp.GetIisSettingsWithFallback((SPUrlZone)zone);
                if (!iisSettings.UseTrustedClaimsAuthenticationProvider)
                    continue;

                // Get the list of authentication providers associated with the zone
                foreach (SPAuthenticationProvider prov in iisSettings.ClaimsAuthenticationProviders)
                {
                    if (prov.GetType() == typeof(Microsoft.SharePoint.Administration.SPTrustedAuthenticationProvider))
                    {
                        // Check if the current SPTrustedAuthenticationProvider is associated with the claim provider
                        if (String.Equals(prov.ClaimProviderName, ProviderInternalName, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Get the first TrustedLoginProvider associated with current claim provider
        /// LIMITATION: The same claims provider (uniquely identified by its name) cannot be associated to multiple TrustedLoginProvider because at runtime there is no way to determine what TrustedLoginProvider is currently calling
        /// </summary>
        /// <param name="providerInternalName"></param>
        /// <returns></returns>
        public static SPTrustedLoginProvider GetSPTrustAssociatedWithCP(string providerInternalName)
        {
            var lp = SPSecurityTokenServiceManager.Local.TrustedLoginProviders.Where(x => String.Equals(x.ClaimProviderName, providerInternalName, StringComparison.OrdinalIgnoreCase));

            if (lp != null && lp.Count() == 1)
                return lp.First();

            if (lp != null && lp.Count() > 1)
                AzureCPLogging.Log(String.Format("[{0}] Claims provider {0} is associated to multiple SPTrustedIdentityTokenIssuer, which is not supported because at runtime there is no way to determine what TrustedLoginProvider is currently calling", providerInternalName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);

            AzureCPLogging.Log(String.Format("[{0}] Claims provider {0} is not associated with any SPTrustedIdentityTokenIssuer so it cannot create permissions.\r\nVisit http://ldapcp.codeplex.com for installation procedure or set property ClaimProviderName with PowerShell cmdlet Get-SPTrustedIdentityTokenIssuer to create association.", providerInternalName), TraceSeverity.High, EventSeverity.Warning, AzureCPLogging.Categories.Core);
            return null;
        }

        /// <summary>
        /// Returns the graph property value of a GraphObject (User, Group, Role)
        /// </summary>
        /// <param name="src"></param>
        /// <param name="propName"></param>
        /// <returns>Null if property doesn't exist. String.Empty if property exists but has no value. Actual value otherwise</returns>
        public static string GetGraphPropertyValue(object src, string propName)
        {
            PropertyInfo pi = src.GetType().GetProperty(propName);
            if (pi == null) return null;    // Property doesn't exist
            object propertyValue = pi.GetValue(src, null);
            return propertyValue == null ? String.Empty : propertyValue.ToString();
        }

        /// <summary>
        /// Create the SPClaim with proper issuer name
        /// </summary>
        /// <param name="type">Claim type</param>
        /// <param name="value">Claim value</param>
        /// <param name="valueType">Claim valueType</param>
        /// <param name="inputHasKeyword">Did the original input contain a keyword?</param>
        /// <returns></returns>
        protected virtual new SPClaim CreateClaim(string type, string value, string valueType)
        {
            string claimValue = String.Empty;
            var obj = ProcessedClaimTypeConfigList.FirstOrDefault(x => String.Equals(x.ClaimType, type, StringComparison.InvariantCultureIgnoreCase));
            claimValue = value;
            // SPClaimProvider.CreateClaim issues with SPOriginalIssuerType.ClaimProvider
            //return CreateClaim(type, claimValue, valueType);
            return new SPClaim(type, claimValue, valueType, IssuerName);
        }

        protected virtual PickerEntity CreatePickerEntityHelper(AzureCPResult result)
        {
            PickerEntity pe = CreatePickerEntity();
            SPClaim claim;
            string permissionValue = result.PermissionValue;
            string permissionClaimType = result.ClaimTypeConfig.ClaimType;
            bool isIdentityClaimType = false;

            if (String.Equals(result.ClaimTypeConfig.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase)
                || result.ClaimTypeConfig.CreateAsIdentityClaim)
            {
                isIdentityClaimType = true;
            }

            if (result.ClaimTypeConfig.CreateAsIdentityClaim)
            {
                // This azureObject is not directly linked to a claim type, so permission is created with identity claim type
                permissionClaimType = IdentityClaimTypeConfig.ClaimType;
                permissionValue = FormatPermissionValue(permissionClaimType, permissionValue, isIdentityClaimType, result);
                claim = CreateClaim(
                    permissionClaimType,
                    permissionValue,
                    IdentityClaimTypeConfig.ClaimValueType);
                pe.EntityType = IdentityClaimTypeConfig.ClaimEntityType;
            }
            else
            {
                permissionValue = FormatPermissionValue(permissionClaimType, permissionValue, isIdentityClaimType, result);
                claim = CreateClaim(
                    permissionClaimType,
                    permissionValue,
                    result.ClaimTypeConfig.ClaimValueType);
                pe.EntityType = result.ClaimTypeConfig.ClaimEntityType;
            }

            pe.DisplayText = FormatPermissionDisplayText(permissionClaimType, permissionValue, isIdentityClaimType, result);
            pe.Description = String.Format(
                PickerEntityOnMouseOver,
                result.ClaimTypeConfig.GraphProperty.ToString(),
                result.QueryMatchValue);
            pe.Claim = claim;
            pe.IsResolved = true;
            //pe.EntityGroupName = "";

            int nbMetadata = 0;
            // Populate metadata attributes of permission created
            foreach (var entityAttrib in UserMetadataClaimTypeConfigList)
            {
                // if there is actally a value in the GraphObject, then it can be set
                string entityAttribValue = GetGraphPropertyValue(result.UserOrGroupResult, entityAttrib.GraphProperty.ToString());
                if (!String.IsNullOrEmpty(entityAttribValue))
                {
                    pe.EntityData[entityAttrib.EntityDataKey] = entityAttribValue;
                    nbMetadata++;
                    AzureCPLogging.Log(String.Format("[{0}] Added metadata \"{1}\" with value \"{2}\" to permission", ProviderInternalName, entityAttrib.EntityDataKey, entityAttribValue), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                }
            }

            AzureCPLogging.Log(String.Format("[{0}] Created permission: display text: \"{1}\", value: \"{2}\", claim type: \"{3}\", and filled with {4} metadata.", ProviderInternalName, pe.DisplayText, pe.Claim.Value, pe.Claim.ClaimType, nbMetadata.ToString()), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
            return pe;
        }

        /// <summary>
        /// Override this method to customize value of permission created
        /// </summary>
        /// <param name="claimType"></param>
        /// <param name="claimValue"></param>
        /// <param name="isIdentityClaimType"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        protected virtual string FormatPermissionValue(string claimType, string claimValue, bool isIdentityClaimType, AzureCPResult result)
        {
            return claimValue;
        }

        /// <summary>
        /// Override this method to customize display text of permission created
        /// </summary>
        /// <param name="displayText"></param>
        /// <param name="claimType"></param>
        /// <param name="claimValue"></param>
        /// <param name="isIdentityClaim"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        protected virtual string FormatPermissionDisplayText(string claimType, string claimValue, bool isIdentityClaimType, AzureCPResult result)
        {
            string permissionDisplayText = String.Empty;
            string valueDisplayedInPermission = String.Empty;

            if (result.ClaimTypeConfig.GraphPropertyToDisplay != GraphProperty.None)
            {
                if (!isIdentityClaimType) permissionDisplayText = "(" + result.ClaimTypeConfig.ClaimTypeMappingName + ") ";

                string graphPropertyToDisplayValue = GetGraphPropertyValue(result.UserOrGroupResult, result.ClaimTypeConfig.GraphPropertyToDisplay.ToString());
                if (!String.IsNullOrEmpty(graphPropertyToDisplayValue)) permissionDisplayText += graphPropertyToDisplayValue;
                else permissionDisplayText += result.PermissionValue;

            }
            else
            {
                if (isIdentityClaimType)
                {
                    permissionDisplayText = result.QueryMatchValue;
                }
                else
                {
                    permissionDisplayText = String.Format(
                        PickerEntityDisplayText,
                        result.ClaimTypeConfig.ClaimTypeMappingName,
                        result.PermissionValue);
                }
            }

            return permissionDisplayText;
        }

        protected virtual PickerEntity CreatePickerEntityForSpecificClaimType(string input, AzureADObject claimTypesToResolve, bool inputHasKeyword)
        {
            List<PickerEntity> entities = CreatePickerEntityForSpecificClaimTypes(
                input,
                new List<AzureADObject>()
                    {
                        claimTypesToResolve,
                    },
                inputHasKeyword);
            return entities == null ? null : entities.First();
        }

        protected virtual List<PickerEntity> CreatePickerEntityForSpecificClaimTypes(string input, List<AzureADObject> claimTypesToResolve, bool inputHasKeyword)
        {
            List<PickerEntity> entities = new List<PickerEntity>();
            foreach (var claimTypeToResolve in claimTypesToResolve)
            {
                PickerEntity pe = CreatePickerEntity();
                SPClaim claim = CreateClaim(claimTypeToResolve.ClaimType, input, claimTypeToResolve.ClaimValueType);

                if (String.Equals(claim.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    pe.DisplayText = input;
                }
                else
                {
                    pe.DisplayText = String.Format(
                        PickerEntityDisplayText,
                        claimTypeToResolve.ClaimTypeMappingName,
                        input);
                }

                pe.EntityType = claimTypeToResolve.ClaimEntityType;
                pe.Description = String.Format(
                    PickerEntityOnMouseOver,
                    claimTypeToResolve.GraphProperty.ToString(),
                    input);

                pe.Claim = claim;
                pe.IsResolved = true;
                //pe.EntityGroupName = "";

                if (claimTypeToResolve.ClaimEntityType == SPClaimEntityTypes.User && !String.IsNullOrEmpty(claimTypeToResolve.EntityDataKey))
                {
                    pe.EntityData[claimTypeToResolve.EntityDataKey] = pe.Claim.Value;
                    AzureCPLogging.Log(String.Format("[{0}] Added metadata \"{1}\" with value \"{2}\" to permission", ProviderInternalName, claimTypeToResolve.EntityDataKey, pe.EntityData[claimTypeToResolve.EntityDataKey]), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                }
                entities.Add(pe);
                AzureCPLogging.Log(String.Format("[{0}] Created permission: display text: \"{1}\", value: \"{2}\", claim type: \"{3}\".", ProviderInternalName, pe.DisplayText, pe.Claim.Value, pe.Claim.ClaimType), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
            }
            return entities.Count > 0 ? entities : null;
        }

        /// <summary>
        /// Called when claims provider is added to the farm. At this point the persisted object is not created yet so we can't pass actual claim type list
        /// If assemblyBinding for Newtonsoft.Json was not correctly added on the server, this method will generate an assembly load exception during feature activation
        /// Also called every 1st query in people picker
        /// </summary>
        /// <param name="claimTypes"></param>
        protected override void FillClaimTypes(List<string> claimTypes)
        {
            if (claimTypes == null) return;
            try
            {
                this.Lock_Config.EnterReadLock();
                if (ProcessedClaimTypeConfigList == null) return;
                foreach (var claimTypeSettings in ProcessedClaimTypeConfigList)
                {
                    claimTypes.Add(claimTypeSettings.ClaimType);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillClaimTypes", AzureCPLogging.Categories.Core, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        protected override void FillClaimValueTypes(List<string> claimValueTypes)
        {
            claimValueTypes.Add(WIF.ClaimValueTypes.String);
        }

        protected override void FillClaimsForEntity(Uri context, SPClaim entity, SPClaimProviderContext claimProviderContext, List<SPClaim> claims)
        {
            AugmentEntity(context, entity, claimProviderContext, claims);
        }

        protected override void FillClaimsForEntity(Uri context, SPClaim entity, List<SPClaim> claims)
        {
            AugmentEntity(context, entity, null, claims);
        }

        /// <summary>
        /// Perform augmentation of entity supplied
        /// </summary>
        /// <param name="context"></param>
        /// <param name="entity">entity to augment</param>
        /// <param name="claimProviderContext">Can be null</param>
        /// <param name="claims"></param>
        protected virtual void AugmentEntity(Uri context, SPClaim entity, SPClaimProviderContext claimProviderContext, List<SPClaim> claims)
        {
            Stopwatch timer = new Stopwatch();
            timer.Start();
            SPClaim decodedEntity;
            if (SPClaimProviderManager.IsUserIdentifierClaim(entity))
                decodedEntity = SPClaimProviderManager.DecodeUserIdentifierClaim(entity);
            else
            {
                if (SPClaimProviderManager.IsEncodedClaim(entity.Value))
                    decodedEntity = SPClaimProviderManager.Local.DecodeClaim(entity.Value);
                else
                    decodedEntity = entity;
            }

            SPOriginalIssuerType loginType = SPOriginalIssuers.GetIssuerType(decodedEntity.OriginalIssuer);
            if (loginType != SPOriginalIssuerType.TrustedProvider && loginType != SPOriginalIssuerType.ClaimProvider)
            {
                AzureCPLogging.Log($"[{ProviderInternalName}] Not trying to augment '{decodedEntity.Value}' because OriginalIssuer is '{decodedEntity.OriginalIssuer}'.",
                    TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Augmentation);
                return;
            }

            if (!Initialize(context, null))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                if (!this.CurrentConfiguration.AugmentAADRoles)
                    return;

                AzureCPLogging.Log($"[{ProviderInternalName}] Starting augmentation for user '{decodedEntity.Value}'.", TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Augmentation);
                AzureADObject groupClaimTypeSettings = this.ProcessedClaimTypeConfigList.FirstOrDefault(x => x.ClaimEntityType == SPClaimEntityTypes.FormsRole);
                if (groupClaimTypeSettings == null)
                {
                    AzureCPLogging.Log($"[{ProviderInternalName}] No role claim type with SPClaimEntityTypes set to 'FormsRole' was found, please check claims mapping table.",
                        TraceSeverity.High, EventSeverity.Error, AzureCPLogging.Categories.Augmentation);
                    return;
                }

                RequestInformation infos = new RequestInformation(CurrentConfiguration, RequestType.Augmentation, ProcessedClaimTypeConfigList, null, decodedEntity, context, null, null, Int32.MaxValue);
                Task<List<SPClaim>> resultsTask = GetGroupMembershipAsync(infos, groupClaimTypeSettings);
                resultsTask.Wait();
                List<SPClaim> groups = resultsTask.Result;
                timer.Stop();
                if (groups?.Count > 0)
                {
                    foreach (SPClaim group in groups)
                    {
                        claims.Add(group);
                        AzureCPLogging.Log($"[{ProviderInternalName}] Added group '{group.Value}' to user '{infos.IncomingEntity.Value}'",
                            TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Augmentation);
                    }
                    AzureCPLogging.Log($"[{ProviderInternalName}] User '{infos.IncomingEntity.Value}' was augmented with {groups.Count.ToString()} groups in {timer.ElapsedMilliseconds.ToString()} ms",
                        TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Augmentation);
                }
                else
                {
                    AzureCPLogging.Log($"[{ProviderInternalName}] No group found for user '{infos.IncomingEntity.Value}', search took {timer.ElapsedMilliseconds.ToString()} ms",
                        TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Augmentation);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in AugmentEntity", AzureCPLogging.Categories.Augmentation, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }

        }

        protected async virtual Task<List<SPClaim>> GetGroupMembershipAsync(RequestInformation requestInfo, AzureADObject groupClaimTypeSettings)
        {
            List<SPClaim> claims = new List<SPClaim>();
            foreach (var tenant in this.CurrentConfiguration.AzureTenants)
            {
                // The logic is that there will always be only 1 tenant returning groups, so as soon as 1 returned groups, foreach can stop
                claims = await GetGroupMembershipFromAzureADAsync(requestInfo, groupClaimTypeSettings, tenant).ConfigureAwait(false);
                if (claims?.Count > 0) break;
            }
            return claims;
        }

        protected async virtual Task<List<SPClaim>> GetGroupMembershipFromAzureADAsync(RequestInformation requestInfo, AzureADObject groupClaimTypeSettings, AzureTenant tenant)
        {
            List<SPClaim> claims = new List<SPClaim>();
            var userResult = await tenant.GraphService.Users.Request().Filter($"{requestInfo.IdentityClaimTypeConfig.GraphProperty} eq '{requestInfo.IncomingEntity.Value}'").GetAsync().ConfigureAwait(false);
            User user = userResult.FirstOrDefault();
            if (user == null) return claims;
            // This only returns a collection of strings, set with group ID:
            //IDirectoryObjectGetMemberGroupsCollectionPage groups = await tenant.GraphService.Users[requestInfo.IncomingEntity.Value].GetMemberGroups(true).Request().PostAsync().ConfigureAwait(false);
            IUserMemberOfCollectionWithReferencesPage groups = await tenant.GraphService.Users[requestInfo.IncomingEntity.Value].MemberOf.Request().GetAsync().ConfigureAwait(false);
            bool continueProcess = groups?.Count > 0;
            while (continueProcess)
            {
                foreach (Group group in groups.OfType<Group>())
                {
                    string groupClaimValue = GetGraphPropertyValue(group, groupClaimTypeSettings.GraphProperty.ToString());
                    claims.Add(CreateClaim(groupClaimTypeSettings.ClaimType, groupClaimValue, groupClaimTypeSettings.ClaimValueType));
                }
                if (groups.NextPageRequest != null) groups = await groups.NextPageRequest.GetAsync().ConfigureAwait(false);
                else continueProcess = false;
            }

            //if (groups?.Count > 0)
            //{
            //    do
            //    {
            //        foreach (Group group in groups.OfType<Group>())
            //        {
            //            string groupClaimValue = GetGraphPropertyValue(group, groupAttribute.GraphProperty.ToString());
            //            claims.Add(CreateClaim(groupAttribute.ClaimType, groupClaimValue, groupAttribute.ClaimValueType));
            //        }
            //        if (groups.NextPageRequest != null) groups = await groups.NextPageRequest.GetAsync().ConfigureAwait(false);
            //    }
            //    while (groups?.Count > 0 && groups.NextPageRequest != null);
            //}
            return claims;
        }

        protected override void FillEntityTypes(List<string> entityTypes)
        {
            entityTypes.Add(SPClaimEntityTypes.User);
            entityTypes.Add(SPClaimEntityTypes.FormsRole);
        }

        protected override void FillHierarchy(Uri context, string[] entityTypes, string hierarchyNodeID, int numberOfLevels, Microsoft.SharePoint.WebControls.SPProviderHierarchyTree hierarchy)
        {
            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                if (hierarchyNodeID == null)
                {
                    // Root level
                    //foreach (var azureObject in FinalAttributeList.Where(x => !String.IsNullOrEmpty(x.peoplePickerAttributeHierarchyNodeId) && !x.CreateAsIdentityClaim && entityTypes.Contains(x.ClaimEntityType)))
                    foreach (var azureObject in this.ProcessedClaimTypeConfigList.FindAll(x => !x.CreateAsIdentityClaim && entityTypes.Contains(x.ClaimEntityType)))
                    {
                        hierarchy.AddChild(
                            new Microsoft.SharePoint.WebControls.SPProviderHierarchyNode(
                                _ProviderInternalName,
                                azureObject.ClaimTypeMappingName,
                                azureObject.ClaimType,
                                true));
                    }
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillHierarchy", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        /// <summary>
        /// Override this method to change / remove permissions created by LDAPCP, or add new ones
        /// </summary>
        /// <param name="context"></param>
        /// <param name="entityTypes"></param>
        /// <param name="input"></param>
        /// <param name="resolved">List of permissions created by LDAPCP</param>
        protected virtual void FillPermissions(Uri context, string[] entityTypes, string input, ref List<PickerEntity> resolved)
        {
        }

        protected override void FillResolve(Uri context, string[] entityTypes, SPClaim resolveInput, List<Microsoft.SharePoint.WebControls.PickerEntity> resolved)
        {
            if (!Initialize(context, entityTypes))
                return;

            // Ensure incoming claim should be validated by AzureCP
            // Must be made after call to Initialize because SPTrustedLoginProvider name must be known
            if (!String.Equals(resolveInput.OriginalIssuer, IssuerName, StringComparison.InvariantCultureIgnoreCase))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                RequestInformation infos = new RequestInformation(CurrentConfiguration, RequestType.Validation, ProcessedClaimTypeConfigList, resolveInput.Value, resolveInput, context, entityTypes, null, Int32.MaxValue);
                List<PickerEntity> permissions = SearchOrValidate(infos);
                if (permissions.Count == 1)
                {
                    resolved.Add(permissions[0]);
                    AzureCPLogging.Log(String.Format("[{0}] Validated permission: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, permissions[0].Claim.Value, permissions[0].Claim.ClaimType),
                        TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                }
                else
                {
                    AzureCPLogging.Log(String.Format("[{0}] Validation of incoming claim returned {1} permissions instead of 1 expected. Aborting operation", ProviderInternalName, permissions.Count.ToString()), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Picking);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillResolve(SPClaim)", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        protected override void FillResolve(Uri context, string[] entityTypes, string resolveInput, List<Microsoft.SharePoint.WebControls.PickerEntity> resolved)
        {
            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                RequestInformation settings = new RequestInformation(CurrentConfiguration, RequestType.Search, ProcessedClaimTypeConfigList, resolveInput, null, context, entityTypes, null, Int32.MaxValue);
                List<PickerEntity> permissions = SearchOrValidate(settings);
                FillPermissions(context, entityTypes, resolveInput, ref permissions);
                foreach (PickerEntity permission in permissions)
                {
                    resolved.Add(permission);
                    AzureCPLogging.Log(String.Format("[{0}] Added permission: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, permission.Claim.Value, permission.Claim.ClaimType),
                        TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillResolve(string)", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        protected override void FillSchema(Microsoft.SharePoint.WebControls.SPProviderSchema schema)
        {
            schema.AddSchemaElement(new SPSchemaElement(PeopleEditorEntityDataKeys.DisplayName, "Display Name", SPSchemaElementType.Both));
        }

        protected override void FillSearch(Uri context, string[] entityTypes, string searchPattern, string hierarchyNodeID, int maxCount, Microsoft.SharePoint.WebControls.SPProviderHierarchyTree searchTree)
        {
            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                RequestInformation settings = new RequestInformation(CurrentConfiguration, RequestType.Search, ProcessedClaimTypeConfigList, searchPattern, null, context, entityTypes, hierarchyNodeID, maxCount);
                List<PickerEntity> permissions = SearchOrValidate(settings);
                FillPermissions(context, entityTypes, searchPattern, ref permissions);
                SPProviderHierarchyNode matchNode = null;
                foreach (PickerEntity permission in permissions)
                {
                    // Add current PickerEntity to the corresponding attribute in the hierarchy
                    if (searchTree.HasChild(permission.Claim.ClaimType))
                    {
                        matchNode = searchTree.Children.First(x => x.HierarchyNodeID == permission.Claim.ClaimType);
                    }
                    else
                    {
                        AzureADObject attrHelper = ProcessedClaimTypeConfigList.FirstOrDefault(x =>
                            !x.CreateAsIdentityClaim &&
                            String.Equals(x.ClaimType, permission.Claim.ClaimType, StringComparison.InvariantCultureIgnoreCase));

                        string nodeName = attrHelper != null ? attrHelper.ClaimTypeMappingName : permission.Claim.ClaimType;
                        matchNode = new SPProviderHierarchyNode(_ProviderInternalName, nodeName, permission.Claim.ClaimType, true);
                        searchTree.AddChild(matchNode);
                    }
                    matchNode.AddEntity(permission);
                    AzureCPLogging.Log(String.Format("[{0}] Added permission: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, permission.Claim.Value, permission.Claim.ClaimType),
                        TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillSearch", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        /// <summary>
        /// Search and validate requests coming from SharePoint
        /// </summary>
        /// <param name="requestInfo">Information about current context and operation</param>
        /// <returns></returns>
        protected virtual List<PickerEntity> SearchOrValidate(RequestInformation requestInfo)
        {
            List<PickerEntity> permissions = new List<PickerEntity>();
            try
            {
                if (this.CurrentConfiguration.AlwaysResolveUserInput)
                {
                    // Completely bypass LDAP lookp
                    List<PickerEntity> entities = CreatePickerEntityForSpecificClaimTypes(
                        requestInfo.Input,
                        requestInfo.ClaimTypeConfigList.FindAll(x => !x.CreateAsIdentityClaim),
                        false);
                    if (entities != null)
                    {
                        foreach (var entity in entities)
                        {
                            permissions.Add(entity);
                            AzureCPLogging.Log(String.Format("[{0}] Added permission created without LDAP lookup because LDAPCP configured to always resolve input: claim value: {1}, claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType),
                                TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                        }
                    }
                    return permissions;
                }

                if (requestInfo.RequestType == RequestType.Search)
                {
                    List<AzureADObject> attribsMatchInputPrefix = requestInfo.ClaimTypeConfigList.FindAll(x =>
                        !String.IsNullOrEmpty(x.PrefixToBypassLookup) &&
                        requestInfo.Input.StartsWith(x.PrefixToBypassLookup, StringComparison.InvariantCultureIgnoreCase));
                    if (attribsMatchInputPrefix.Count > 0)
                    {
                        // Input has a prefix, so it should be validated with no lookup
                        AzureADObject attribMatchInputPrefix = attribsMatchInputPrefix.First();
                        if (attribsMatchInputPrefix.Count > 1)
                        {
                            // Multiple attributes have same prefix, which is not allowed
                            AzureCPLogging.Log(String.Format("[{0}] Multiple attributes have same prefix ({1}), which is not allowed.", ProviderInternalName, attribMatchInputPrefix.PrefixToBypassLookup), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Picking);
                            return permissions;
                        }

                        // Check if a keyword was typed to bypass lookup and create permission manually
                        requestInfo.Input = requestInfo.Input.Substring(attribMatchInputPrefix.PrefixToBypassLookup.Length);
                        if (String.IsNullOrEmpty(requestInfo.Input)) return permissions;    // Keyword was found but nothing typed after, give up
                        PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                            requestInfo.Input,
                            attribMatchInputPrefix,
                            true);
                        if (entity != null)
                        {
                            permissions.Add(entity);
                            AzureCPLogging.Log(String.Format("[{0}] Added permission created without LDAP lookup because input matches a keyword: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType),
                                TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                            return permissions;
                        }
                    }
                    SearchOrValidateInAzureAD(requestInfo, ref permissions);
                }
                else if (requestInfo.RequestType == RequestType.Validation)
                {
                    SearchOrValidateInAzureAD(requestInfo, ref permissions);
                    if (!String.IsNullOrEmpty(requestInfo.IdentityClaimTypeConfig.PrefixToBypassLookup))
                    {
                        // At this stage, it is impossible to know if input was originally created with the keyword that bypasses LDAP lookup
                        // But it should be validated anyway since keyword is set for this claim type
                        // If previous LDAP lookup found the permission, return it as is
                        if (permissions.Count == 1) return permissions;

                        // If we don't get exactly 1 permission, create it manually
                        PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                            requestInfo.Input,
                            requestInfo.IdentityClaimTypeConfig,
                            requestInfo.InputHasKeyword);
                        if (entity != null)
                        {
                            permissions.Add(entity);
                            AzureCPLogging.Log(String.Format("[{0}] Added permission without LDAP lookup because corresponding claim type has a keyword associated. Claim value: \"{1}\", Claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType),
                                TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                        }
                        return permissions;
                    }
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in SearchOrValidate", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            return permissions;
        }

        protected virtual void SearchOrValidateInAzureAD(RequestInformation requestInfo, ref List<PickerEntity> permissions)
        {
            string userFilter = String.Empty;
            string groupFilter = String.Empty;
            string userSelect = String.Empty;
            string groupSelect = String.Empty;
            BuildFilter(requestInfo, out userFilter, out groupFilter, out userSelect, out groupSelect);

            List<AzureADResult> aadResults;
            using (new SPMonitoredScope(String.Format("[{0}] Total time spent in all LDAP server(s)", ProviderInternalName), 1000))
            {

                Task<List<AzureADResult>> taskAadResults = QueryAzureADCollectionAsync(requestInfo, userFilter, groupFilter, userSelect, groupSelect);
                taskAadResults.Wait();
                aadResults = taskAadResults.Result;
            }

            if (aadResults?.Count > 0)
            {
                List<AzureCPResult> results = ProcessAzureADResults(requestInfo, aadResults);

                if (results?.Count > 0)
                {
                    foreach (var result in results)
                    {
                        permissions.Add(result.PickerEntity);
                        AzureCPLogging.Log(String.Format("[{0}] Added permission created with LDAP lookup: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, result.PickerEntity.Claim.Value, result.PickerEntity.Claim.ClaimType),
                            TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                    }
                }
            }
        }

        /// <summary>
        /// Filter must be URL encoded as documented in https://developer.microsoft.com/en-us/graph/docs/concepts/query_parameters#encoding-query-parameters
        /// </summary>
        /// <param name="requestInfo"></param>
        /// <param name="userFilter"></param>
        /// <param name="groupFilter"></param>
        protected virtual void BuildFilter(RequestInformation requestInfo, out string userFilter, out string groupFilter, out string userSelect, out string groupSelect)
        {
            StringBuilder userFilterBuilder = new StringBuilder("accountEnabled eq true and (");
            StringBuilder groupFilterBuilder = new StringBuilder();
            StringBuilder userSelectBuilder = new StringBuilder("UserType,");
            StringBuilder groupSelectBuilder = new StringBuilder();

            string searchPattern;
            string input = requestInfo.Input;
            if (requestInfo.ExactSearch) searchPattern = "{0} eq '" + input + "'";
            else searchPattern = "startswith({0},'" + input + "')";

            bool firstUserObject = true;
            bool firstGroupObject = true;
            foreach (AzureADObject adObject in requestInfo.ClaimTypeConfigList)
            {
                string property = adObject.GraphProperty.ToString();
                string objectFilter = String.Format(searchPattern, property);
                string objectSelect = property;
                if (adObject.ClaimEntityType == SPClaimEntityTypes.User)
                {
                    if (firstUserObject) firstUserObject = false;
                    else
                    {
                        objectFilter = " or " + objectFilter;
                        objectSelect = ", " + objectSelect;
                    }
                    userFilterBuilder.Append(objectFilter);
                    userSelectBuilder.Append(objectSelect);
                }
                else
                {
                    // else with no further test assumes everything that is not a User is a Group
                    if (firstGroupObject) firstGroupObject = false;
                    else
                    {
                        objectFilter = objectFilter + " or ";
                        objectSelect = ", " + objectSelect;
                    }
                    groupFilterBuilder.Append(objectFilter);
                    groupSelectBuilder.Append(objectSelect);
                }
            }

            userFilterBuilder.Append(")");
            //groupFilterBuilder.Append(")");

            userFilter = HttpUtility.UrlEncode(userFilterBuilder.ToString());
            groupFilter = HttpUtility.UrlEncode(groupFilterBuilder.ToString());
            userSelect = HttpUtility.UrlEncode(userSelectBuilder.ToString());
            groupSelect = HttpUtility.UrlEncode(groupSelectBuilder.ToString());
        }

        protected virtual async Task<List<AzureADResult>> QueryAzureADCollectionAsync(RequestInformation requestInfo, string userFilter, string groupFilter, string userSelect, string groupSelect)
        {
            if (userFilter == null && groupFilter == null) return null;
            List<AzureADResult> allSearchResults = new List<AzureADResult>();
            var lockResults = new object();

            foreach (AzureTenant coco in this.CurrentConfiguration.AzureTenants)
            //Parallel.ForEach(this.CurrentConfiguration.AzureTenants, async coco =>
            //var queryTenantTasks = this.CurrentConfiguration.AzureTenants.Select (async coco =>
            {
                Stopwatch timer = new Stopwatch();
                AzureADResult searchResult = null;
                try
                {
                    timer.Start();
                    searchResult = await QueryAzureADAsync(requestInfo, coco, userFilter, groupFilter, userSelect, groupSelect, true).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    AzureCPLogging.LogException(ProviderInternalName, String.Format("in QueryAzureADCollectionAsync while querying tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, ex);
                }
                finally
                {
                    timer.Stop();
                }

                if (searchResult != null)
                {
                    lock (lockResults)
                    {
                        allSearchResults.Add(searchResult);
                        AzureCPLogging.Log($"[{ProviderInternalName}] Got {searchResult.UserOrGroupResultList.Count().ToString()} users/groups and {searchResult.DomainsRegisteredInAzureADTenant.Count().ToString()} registered domains in {timer.ElapsedMilliseconds.ToString()} ms from '{coco.TenantName}' with input '{requestInfo.Input}'",
                            TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                    }
                }
                else AzureCPLogging.Log($"[{ProviderInternalName}] Got no result from '{coco.TenantName}' with input '{requestInfo.Input}', search took {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                //});
            }
            return allSearchResults;
        }

        protected virtual async Task<AzureADResult> QueryAzureADAsync(RequestInformation requestInfo, AzureTenant coco, string userFilter, string groupFilter, string userSelect, string groupSelect, bool firstAttempt)
        {
            AzureCPLogging.Log($"[{ProviderInternalName}] Querying Azure AD tenant '{coco.TenantName}' for users/groups/domains, with input '{requestInfo.Input}'", TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
            AzureADResult tenantResults = new AzureADResult();
            bool tryAgain = false;
            object lockAddResultToCollection = new object();
            CancellationTokenSource cts = new CancellationTokenSource(Constants.timeout);
            try
            {
                using (new SPMonitoredScope($"[{ProviderInternalName}] Querying Azure AD tenant '{coco.TenantName}' for users/groups/domains, with input '{requestInfo.Input}'", 1000))
                {
                    using (AADContextLock.Lock())
                    {
                        Task userQueryTask = Task.Run(async () =>
                        {
                            if (userFilter == null) return;
                            AzureCPLogging.Log(String.Format("[{0}] UserQueryTask starting for tenant '{1}'", ProviderInternalName, coco.TenantName), TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                            IGraphServiceUsersCollectionPage users = await coco.GraphService.Users.Request().Select(userSelect).Filter(userFilter).GetAsync();
                            if (users?.Count > 0)
                            {
                                do
                                {
                                    lock (lockAddResultToCollection)
                                    {
                                        tenantResults.UserOrGroupResultList.AddRange(users.CurrentPage);
                                    }
                                    if (users.NextPageRequest != null) users = await users.NextPageRequest.GetAsync().ConfigureAwait(false);
                                }
                                while (users?.Count > 0 && users.NextPageRequest != null);
                            }
                            AzureCPLogging.Log(String.Format("[{0}] UserQueryTask ended for tenant '{1}'", ProviderInternalName, coco.TenantName), TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                        }, cts.Token);
                        Task groupQueryTask = Task.Run(async () =>
                        {
                            if (groupFilter == null) return;
                            AzureCPLogging.Log(String.Format("[{0}] GroupQueryTask starting for tenant '{1}'", ProviderInternalName, coco.TenantName), TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                            IGraphServiceGroupsCollectionPage groups = await coco.GraphService.Groups.Request().Select(groupSelect).Filter(groupFilter).GetAsync();
                            if (groups?.Count > 0)
                            {
                                do
                                {
                                    lock (lockAddResultToCollection)
                                    {
                                        tenantResults.UserOrGroupResultList.AddRange(groups.CurrentPage);
                                    }
                                    if (groups.NextPageRequest != null) groups = await groups.NextPageRequest.GetAsync().ConfigureAwait(false);
                                }
                                while (groups?.Count > 0 && groups.NextPageRequest != null);
                            }
                            AzureCPLogging.Log(String.Format("[{0}] GroupQueryTask ended for tenant '{1}'", ProviderInternalName, coco.TenantName), TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                        }, cts.Token);
                        Task domainQueryTask = Task.Run(async () =>
                        {
                            AzureCPLogging.Log(String.Format("[{0}] DomainQueryTask starting for tenant '{1}'", ProviderInternalName, coco.TenantName), TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                            IGraphServiceDomainsCollectionPage domains = await coco.GraphService.Domains.Request().GetAsync();
                            lock (lockAddResultToCollection)
                            {
                                tenantResults.DomainsRegisteredInAzureADTenant.AddRange(domains.Where(x => x.IsVerified == true).Select(x => x.Id));
                            }
                            AzureCPLogging.Log(String.Format("[{0}] DomainQueryTask ended for tenant '{1}'", ProviderInternalName, coco.TenantName), TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                        }, cts.Token);

                        Task.WaitAll(new Task[3] { userQueryTask, groupQueryTask, domainQueryTask }, Constants.timeout, cts.Token);
                        //await Task.WhenAll(userQueryTask, groupQueryTask).ConfigureAwait(false);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                AzureCPLogging.Log($"[{ProviderInternalName}] Query on Azure AD tenant '{coco.TenantName}' exceeded timeout of {Constants.timeout} ms and was cancelled.", TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                tryAgain = true;
            }
            catch (AggregateException ex)
            {
                // Task.WaitAll throws an AggregateException, which contains all exceptions thrown by tasks it waited on
                AzureCPLogging.LogException(ProviderInternalName, $"while querying tenant '{coco.TenantName}'", AzureCPLogging.Categories.Lookup, ex);
                tryAgain = true;
            }
            finally
            {
                //AzureCPLogging.LogDebug(String.Format("Releasing cancellation token of tenant '{0}'", coco.TenantName));
                cts.Dispose();
            }

            if (firstAttempt && tryAgain)
            {
                AzureCPLogging.Log($"[{ProviderInternalName}] Doing new attempt to query tenant '{coco.TenantName}'...",
                    TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);

                tenantResults = await QueryAzureADAsync(requestInfo, coco, userFilter, groupFilter, userSelect, groupSelect, false).ConfigureAwait(false);
            }
            return tenantResults;
        }

        protected virtual List<AzureCPResult> ProcessAzureADResults(RequestInformation requestInfo, List<AzureADResult> azureADResults)
        {
            // Split results between users/groups and list of registered domains in the tenant
            List<DirectoryObject> usersAndGroupsResults = new List<DirectoryObject>();
            List<string> domains = new List<string>();
            // For each Azure AD tenant
            foreach (AzureADResult tenantResults in azureADResults)
            {
                usersAndGroupsResults.AddRange(tenantResults.UserOrGroupResultList);
                domains.AddRange(tenantResults.DomainsRegisteredInAzureADTenant);
            }

            // Return if no user / groups is found, or if no registered domain is found
            if (usersAndGroupsResults == null || !usersAndGroupsResults.Any() || domains == null || !domains.Any())
            {
                return null;
            };

            // If exactSearch is true, we don't care about attributes with CreateAsIdentityClaim = true
            List<AzureADObject> claimTypeConfigList;
            if (requestInfo.ExactSearch) claimTypeConfigList = requestInfo.ClaimTypeConfigList.FindAll(x => !x.CreateAsIdentityClaim);
            else claimTypeConfigList = requestInfo.ClaimTypeConfigList;

            List<AzureCPResult> processedResults = new List<AzureCPResult>();
            foreach (DirectoryObject userOrGroup in usersAndGroupsResults)
            {
                DirectoryObject currentObject = null;
                string claimEntityType = null;
                if (userOrGroup is User)
                {
                    // Always skip shadow users: UserType is Guest and his mail matches a verified domain in AAD tenant
                    string userType = GetGraphPropertyValue(userOrGroup, "UserType");
                    if (String.IsNullOrEmpty(userType))
                    {
                        AzureCPLogging.Log(
                            String.Format("[{0}] User {1} filtered out because his property UserType is empty.", ProviderInternalName, ((User)userOrGroup).UserPrincipalName),
                            TraceSeverity.Unexpected, EventSeverity.Warning, AzureCPLogging.Categories.Lookup);
                        continue;
                    }
                    if (String.Equals(userType, Constants.GraphUserType.Guest, StringComparison.InvariantCultureIgnoreCase))
                    {
                        string mail = GetGraphPropertyValue(userOrGroup, "Mail");
                        if (String.IsNullOrEmpty(mail))
                        {
                            AzureCPLogging.Log(
                                String.Format("[{0}] Guest user {1} filtered out because his mail is empty.", ProviderInternalName, ((User)userOrGroup).UserPrincipalName),
                                TraceSeverity.Unexpected, EventSeverity.Warning, AzureCPLogging.Categories.Lookup);
                            continue;
                        }
                        if (!mail.Contains('@')) continue;
                        string maildomain = mail.Split('@')[1];
                        if (domains.Any(x => String.Equals(x, maildomain, StringComparison.InvariantCultureIgnoreCase)))
                        {
                            AzureCPLogging.Log(
                                String.Format("[{0}] Guest user {1} filtered out because he is in a domain registered in AAD tenant.", ProviderInternalName, mail),
                                TraceSeverity.Verbose, EventSeverity.Verbose, AzureCPLogging.Categories.Lookup);
                            continue;
                        }
                    }
                    currentObject = userOrGroup;
                    claimEntityType = SPClaimEntityTypes.User;
                }
                else
                {
                    currentObject = userOrGroup;
                    claimEntityType = SPClaimEntityTypes.FormsRole;
                }

                // Start filter
                foreach (AzureADObject claimTypeConfig in claimTypeConfigList.Where(x => x.ClaimEntityType == claimEntityType))
                {
                    // Get value with of current GraphProperty
                    string graphPropertyValue = GetGraphPropertyValue(currentObject, claimTypeConfig.GraphProperty.ToString());

                    // Check if property exists (no null) and has a value (not String.Empty)
                    if (String.IsNullOrEmpty(graphPropertyValue)) continue;

                    // Check if current value mathes input, otherwise go to next GraphProperty to check
                    if (requestInfo.ExactSearch)
                    {
                        if (!String.Equals(graphPropertyValue, requestInfo.Input, StringComparison.InvariantCultureIgnoreCase)) continue;
                    }
                    else
                    {
                        if (!graphPropertyValue.StartsWith(requestInfo.Input, StringComparison.InvariantCultureIgnoreCase)) continue;
                    }

                    // Current GraphProperty value matches user input. Add current object in search results if it passes following checks
                    string queryMatchValue = graphPropertyValue;
                    string valueToCheck = queryMatchValue;
                    // Check if current object is not already in the collection
                    AzureADObject objCompare;
                    if (claimTypeConfig.CreateAsIdentityClaim)
                    {
                        objCompare = IdentityClaimTypeConfig;
                        // Get the value of the GraphProperty linked to IdentityAzureObject
                        valueToCheck = GetGraphPropertyValue(currentObject, IdentityClaimTypeConfig.GraphProperty.ToString());
                        if (String.IsNullOrEmpty(valueToCheck)) continue;
                    }
                    else
                    {
                        objCompare = claimTypeConfig;
                    }

                    // if claim type, GraphProperty and value are identical, then result is already in collection
                    int numberResultFound = processedResults.FindAll(x =>
                        String.Equals(x.ClaimTypeConfig.ClaimType, objCompare.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        //x.AzureObject.GraphProperty == objCompare.GraphProperty &&
                        String.Equals(x.PermissionValue, valueToCheck, StringComparison.InvariantCultureIgnoreCase)).Count;
                    if (numberResultFound > 0) continue;

                    // Passed the checks, add it to the searchResults list
                    processedResults.Add(
                        new AzureCPResult(currentObject)
                        {
                            ClaimTypeConfig = claimTypeConfig,
                            //GraphPropertyValue = graphPropertyValue,
                            PermissionValue = valueToCheck,
                            QueryMatchValue = queryMatchValue,
                        });
                }
            }

            AzureCPLogging.Log(
                String.Format(
                    "[{0}] {1} permission(s) to create after filtering",
                    ProviderInternalName, processedResults.Count),
                TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
            foreach (AzureCPResult result in processedResults)
            {
                PickerEntity pe = CreatePickerEntityHelper(result);
                result.PickerEntity = pe;
            }

            return processedResults;
        }

        public override string Name { get { return ProviderInternalName; } }
        public override bool SupportsEntityInformation { get { return true; } }
        public override bool SupportsHierarchy { get { return true; } }
        public override bool SupportsResolve { get { return true; } }
        public override bool SupportsSearch { get { return true; } }
        public override bool SupportsUserKey { get { return true; } }

        /// <summary>
        /// Return the identity claim type
        /// </summary>
        /// <returns></returns>
        public override string GetClaimTypeForUserKey()
        {
            if (!Initialize(null, null))
                return null;

            this.Lock_Config.EnterReadLock();
            try
            {
                return IdentityClaimTypeConfig.ClaimType;
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in GetClaimTypeForUserKey", AzureCPLogging.Categories.Rehydration, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
            return null;
        }

        /// <summary>
        /// Return the user key (SPClaim with identity claim type) from the incoming entity
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        protected override SPClaim GetUserKeyForEntity(SPClaim entity)
        {
            if (!Initialize(null, null))
                return null;

            // There are 2 scenarios:
            // 1: OriginalIssuer is "SecurityTokenService": Value looks like "05.t|yvanhost|yvand@yvanhost.local", claim type is "http://schemas.microsoft.com/sharepoint/2009/08/claims/userid" and it must be decoded properly
            // 2: OriginalIssuer is AzureCP: in this case incoming entity is valid and returned as is
            if (String.Equals(entity.OriginalIssuer, IssuerName, StringComparison.InvariantCultureIgnoreCase))
                return entity;

            SPClaimProviderManager cpm = SPClaimProviderManager.Local;
            SPClaim curUser = SPClaimProviderManager.DecodeUserIdentifierClaim(entity);

            this.Lock_Config.EnterReadLock();
            try
            {
                AzureCPLogging.Log($"[{ProviderInternalName}] Returning user key for '{entity.Value}'",
                    TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Rehydration);
                return CreateClaim(IdentityClaimTypeConfig.ClaimType, curUser.Value, curUser.ValueType);
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in GetUserKeyForEntity", AzureCPLogging.Categories.Rehydration, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
            return null;
        }
    }

    public class AzureADResult
    {
        public List<DirectoryObject> UserOrGroupResultList;
        public List<string> DomainsRegisteredInAzureADTenant;
        //public string TenantName;

        public AzureADResult()
        {
            UserOrGroupResultList = new List<DirectoryObject>();
            DomainsRegisteredInAzureADTenant = new List<string>();
            //this.TenantName = tenantName;
        }
    }

    /// <summary>
    /// User / group found in Azure AD, with additional information
    /// </summary>
    public class AzureCPResult
    {
        public DirectoryObject UserOrGroupResult;
        public AzureADObject ClaimTypeConfig;
        public PickerEntity PickerEntity;
        public string PermissionValue;
        public string QueryMatchValue;
        //public string TenantName;

        public AzureCPResult(DirectoryObject directoryObject)
        {
            UserOrGroupResult = directoryObject;
            //TenantName = tenantName;
        }
    }
}
