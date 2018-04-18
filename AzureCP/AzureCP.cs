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
using static azurecp.AzureCPLogging;
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
        public virtual string ProviderInternalName => "AzureCP";
        public virtual string PersistedObjectName => ClaimsProviderConstants.AZURECPCONFIG_NAME;

        private object Sync_Init = new object();
        private ReaderWriterLockSlim Lock_Config = new ReaderWriterLockSlim();
        private long AzureCPConfigVersion = 0;

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
        ClaimTypeConfig IdentityClaimTypeConfig;

        /// <summary>
        /// Processed list to use. It is guarranted to never contain an empty ClaimType
        /// </summary>
        public List<ClaimTypeConfig> ProcessedClaimTypesList;
        protected IEnumerable<ClaimTypeConfig> ClaimTypesWithUserMetadata;
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
                IAzureCPConfiguration globalConfiguration = null;
                bool refreshConfig = false;
                bool success = true;
                try
                {
                    if (SPTrust == null)
                    {
                        SPTrust = GetSPTrustAssociatedWithCP(ProviderInternalName);
                        if (SPTrust == null) return false;
                    }
                    if (!CheckIfShouldProcessInput(context)) return false;

                    globalConfiguration = GetConfiguration(context, entityTypes, PersistedObjectName);
                    if (globalConfiguration == null)
                    {
                        AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject not found. Visit AzureCP admin pages in central administration to create it.", ProviderInternalName),
                            TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        // Create a fake persisted object just to get the default settings, it will not be saved in config database
                        globalConfiguration = AzureCPConfig.GetDefaultConfiguration();
                        refreshConfig = true;
                    }
                    else if (globalConfiguration.ClaimTypes == null || globalConfiguration.ClaimTypes.Count == 0)
                    {
                        AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject was found but there are no AzureADObject set. Visit AzureCP admin pages in central administration to create it.", ProviderInternalName),
                            TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        // Cannot continue 
                        success = false;
                    }
                    else if (globalConfiguration.AzureTenants == null || globalConfiguration.AzureTenants.Count == 0)
                    {
                        AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject was found but there are no Azure tenant set. Visit AzureCP admin pages in central administration to add one.", ProviderInternalName),
                            TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        // Cannot continue 
                        success = false;
                    }
                    else
                    {
                        // Persisted object is found and seems valid
                        AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject found, version: {1}, previous version: {2}", ProviderInternalName, ((SPPersistedObject)globalConfiguration).Version.ToString(), this.AzureCPConfigVersion.ToString()),
                            TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Core);
                        if (this.AzureCPConfigVersion != ((SPPersistedObject)globalConfiguration).Version)
                        {
                            refreshConfig = true;
                            this.AzureCPConfigVersion = ((SPPersistedObject)globalConfiguration).Version;
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject changed, refreshing configuration", ProviderInternalName),
                                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
                        }
                    }
                }
                catch (Exception ex)
                {
                    success = false;
                    AzureCPLogging.LogException(ProviderInternalName, "in Initialize", TraceCategory.Core, ex);
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
                        TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);

                    // Create local persisted object that will never be saved in config DB, it's just a local copy
                    // This copy is unique to current object instance to avoid thread safety issues
                    this.CurrentConfiguration = ((AzureCPConfig)globalConfiguration).CloneInReadOnlyObject();
                    //// All settings come from persisted object
                    //this.CurrentConfiguration.AlwaysResolveUserInput = globalConfiguration.AlwaysResolveUserInput;
                    //this.CurrentConfiguration.FilterExactMatchOnly = globalConfiguration.FilterExactMatchOnly;
                    //this.CurrentConfiguration.EnableAugmentation = globalConfiguration.EnableAugmentation;

                    //// Retrieve AzureADObjects
                    //// A copy of collection AzureADObjects must be created because SetActualAADObjectCollection() may change it and it should be made in a copy totally independant from the persisted object
                    //this.CurrentConfiguration.ClaimTypes = new List<ClaimTypeConfig>();
                    //foreach (ClaimTypeConfig currentObject in globalConfiguration.ClaimTypes)
                    //{
                    //    // Create a new AzureADObject
                    //    this.CurrentConfiguration.ClaimTypes.Add(currentObject.CopyPersistedProperties());
                    //}

                    //// Retrieve AzureTenants
                    //// Create a copy of the collection to work in an copy separated from persisted object
                    //this.CurrentConfiguration.AzureTenants = new List<AzureTenant>();
                    //foreach (AzureTenant currentObject in globalConfiguration.AzureTenants)
                    //{
                    //    // Create a copy from persisted object
                    //    this.CurrentConfiguration.AzureTenants.Add(currentObject.CopyPersistedProperties());
                    //}


                    SetCustomConfiguration(context, entityTypes);
                    if (this.CurrentConfiguration.ClaimTypes == null)
                    {
                        // this.CurrentConfiguration.ClaimTypes was set to null in SetCustomConfiguration, which is bad
                        AzureCPLogging.Log(String.Format("[{0}] ClaimTypes was set to null in SetCustomConfiguration, don't set it or set it with actual entries.", ProviderInternalName), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        return false;
                    }

                    if (this.CurrentConfiguration.AzureTenants == null || this.CurrentConfiguration.AzureTenants.Count == 0)
                    {
                        AzureCPLogging.Log(String.Format("[{0}] AzureTenants was not set. Override method SetCustomConfiguration to set it.", ProviderInternalName), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        return false;
                    }

                    // Set properties AuthenticationProvider and GraphService
                    foreach (var coco in this.CurrentConfiguration.AzureTenants)
                    {
                        coco.SetAzureADContext();
                    }
                    success = this.InitializeClaimTypeConfigList(this.CurrentConfiguration.ClaimTypes);
                }
                catch (Exception ex)
                {
                    success = false;
                    AzureCPLogging.LogException(ProviderInternalName, "in Initialize, while refreshing configuration", TraceCategory.Core, ex);
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
        /// <param name="nonProcessedClaimTypes"></param>
        /// <returns></returns>
        private bool InitializeClaimTypeConfigList(ClaimTypeConfigCollection nonProcessedClaimTypes)
        {
            bool success = true;
            try
            {
                bool identityClaimTypeFound = false;
                // Get attributes defined in trust based on their claim type (unique way to map them)
                List<ClaimTypeConfig> claimTypesSetInTrust = new List<ClaimTypeConfig>();
                // There is a bug in the SharePoint API: SPTrustedLoginProvider.ClaimTypes should retrieve SPTrustedClaimTypeInformation.MappedClaimType, but it returns SPTrustedClaimTypeInformation.InputClaimType instead, so we cannot rely on it
                //foreach (var attr in _AttributesDefinitionList.Where(x => AssociatedSPTrustedLoginProvider.ClaimTypes.Contains(x.claimType)))
                //{
                //    attributesDefinedInTrust.Add(attr);
                //}
                foreach (SPTrustedClaimTypeInformation claimTypeInformation in SPTrust.ClaimTypeInformation)
                {
                    // Search if current claim type in trust exists in AzureADObjects
                    // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                    ClaimTypeConfig claimTypeConfig = nonProcessedClaimTypes.FirstOrDefault(x =>
                        String.Equals(x.ClaimType, claimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        !x.CreateAsIdentityClaim &&
                        x.DirectoryObjectProperty != AzureADObjectProperty.None);

                    if (claimTypeConfig == null) continue;
                    claimTypesSetInTrust.Add(claimTypeConfig);
                    if (String.Equals(SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                    {
                        // Identity claim type found, set IdentityAzureADObject property
                        identityClaimTypeFound = true;
                        IdentityClaimTypeConfig = claimTypeConfig;
                    }
                }

                // Check if identity claim is there. Should always check property SPTrustedClaimTypeInformation.MappedClaimType: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.administration.claims.sptrustedclaimtypeinformation.mappedclaimtype.aspx
                if (!identityClaimTypeFound)
                {
                    AzureCPLogging.Log(String.Format("[{0}] Impossible to continue because identity claim type \"{1}\" set in the SPTrustedIdentityTokenIssuer \"{2}\" is missing in AzureADObjects.", ProviderInternalName, SPTrust.IdentityClaimTypeInformation.MappedClaimType, SPTrust.Name), TraceSeverity.Unexpected, EventSeverity.ErrorCritical, TraceCategory.Core);
                    return false;
                }

                // This check is to find if there is a duplicate of the identity claim type that uses the same GraphProperty
                //AzureADObject objectToDelete = claimTypesSetInTrust.Find(x =>
                //    !String.Equals(x.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                //    !x.CreateAsIdentityClaim &&
                //    x.GraphProperty == GraphProperty.UserPrincipalName);
                //if (objectToDelete != null) claimTypesSetInTrust.Remove(objectToDelete);

                // Check if there are objects that should be always queried (CreateAsIdentityClaim) to add in the list
                List<ClaimTypeConfig> additionalClaimTypeConfigList = new List<ClaimTypeConfig>();
                foreach (ClaimTypeConfig claimTypeConfig in nonProcessedClaimTypes.Where(x => x.CreateAsIdentityClaim))// && !claimTypesSetInTrust.Contains(x, new LDAPPropertiesComparer())))
                {
                    // Check if identity claim type is already using same GraphProperty, and ignore current object if so
                    if (IdentityClaimTypeConfig.DirectoryObjectProperty == claimTypeConfig.DirectoryObjectProperty) continue;

                    // Normally ClaimType should be null if CreateAsIdentityClaim is set to true, but we check here it and handle this scenario
                    if (!String.IsNullOrEmpty(claimTypeConfig.ClaimType))
                    {
                        if (String.Equals(SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimTypeConfig.ClaimType))
                        {
                            // Not a big deal since it's set with identity claim type, so no inconsistent behavior to expect, just record an information
                            AzureCPLogging.Log(String.Format("[{0}] Object with GraphProperty {1} is set with CreateAsIdentityClaim to true and ClaimType {2}. Remove ClaimType property as it is useless.", ProviderInternalName, claimTypeConfig.DirectoryObjectProperty, claimTypeConfig.ClaimType), TraceSeverity.Monitorable, EventSeverity.Information, TraceCategory.Core);
                        }
                        else if (claimTypesSetInTrust.Count(x => String.Equals(x.ClaimType, claimTypeConfig.ClaimType)) > 0)
                        {
                            // Same claim type already exists with CreateAsIdentityClaim == false. 
                            // Current object is a bad one and shouldn't be added. Don't add it but continue to build objects list
                            AzureCPLogging.Log(String.Format("[{0}] Claim type {1} is defined twice with CreateAsIdentityClaim set to true and false, which is invalid. Remove entry with CreateAsIdentityClaim set to true.", ProviderInternalName, claimTypeConfig.ClaimType), TraceSeverity.Monitorable, EventSeverity.Information, TraceCategory.Core);
                            continue;
                        }
                    }

                    claimTypeConfig.ClaimType = SPTrust.IdentityClaimTypeInformation.MappedClaimType;    // Give those objects the identity claim type
                    //claimTypeConfig.ClaimEntityType = SPClaimEntityTypes.User;
                    claimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText = IdentityClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText; // Must be set otherwise display text of permissions will be inconsistent
                    additionalClaimTypeConfigList.Add(claimTypeConfig);
                }

                ProcessedClaimTypesList = new List<ClaimTypeConfig>(claimTypesSetInTrust.Count + additionalClaimTypeConfigList.Count);
                ProcessedClaimTypesList.AddRange(claimTypesSetInTrust);
                ProcessedClaimTypesList.AddRange(additionalClaimTypeConfigList);

                // Parse objects to configure some settings
                // An object can have ClaimType set to null if only used to populate metadata of permission created
                foreach (var attr in ProcessedClaimTypesList.Where(x => x.ClaimType != null))
                {
                    var trustedClaim = SPTrust.GetClaimTypeInformationFromMappedClaimType(attr.ClaimType);
                    // It should never be null
                    if (trustedClaim == null) continue;
                    attr.ClaimTypeDisplayName = trustedClaim.DisplayName;
                }

                // Any metadata for a user with GraphProperty actually set is valid
                this.ClaimTypesWithUserMetadata = nonProcessedClaimTypes.Where(x =>
                    !String.IsNullOrEmpty(x.EntityDataKey) &&
                    x.DirectoryObjectProperty != AzureADObjectProperty.None &&
                    //x.ClaimEntityType == SPClaimEntityTypes.User);
                    x.DirectoryObjectType == AzureADObjectType.User);
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in InitializeClaimTypeConfigList", TraceCategory.Core, ex);
                success = false;
            }
            return success;
        }

        /// <summary>
        /// DO NOT Override this method if you use a custom persisted object to hold your configuration.
        /// To get you custom persisted object, you must override property LDAPCP.PersistedObjectName and set its name
        /// </summary>
        /// <returns></returns>
        protected virtual IAzureCPConfiguration GetConfiguration(Uri context, string[] entityTypes, string persistedObjectName)
        {
            return AzureCPConfig.GetConfiguration(persistedObjectName);
            //if (String.Equals(ProviderInternalName, LDAPCP._ProviderInternalName, StringComparison.InvariantCultureIgnoreCase))
            //    return LDAPCPConfig.GetFromConfigDB(persistedObjectName);
            //else
            //    return null;
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
                AzureCPLogging.Log(String.Format("[{0}] Claims provider {0} is associated to multiple SPTrustedIdentityTokenIssuer, which is not supported because at runtime there is no way to determine what TrustedLoginProvider is currently calling", providerInternalName), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);

            AzureCPLogging.Log(String.Format("[{0}] Claims provider {0} is not associated with any SPTrustedIdentityTokenIssuer so it cannot create permissions.\r\nVisit http://ldapcp.codeplex.com for installation procedure or set property ClaimProviderName with PowerShell cmdlet Get-SPTrustedIdentityTokenIssuer to create association.", providerInternalName), TraceSeverity.High, EventSeverity.Warning, TraceCategory.Core);
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
            string claimValue = value;
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
                //pe.EntityType = IdentityClaimTypeConfig.ClaimEntityType;
                pe.EntityType = SPClaimEntityTypes.User;
            }
            else
            {
                permissionValue = FormatPermissionValue(permissionClaimType, permissionValue, isIdentityClaimType, result);
                claim = CreateClaim(
                    permissionClaimType,
                    permissionValue,
                    result.ClaimTypeConfig.ClaimValueType);
                //pe.EntityType = result.ClaimTypeConfig.ClaimEntityType;
                pe.EntityType = result.ClaimTypeConfig.DirectoryObjectType == AzureADObjectType.User ? SPClaimEntityTypes.User : SPClaimEntityTypes.FormsRole;
            }

            pe.DisplayText = FormatPermissionDisplayText(permissionClaimType, permissionValue, isIdentityClaimType, result);
            pe.Description = String.Format(
                PickerEntityOnMouseOver,
                result.ClaimTypeConfig.DirectoryObjectProperty.ToString(),
                result.QueryMatchValue);
            pe.Claim = claim;
            pe.IsResolved = true;
            //pe.EntityGroupName = "";

            int nbMetadata = 0;
            // Populate metadata attributes of permission created
            foreach (var entityAttrib in ClaimTypesWithUserMetadata)
            {
                // if there is actally a value in the GraphObject, then it can be set
                string entityAttribValue = GetGraphPropertyValue(result.UserOrGroupResult, entityAttrib.DirectoryObjectProperty.ToString());
                if (!String.IsNullOrEmpty(entityAttribValue))
                {
                    pe.EntityData[entityAttrib.EntityDataKey] = entityAttribValue;
                    nbMetadata++;
                    AzureCPLogging.Log(String.Format("[{0}] Added metadata \"{1}\" with value \"{2}\" to permission", ProviderInternalName, entityAttrib.EntityDataKey, entityAttribValue), TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
            }

            AzureCPLogging.Log($"[{ProviderInternalName}] Created entity: display text: \"{pe.DisplayText}\", value: \"{pe.Claim.Value}\", claim type: \"{pe.Claim.ClaimType}\", and filled with {nbMetadata.ToString()} metadata.", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
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

            if (result.ClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText != AzureADObjectProperty.None)
            {
                if (!isIdentityClaimType) permissionDisplayText = "(" + result.ClaimTypeConfig.ClaimTypeDisplayName + ") ";

                string graphPropertyToDisplayValue = GetGraphPropertyValue(result.UserOrGroupResult, result.ClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText.ToString());
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
                        result.ClaimTypeConfig.ClaimTypeDisplayName,
                        result.PermissionValue);
                }
            }

            return permissionDisplayText;
        }

        protected virtual PickerEntity CreatePickerEntityForSpecificClaimType(string input, ClaimTypeConfig claimTypesToResolve, bool inputHasKeyword)
        {
            List<PickerEntity> entities = CreatePickerEntityForSpecificClaimTypes(
                input,
                new List<ClaimTypeConfig>()
                    {
                        claimTypesToResolve,
                    },
                inputHasKeyword);
            return entities == null ? null : entities.First();
        }

        protected virtual List<PickerEntity> CreatePickerEntityForSpecificClaimTypes(string input, List<ClaimTypeConfig> claimTypesToResolve, bool inputHasKeyword)
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
                        claimTypeToResolve.ClaimTypeDisplayName,
                        input);
                }

                //pe.EntityType = claimTypeToResolve.ClaimEntityType;
                pe.EntityType = claimTypeToResolve.DirectoryObjectType == AzureADObjectType.User ? SPClaimEntityTypes.User : SPClaimEntityTypes.FormsRole;
                pe.Description = String.Format(
                    PickerEntityOnMouseOver,
                    claimTypeToResolve.DirectoryObjectProperty.ToString(),
                    input);

                pe.Claim = claim;
                pe.IsResolved = true;
                //pe.EntityGroupName = "";

                //if (claimTypeToResolve.ClaimEntityType == SPClaimEntityTypes.User && !String.IsNullOrEmpty(claimTypeToResolve.EntityDataKey))
                if (claimTypeToResolve.DirectoryObjectType == AzureADObjectType.User && !String.IsNullOrEmpty(claimTypeToResolve.EntityDataKey))
                {
                    pe.EntityData[claimTypeToResolve.EntityDataKey] = pe.Claim.Value;
                    AzureCPLogging.Log(String.Format("[{0}] Added metadata \"{1}\" with value \"{2}\" to permission", ProviderInternalName, claimTypeToResolve.EntityDataKey, pe.EntityData[claimTypeToResolve.EntityDataKey]), TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
                entities.Add(pe);
                AzureCPLogging.Log(String.Format("[{0}] Created permission: display text: \"{1}\", value: \"{2}\", claim type: \"{3}\".", ProviderInternalName, pe.DisplayText, pe.Claim.Value, pe.Claim.ClaimType), TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Claims_Picking);
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
                if (ProcessedClaimTypesList == null) return;
                foreach (var claimTypeSettings in ProcessedClaimTypesList)
                {
                    claimTypes.Add(claimTypeSettings.ClaimType);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillClaimTypes", TraceCategory.Core, ex);
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
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Augmentation);
                return;
            }

            if (!Initialize(context, null))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                if (!this.CurrentConfiguration.EnableAugmentation)
                    return;

                AzureCPLogging.Log($"[{ProviderInternalName}] Starting augmentation for user '{decodedEntity.Value}'.", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                //ClaimTypeConfig groupClaimTypeSettings = this.ProcessedClaimTypesList.FirstOrDefault(x => x.ClaimEntityType == SPClaimEntityTypes.FormsRole);
                ClaimTypeConfig groupClaimTypeSettings = this.ProcessedClaimTypesList.FirstOrDefault(x => x.DirectoryObjectType == AzureADObjectType.Group);
                if (groupClaimTypeSettings == null)
                {
                    AzureCPLogging.Log($"[{ProviderInternalName}] No role claim type with SPClaimEntityTypes set to 'FormsRole' was found, please check claims mapping table.",
                        TraceSeverity.High, EventSeverity.Error, TraceCategory.Augmentation);
                    return;
                }

                RequestInformation infos = new RequestInformation(CurrentConfiguration, RequestType.Augmentation, ProcessedClaimTypesList, null, decodedEntity, context, null, null, Int32.MaxValue);
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
                            TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                    }
                    AzureCPLogging.Log($"[{ProviderInternalName}] User '{infos.IncomingEntity.Value}' was augmented with {groups.Count.ToString()} groups in {timer.ElapsedMilliseconds.ToString()} ms",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Augmentation);
                }
                else
                {
                    AzureCPLogging.Log($"[{ProviderInternalName}] No group found for user '{infos.IncomingEntity.Value}', search took {timer.ElapsedMilliseconds.ToString()} ms",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Augmentation);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in AugmentEntity", TraceCategory.Augmentation, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }

        }

        protected async virtual Task<List<SPClaim>> GetGroupMembershipAsync(RequestInformation requestInfo, ClaimTypeConfig groupClaimTypeSettings)
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

        protected async virtual Task<List<SPClaim>> GetGroupMembershipFromAzureADAsync(RequestInformation requestInfo, ClaimTypeConfig groupClaimTypeSettings, AzureTenant tenant)
        {
            List<SPClaim> claims = new List<SPClaim>();
            var userResult = await tenant.GraphService.Users.Request().Filter($"{requestInfo.IdentityClaimTypeConfig.DirectoryObjectProperty} eq '{requestInfo.IncomingEntity.Value}'").GetAsync().ConfigureAwait(false);
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
                    string groupClaimValue = GetGraphPropertyValue(group, groupClaimTypeSettings.DirectoryObjectProperty.ToString());
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
            List<AzureADObjectType> aadEntityTypes = new List<AzureADObjectType>();
            if (entityTypes.Contains(SPClaimEntityTypes.User))
                aadEntityTypes.Add(AzureADObjectType.User);
            if (entityTypes.Contains(SPClaimEntityTypes.FormsRole))
                aadEntityTypes.Add(AzureADObjectType.Group);

            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                if (hierarchyNodeID == null)
                {
                    // Root level
                    //foreach (var azureObject in FinalAttributeList.Where(x => !String.IsNullOrEmpty(x.peoplePickerAttributeHierarchyNodeId) && !x.CreateAsIdentityClaim && entityTypes.Contains(x.ClaimEntityType)))
                    foreach (var azureObject in this.ProcessedClaimTypesList.FindAll(x => !x.CreateAsIdentityClaim && aadEntityTypes.Contains(x.DirectoryObjectType)))
                    {
                        hierarchy.AddChild(
                            new Microsoft.SharePoint.WebControls.SPProviderHierarchyNode(
                                _ProviderInternalName,
                                azureObject.ClaimTypeDisplayName,
                                azureObject.ClaimType,
                                true));
                    }
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillHierarchy", TraceCategory.Claims_Picking, ex);
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
            AzureCPLogging.LogDebug($"context passed to FillResolve (SPClaim): {context.ToString()}");
            List<AzureADObjectType> aadEntityTypes = new List<AzureADObjectType>();
            if (entityTypes.Contains(SPClaimEntityTypes.User))
                aadEntityTypes.Add(AzureADObjectType.User);
            if (entityTypes.Contains(SPClaimEntityTypes.FormsRole))
                aadEntityTypes.Add(AzureADObjectType.Group);

            if (!Initialize(context, entityTypes))
                return;

            // Ensure incoming claim should be validated by AzureCP
            // Must be made after call to Initialize because SPTrustedLoginProvider name must be known
            if (!String.Equals(resolveInput.OriginalIssuer, IssuerName, StringComparison.InvariantCultureIgnoreCase))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                RequestInformation infos = new RequestInformation(CurrentConfiguration, RequestType.Validation, ProcessedClaimTypesList, resolveInput.Value, resolveInput, context, aadEntityTypes.ToArray(), null, Int32.MaxValue);
                List<PickerEntity> permissions = SearchOrValidate(infos);
                if (permissions.Count == 1)
                {
                    resolved.Add(permissions[0]);
                    AzureCPLogging.Log(String.Format("[{0}] Validated permission: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, permissions[0].Claim.Value, permissions[0].Claim.ClaimType),
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
                else
                {
                    AzureCPLogging.Log(String.Format("[{0}] Validation of incoming claim returned {1} permissions instead of 1 expected. Aborting operation", ProviderInternalName, permissions.Count.ToString()), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Claims_Picking);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillResolve(SPClaim)", TraceCategory.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        protected override void FillResolve(Uri context, string[] entityTypes, string resolveInput, List<Microsoft.SharePoint.WebControls.PickerEntity> resolved)
        {
            List<AzureADObjectType> aadEntityTypes = new List<AzureADObjectType>();
            if (entityTypes.Contains(SPClaimEntityTypes.User))
                aadEntityTypes.Add(AzureADObjectType.User);
            if (entityTypes.Contains(SPClaimEntityTypes.FormsRole))
                aadEntityTypes.Add(AzureADObjectType.Group);

            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                RequestInformation settings = new RequestInformation(CurrentConfiguration, RequestType.Search, ProcessedClaimTypesList, resolveInput, null, context, aadEntityTypes.ToArray(), null, Int32.MaxValue);
                List<PickerEntity> permissions = SearchOrValidate(settings);
                FillPermissions(context, entityTypes, resolveInput, ref permissions);
                foreach (PickerEntity entity in permissions)
                {
                    resolved.Add(entity);
                    AzureCPLogging.Log(String.Format("[{0}] Added entity: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType),
                        TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillResolve(string)", TraceCategory.Claims_Picking, ex);
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
            List<AzureADObjectType> aadEntityTypes = new List<AzureADObjectType>();
            if (entityTypes.Contains(SPClaimEntityTypes.User))
                aadEntityTypes.Add(AzureADObjectType.User);
            if (entityTypes.Contains(SPClaimEntityTypes.FormsRole))
                aadEntityTypes.Add(AzureADObjectType.Group);

            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                RequestInformation settings = new RequestInformation(CurrentConfiguration, RequestType.Search, ProcessedClaimTypesList, searchPattern, null, context, aadEntityTypes.ToArray(), hierarchyNodeID, maxCount);
                List<PickerEntity> permissions = SearchOrValidate(settings);
                FillPermissions(context, entityTypes, searchPattern, ref permissions);
                SPProviderHierarchyNode matchNode = null;
                foreach (PickerEntity entity in permissions)
                {
                    // Add current PickerEntity to the corresponding attribute in the hierarchy
                    if (searchTree.HasChild(entity.Claim.ClaimType))
                    {
                        matchNode = searchTree.Children.First(x => x.HierarchyNodeID == entity.Claim.ClaimType);
                    }
                    else
                    {
                        ClaimTypeConfig attrHelper = ProcessedClaimTypesList.FirstOrDefault(x =>
                            !x.CreateAsIdentityClaim &&
                            String.Equals(x.ClaimType, entity.Claim.ClaimType, StringComparison.InvariantCultureIgnoreCase));

                        string nodeName = attrHelper != null ? attrHelper.ClaimTypeDisplayName : entity.Claim.ClaimType;
                        matchNode = new SPProviderHierarchyNode(_ProviderInternalName, nodeName, entity.Claim.ClaimType, true);
                        searchTree.AddChild(matchNode);
                    }
                    matchNode.AddEntity(entity);
                    AzureCPLogging.Log(String.Format("[{0}] Added entity: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType),
                        TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillSearch", TraceCategory.Claims_Picking, ex);
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
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                        }
                    }
                    return permissions;
                }

                if (requestInfo.RequestType == RequestType.Search)
                {
                    List<ClaimTypeConfig> attribsMatchInputPrefix = requestInfo.ClaimTypeConfigList.FindAll(x =>
                        !String.IsNullOrEmpty(x.PrefixToBypassLookup) &&
                        requestInfo.Input.StartsWith(x.PrefixToBypassLookup, StringComparison.InvariantCultureIgnoreCase));
                    if (attribsMatchInputPrefix.Count > 0)
                    {
                        // Input has a prefix, so it should be validated with no lookup
                        ClaimTypeConfig attribMatchInputPrefix = attribsMatchInputPrefix.First();
                        if (attribsMatchInputPrefix.Count > 1)
                        {
                            // Multiple attributes have same prefix, which is not allowed
                            AzureCPLogging.Log(String.Format("[{0}] Multiple attributes have same prefix ({1}), which is not allowed.", ProviderInternalName, attribMatchInputPrefix.PrefixToBypassLookup), TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Claims_Picking);
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
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
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
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                        }
                        return permissions;
                    }
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in SearchOrValidate", TraceCategory.Claims_Picking, ex);
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

            List<AzureADResult> aadResults = null;
            using (new SPMonitoredScope($"[{ProviderInternalName}] Total time spent to query Azure AD tenant(s)", 1000))
            {
                // Call async method in a task to avoid error "Asynchronous operations are not allowed in this context" error when permission is validated (POST from people picker)
                // More info on the error: https://stackoverflow.com/questions/672237/running-an-asynchronous-operation-triggered-by-an-asp-net-web-page-request
                Task azureADQueryTask = Task.Run(async () =>
                {
                    //Task<List<AzureADResult>> taskAadResults = QueryAzureADTenantsAsync(requestInfo, userFilter, groupFilter, userSelect, groupSelect);
                    //taskAadResults.ConfigureAwait(false);
                    //taskAadResults.Wait();
                    //aadResults = taskAadResults.Result;
                    aadResults = await QueryAzureADTenantsAsync(requestInfo, userFilter, groupFilter, userSelect, groupSelect);
                });
                azureADQueryTask.Wait();
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
                            TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                    }
                }
            }
        }

        ///// <summary>
        ///// Build filter and select statements sent to Azure AD
        ///// $filter and $select must be URL encoded as documented in https://developer.microsoft.com/en-us/graph/docs/concepts/query_parameters#encoding-query-parameters
        ///// </summary>
        ///// <param name="requestInfo"></param>
        ///// <param name="userFilter">User filter</param>
        ///// <param name="groupFilter">Group filter</param>
        ///// <param name="userSelect">User properties to get from AAD</param>
        ///// <param name="groupSelect">Group properties to get from AAD</param>
        //protected virtual void BuildFilter(RequestInformation requestInfo, out string userFilter, out string groupFilter, out string userSelect, out string groupSelect)
        //{
        //    // TODO: Move this outside of this method
        //    /////
        //    List<GraphPropertyQuery> propertyQueries = new List<GraphPropertyQuery>();
        //    foreach (GraphProperty prop in Enum.GetValues(typeof(GraphProperty)))
        //    {
        //        GraphPropertyQuery querySyntax = new GraphPropertyQuery(prop);
        //        //if (prop == GraphProperty.Id) querySyntax.SearchQuery = querySyntax.ValidationQuery;    // ID does not support 'startswith'
        //        if (prop == GraphProperty.Id) querySyntax.FieldType = typeof(Guid);
        //        propertyQueries.Add(querySyntax);
        //    }
        //    //propertyQueries.FirstOrDefault(x => x.PropertyName == )
        //    /////

        //    StringBuilder userFilterBuilder = new StringBuilder("accountEnabled eq true and (");
        //    StringBuilder groupFilterBuilder = new StringBuilder();
        //    StringBuilder userSelectBuilder = new StringBuilder("UserType, Mail, ");    // UserType and Mail are always needed to deal with Guest users
        //    StringBuilder groupSelectBuilder = new StringBuilder();

        //    string searchPattern;
        //    string input = requestInfo.Input;
        //    if (requestInfo.ExactSearch) searchPattern = "{0} eq '" + input + "'";
        //    else searchPattern = "startswith({0},'" + input + "')";

        //    bool firstUserObject = true;
        //    bool firstGroupObject = true;
        //    foreach (AzureADObject adObject in requestInfo.ClaimTypeConfigList)
        //    {
        //        GraphPropertyQuery querySyntax = propertyQueries.FirstOrDefault(x => x.PropertyName == adObject.GraphProperty);
        //        //if (requestInfo.ExactSearch) searchPattern = String.Format(querySyntax.ValidationQuery, "{0}", input);
        //        //else searchPattern = String.Format(querySyntax.SearchQuery, "{0}", input);

        //        string property = adObject.GraphProperty.ToString();
        //        string objectFilter = String.Format(searchPattern, property);
        //        string objectSelect = property;
        //        if (adObject.ClaimEntityType == SPClaimEntityTypes.User)
        //        {
        //            if (firstUserObject) firstUserObject = false;
        //            else
        //            {
        //                objectFilter = " or " + objectFilter;
        //                objectSelect = ", " + objectSelect;
        //            }
        //            userFilterBuilder.Append(objectFilter);
        //            userSelectBuilder.Append(objectSelect);
        //        }
        //        else
        //        {
        //            // else with no further test assumes everything that is not a User is a Group
        //            if (firstGroupObject) firstGroupObject = false;
        //            else
        //            {
        //                objectFilter = objectFilter + " or ";
        //                objectSelect = ", " + objectSelect;
        //            }
        //            groupFilterBuilder.Append(objectFilter);
        //            groupSelectBuilder.Append(objectSelect);
        //        }
        //    }

        //    // Also add properties in user metadata list to $select
        //    foreach (AzureADObject adObject in UserMetadataClaimTypeConfigList)
        //    {
        //        string property = adObject.GraphProperty.ToString();
        //        string objectSelect = property;
        //        if (firstUserObject) firstUserObject = false;
        //        else
        //        {
        //            objectSelect = ", " + objectSelect;
        //        }
        //        userSelectBuilder.Append(objectSelect);
        //    }

        //    userFilterBuilder.Append(")");

        //    userFilter = HttpUtility.UrlEncode(userFilterBuilder.ToString());
        //    groupFilter = HttpUtility.UrlEncode(groupFilterBuilder.ToString());
        //    userSelect = HttpUtility.UrlEncode(userSelectBuilder.ToString());
        //    groupSelect = HttpUtility.UrlEncode(groupSelectBuilder.ToString());
        //}

        /// <summary>
        /// Build filter and select statements sent to Azure AD
        /// $filter and $select must be URL encoded as documented in https://developer.microsoft.com/en-us/graph/docs/concepts/query_parameters#encoding-query-parameters
        /// </summary>
        /// <param name="requestInfo"></param>
        /// <param name="userFilter">User filter</param>
        /// <param name="groupFilter">Group filter</param>
        /// <param name="userSelect">User properties to get from AAD</param>
        /// <param name="groupSelect">Group properties to get from AAD</param>
        protected virtual void BuildFilter(RequestInformation requestInfo, out string userFilter, out string groupFilter, out string userSelect, out string groupSelect)
        {
            StringBuilder userFilterBuilder = new StringBuilder("accountEnabled eq true and (");
            StringBuilder groupFilterBuilder = new StringBuilder();
            StringBuilder userSelectBuilder = new StringBuilder("UserType, Mail, ");    // UserType and Mail are always needed to deal with Guest users
            StringBuilder groupSelectBuilder = new StringBuilder("Id, ");               // Id is always required for groups

            string preferredSearchPattern;
            string input = requestInfo.Input;
            //if (requestInfo.ExactSearch) preferredSearchPattern = "{0} eq '" + input + "'";
            //else preferredSearchPattern = "startswith({0},'" + input + "')";

            if (requestInfo.ExactSearch) preferredSearchPattern = String.Format(ClaimsProviderConstants.SearchPatternEquals, "{0}", input);
            else preferredSearchPattern = String.Format(ClaimsProviderConstants.SearchPatternStartsWith, "{0}", input);

            bool firstUserObject = true;
            bool firstGroupObject = true;
            foreach (ClaimTypeConfig adObject in requestInfo.ClaimTypeConfigList)
            {
                string property = adObject.DirectoryObjectProperty.ToString();
                string objectFilter = String.Format(preferredSearchPattern, property);
                if (adObject.DirectoryObjectProperty == AzureADObjectProperty.Id)
                {
                    Guid idGuid = new Guid();
                    if (!Guid.TryParse(input, out idGuid)) continue;
                    else objectFilter = String.Format(ClaimsProviderConstants.SearchPatternEquals, property, idGuid.ToString());
                }

                string objectSelect = property;
                //if (adObject.ClaimEntityType == SPClaimEntityTypes.User)
                if (adObject.DirectoryObjectType == AzureADObjectType.User)
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

            // Also add properties in user metadata list to $select
            foreach (ClaimTypeConfig adObject in ClaimTypesWithUserMetadata)
            {
                string property = adObject.DirectoryObjectProperty.ToString();
                string objectSelect = property;
                if (firstUserObject) firstUserObject = false;
                else
                {
                    objectSelect = ", " + objectSelect;
                }
                userSelectBuilder.Append(objectSelect);
            }

            userFilterBuilder.Append(")");

            // Detect if some properties were actually found for each object type, if not, return an empty filter
            if (firstUserObject) userFilterBuilder.Clear();
            if (firstGroupObject) groupFilterBuilder.Clear();

            userFilter = HttpUtility.UrlEncode(userFilterBuilder.ToString());
            groupFilter = HttpUtility.UrlEncode(groupFilterBuilder.ToString());
            userSelect = HttpUtility.UrlEncode(userSelectBuilder.ToString());
            groupSelect = HttpUtility.UrlEncode(groupSelectBuilder.ToString());
        }

        protected virtual async Task<List<AzureADResult>> QueryAzureADTenantsAsync(RequestInformation requestInfo, string userFilter, string groupFilter, string userSelect, string groupSelect)
        {
            if (userFilter == null && groupFilter == null) return null;
            List<AzureADResult> allSearchResults = new List<AzureADResult>();
            var lockResults = new object();

            //foreach (AzureTenant coco in this.CurrentConfiguration.AzureTenants)
            Parallel.ForEach(this.CurrentConfiguration.AzureTenants, async coco =>
            //var queryTenantTasks = this.CurrentConfiguration.AzureTenants.Select (async coco =>
            {
                Stopwatch timer = new Stopwatch();
                AzureADResult searchResult = null;
                try
                {
                    timer.Start();
                    searchResult = await QueryAzureADTenantAsync(requestInfo, coco, userFilter, groupFilter, userSelect, groupSelect, true).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    AzureCPLogging.LogException(ProviderInternalName, String.Format("in QueryAzureADTenantsAsync while querying tenant {0}", coco.TenantName), TraceCategory.Lookup, ex);
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
                    }
                    AzureCPLogging.Log($"[{ProviderInternalName}] Got {searchResult.UserOrGroupResultList.Count().ToString()} users/groups and {searchResult.DomainsRegisteredInAzureADTenant.Count().ToString()} registered domains in {timer.ElapsedMilliseconds.ToString()} ms from '{coco.TenantName}' with input '{requestInfo.Input}'",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                }
                else AzureCPLogging.Log($"[{ProviderInternalName}] Got no result from '{coco.TenantName}' with input '{requestInfo.Input}', search took {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
            });
            //}
            return allSearchResults;
        }

        protected virtual async Task<AzureADResult> QueryAzureADTenantAsync(RequestInformation requestInfo, AzureTenant coco, string userFilter, string groupFilter, string userSelect, string groupSelect, bool firstAttempt)
        {
            AzureCPLogging.Log($"[{ProviderInternalName}] Querying Azure AD tenant '{coco.TenantName}' for users/groups/domains, with input '{requestInfo.Input}'", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Lookup);
            AzureADResult tenantResults = new AzureADResult();
            bool tryAgain = false;
            object lockAddResultToCollection = new object();
            CancellationTokenSource cts = new CancellationTokenSource(ClaimsProviderConstants.timeout);
            try
            {
                using (new SPMonitoredScope($"[{ProviderInternalName}] Querying Azure AD tenant '{coco.TenantName}' for users/groups/domains, with input '{requestInfo.Input}'", 1000))
                {
                    // No need to lock here: as per https://stackoverflow.com/questions/49108179/need-advice-on-getting-access-token-with-multiple-task-in-microsoft-graph:
                    // The Graph client object is thread-safe and re-entrant
                    Task userQueryTask = Task.Run(async () =>
                    {
                        AzureCPLogging.LogDebug($"[{ProviderInternalName}] UserQueryTask starting for tenant '{coco.TenantName}'");
                        if (String.IsNullOrEmpty(userFilter)) return;
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
                        AzureCPLogging.LogDebug($"[{ProviderInternalName}] UserQueryTask ended for tenant '{coco.TenantName}'");
                    }, cts.Token);
                    Task groupQueryTask = Task.Run(async () =>
                    {
                        AzureCPLogging.LogDebug($"[{ProviderInternalName}] GroupQueryTask starting for tenant '{coco.TenantName}'");
                        if (String.IsNullOrEmpty(groupFilter)) return;
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
                        AzureCPLogging.LogDebug($"[{ProviderInternalName}] GroupQueryTask ended for tenant '{coco.TenantName}'");
                    }, cts.Token);
                    Task domainQueryTask = Task.Run(async () =>
                    {
                        AzureCPLogging.LogDebug($"[{ProviderInternalName}] DomainQueryTask starting for tenant '{coco.TenantName}'");
                        IGraphServiceDomainsCollectionPage domains = await coco.GraphService.Domains.Request().GetAsync();
                        lock (lockAddResultToCollection)
                        {
                            tenantResults.DomainsRegisteredInAzureADTenant.AddRange(domains.Where(x => x.IsVerified == true).Select(x => x.Id));
                        }
                        AzureCPLogging.LogDebug($"[{ProviderInternalName}] DomainQueryTask ended for tenant '{coco.TenantName}'");
                    }, cts.Token);

                    Task.WaitAll(new Task[3] { userQueryTask, groupQueryTask, domainQueryTask }, ClaimsProviderConstants.timeout, cts.Token);
                    //await Task.WhenAll(userQueryTask, groupQueryTask).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
                AzureCPLogging.Log($"[{ProviderInternalName}] Query on Azure AD tenant '{coco.TenantName}' exceeded timeout of {ClaimsProviderConstants.timeout} ms and was cancelled.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
                tryAgain = true;
            }
            catch (AggregateException ex)
            {
                // Task.WaitAll throws an AggregateException, which contains all exceptions thrown by tasks it waited on
                AzureCPLogging.LogException(ProviderInternalName, $"while querying tenant '{coco.TenantName}'", TraceCategory.Lookup, ex);
                tryAgain = true;
            }
            finally
            {
                AzureCPLogging.LogDebug($"[{ProviderInternalName}] Releasing cancellation token of tenant '{coco.TenantName}'");
                cts.Dispose();
            }

            if (firstAttempt && tryAgain)
            {
                AzureCPLogging.Log($"[{ProviderInternalName}] Doing new attempt to query tenant '{coco.TenantName}'...",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                tenantResults = await QueryAzureADTenantAsync(requestInfo, coco, userFilter, groupFilter, userSelect, groupSelect, false).ConfigureAwait(false);
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
            List<ClaimTypeConfig> claimTypeConfigList;
            if (requestInfo.ExactSearch) claimTypeConfigList = requestInfo.ClaimTypeConfigList.FindAll(x => !x.CreateAsIdentityClaim);
            else claimTypeConfigList = requestInfo.ClaimTypeConfigList;

            List<AzureCPResult> processedResults = new List<AzureCPResult>();
            foreach (DirectoryObject userOrGroup in usersAndGroupsResults)
            {
                DirectoryObject currentObject = null;
                //string claimEntityType = null;
                AzureADObjectType objectType;
                if (userOrGroup is User)
                {
                    // Always skip shadow users: UserType is Guest and his mail matches a verified domain in AAD tenant
                    string userType = GetGraphPropertyValue(userOrGroup, "UserType");
                    if (String.IsNullOrEmpty(userType))
                    {
                        AzureCPLogging.Log(
                            String.Format("[{0}] User {1} filtered out because his property UserType is empty.", ProviderInternalName, ((User)userOrGroup).UserPrincipalName),
                            TraceSeverity.Unexpected, EventSeverity.Warning, TraceCategory.Lookup);
                        continue;
                    }
                    if (String.Equals(userType, ClaimsProviderConstants.GraphUserType.Guest, StringComparison.InvariantCultureIgnoreCase))
                    {
                        string mail = GetGraphPropertyValue(userOrGroup, "Mail");
                        if (String.IsNullOrEmpty(mail))
                        {
                            AzureCPLogging.Log(
                                String.Format("[{0}] Guest user {1} filtered out because his mail is empty.", ProviderInternalName, ((User)userOrGroup).UserPrincipalName),
                                TraceSeverity.Unexpected, EventSeverity.Warning, TraceCategory.Lookup);
                            continue;
                        }
                        if (!mail.Contains('@')) continue;
                        string maildomain = mail.Split('@')[1];
                        if (domains.Any(x => String.Equals(x, maildomain, StringComparison.InvariantCultureIgnoreCase)))
                        {
                            AzureCPLogging.Log(
                                String.Format("[{0}] Guest user {1} filtered out because he is in a domain registered in AAD tenant.", ProviderInternalName, mail),
                                TraceSeverity.Verbose, EventSeverity.Verbose, TraceCategory.Lookup);
                            continue;
                        }
                    }
                    currentObject = userOrGroup;
                    //claimEntityType = SPClaimEntityTypes.User;
                    objectType = AzureADObjectType.User;
                }
                else
                {
                    currentObject = userOrGroup;
                    //claimEntityType = SPClaimEntityTypes.FormsRole;
                    objectType = AzureADObjectType.Group;
                }

                // Start filter
                //foreach (ClaimTypeConfig claimTypeConfig in claimTypeConfigList.Where(x => x.ClaimEntityType == claimEntityType))
                foreach (ClaimTypeConfig claimTypeConfig in claimTypeConfigList.Where(x => x.DirectoryObjectType == objectType))
                {
                    // Get value with of current GraphProperty
                    string graphPropertyValue = GetGraphPropertyValue(currentObject, claimTypeConfig.DirectoryObjectProperty.ToString());

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
                    ClaimTypeConfig objCompare;
                    if (claimTypeConfig.CreateAsIdentityClaim)
                    {
                        objCompare = IdentityClaimTypeConfig;
                        // Get the value of the GraphProperty linked to IdentityAzureObject
                        valueToCheck = GetGraphPropertyValue(currentObject, IdentityClaimTypeConfig.DirectoryObjectProperty.ToString());
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
                TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Lookup);
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
                AzureCPLogging.LogException(ProviderInternalName, "in GetClaimTypeForUserKey", TraceCategory.Rehydration, ex);
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
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Rehydration);
                return CreateClaim(IdentityClaimTypeConfig.ClaimType, curUser.Value, curUser.ValueType);
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in GetUserKeyForEntity", TraceCategory.Rehydration, ex);
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
        public ClaimTypeConfig ClaimTypeConfig;
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
