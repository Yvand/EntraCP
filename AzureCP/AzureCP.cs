using Microsoft.Graph;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using static azurecp.ClaimsProviderLogging;
using WIF4_5 = System.Security.Claims;

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
        public static string _ProviderInternalName => "AzureCP";
        public virtual string ProviderInternalName => "AzureCP";
        public virtual string PersistedObjectName => ClaimsProviderConstants.CONFIG_NAME;

        private object Lock_Init = new object();
        private ReaderWriterLockSlim Lock_Config = new ReaderWriterLockSlim();
        private long CurrentConfigurationVersion = 0;

        /// <summary>
        /// Contains configuration currently used by claims provider
        /// </summary>
        public IAzureCPConfiguration CurrentConfiguration
        {
            get => _CurrentConfiguration;
            set => _CurrentConfiguration = value;
        }
        private IAzureCPConfiguration _CurrentConfiguration;

        /// <summary>
        /// SPTrust associated with the claims provider
        /// </summary>
        protected SPTrustedLoginProvider SPTrust;

        /// <summary>
        /// ClaimTypeConfig mapped to the identity claim in the SPTrustedIdentityTokenIssuer
        /// </summary>
        IdentityClaimTypeConfig IdentityClaimTypeConfig;

        /// <summary>
        /// Group ClaimTypeConfig used to set the claim type for other group ClaimTypeConfig that have UseMainClaimTypeOfDirectoryObject set to true
        /// </summary>
        ClaimTypeConfig MainGroupClaimTypeConfig;

        /// <summary>
        /// Processed list to use. It is guarranted to never contain an empty ClaimType
        /// </summary>
        public List<ClaimTypeConfig> ProcessedClaimTypesList
        {
            get => _ProcessedClaimTypesList;
            set => _ProcessedClaimTypesList = value;
        }
        private List<ClaimTypeConfig> _ProcessedClaimTypesList;

        protected IEnumerable<ClaimTypeConfig> MetadataConfig;
        protected virtual string PickerEntityDisplayText => "({0}) {1}";
        protected virtual string PickerEntityOnMouseOver => "{0}={1}";

        /// <summary>
        /// Returned issuer formatted like the property SPClaim.OriginalIssuer: "TrustedProvider:TrustedProviderName"
        /// </summary>
        protected string IssuerName => SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, SPTrust.Name);

        public AzureCP(string displayName) : base(displayName) { }

        /// <summary>
        /// Initializes claim provider. This method is reserved for internal use and is not intended to be called from external code or changed
        /// </summary>
        public bool Initialize(Uri context, string[] entityTypes)
        {
            // Ensures thread safety to initialize class variables
            lock (Lock_Init)
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

                    globalConfiguration = GetConfiguration(context, entityTypes, PersistedObjectName, SPTrust.Name);
                    if (globalConfiguration == null)
                    {
                        ClaimsProviderLogging.Log($"[{ProviderInternalName}] Configuration '{PersistedObjectName}' was not found in configuration database, use default configuration instead. Visit AzureCP admin pages in central administration to create it.",
                            TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        // Return default configuration and set refreshConfig to true to give a chance to deprecated method SetCustomConfiguration() to set AzureTenants list
                        globalConfiguration = AzureCPConfig.ReturnDefaultConfiguration(SPTrust.Name);
                        refreshConfig = true;
                    }
                    else
                    {
                        ((AzureCPConfig)globalConfiguration).CheckAndCleanConfiguration(SPTrust.Name);
                    }

                    if (globalConfiguration.ClaimTypes == null || globalConfiguration.ClaimTypes.Count == 0)
                    {
                        ClaimsProviderLogging.Log($"[{ProviderInternalName}] Configuration '{PersistedObjectName}' was found but collection ClaimTypes is null or empty. Visit AzureCP admin pages in central administration to create it.",
                            TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        // Cannot continue 
                        success = false;
                    }

                    if (success)
                    {
                        if (this.CurrentConfigurationVersion == ((SPPersistedObject)globalConfiguration).Version)
                        {
                            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Configuration '{PersistedObjectName}' was found, version {((SPPersistedObject)globalConfiguration).Version.ToString()}",
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Core);
                        }
                        else
                        {
                            refreshConfig = true;
                            this.CurrentConfigurationVersion = ((SPPersistedObject)globalConfiguration).Version;
                            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Configuration '{PersistedObjectName}' changed to version {((SPPersistedObject)globalConfiguration).Version.ToString()}, refreshing local copy",
                                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
                        }
                    }

                    // ProcessedClaimTypesList can be null if:
                    // - 1st initialization
                    // - Initialized before but it failed. If so, try again to refresh config
                    if (this.ProcessedClaimTypesList == null)
                    {
                        refreshConfig = true;
                    }

                    // If config is already initialized, double check that property GraphService is not null as it is required to query AAD tenants
                    if (!refreshConfig)
                    {
                        foreach (var tenant in this.CurrentConfiguration.AzureTenants)
                        {
                            if (tenant.GraphService == null)
                            {
                                // Mark config to be refreshed in the write lock
                                refreshConfig = true;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    success = false;
                    ClaimsProviderLogging.LogException(ProviderInternalName, "in Initialize", TraceCategory.Core, ex);
                }

                if (!success || !refreshConfig)
                {
                    return success;
                }

                // 2ND PART: APPLY CONFIGURATION
                // Configuration needs to be refreshed, lock current thread in write mode
                Lock_Config.EnterWriteLock();
                try
                {
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Refreshing local copy of configuration '{PersistedObjectName}'",
                        TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Core);

                    // Create local persisted object that will never be saved in config DB, it's just a local copy
                    // This copy is unique to current object instance to avoid thread safety issues
                    this.CurrentConfiguration = ((AzureCPConfig)globalConfiguration).CopyConfiguration();

#pragma warning disable CS0618 // Type or member is obsolete
                    SetCustomConfiguration(context, entityTypes);
#pragma warning restore CS0618 // Type or member is obsolete
                    if (this.CurrentConfiguration.ClaimTypes == null)
                    {
                        ClaimsProviderLogging.Log($"[{ProviderInternalName}] List if claim types was set to null in method SetCustomConfiguration for configuration '{PersistedObjectName}'.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        return false;
                    }

                    if (this.CurrentConfiguration.AzureTenants == null || this.CurrentConfiguration.AzureTenants.Count == 0)
                    {
                        ClaimsProviderLogging.Log($"[{ProviderInternalName}] There is no Azure tenant registered in the configuration '{PersistedObjectName}'. Visit AzureCP in central administration to add it, or override method GetConfiguration.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                        return false;
                    }

                    // Set properties AuthenticationProvider and GraphService
                    foreach (var tenant in this.CurrentConfiguration.AzureTenants)
                    {
                        tenant.SetAzureADContext(ProviderInternalName, this.CurrentConfiguration.Timeout);
                    }
                    success = this.InitializeClaimTypeConfigList(this.CurrentConfiguration.ClaimTypes);
                }
                catch (Exception ex)
                {
                    success = false;
                    ClaimsProviderLogging.LogException(ProviderInternalName, "in Initialize, while refreshing configuration", TraceCategory.Core, ex);
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
                bool groupClaimTypeFound = false;
                List<ClaimTypeConfig> claimTypesSetInTrust = new List<ClaimTypeConfig>();
                // Foreach MappedClaimType in the SPTrustedLoginProvider
                foreach (SPTrustedClaimTypeInformation claimTypeInformation in SPTrust.ClaimTypeInformation)
                {
                    // Search if current claim type in trust exists in ClaimTypeConfigCollection
                    ClaimTypeConfig claimTypeConfig = nonProcessedClaimTypes.FirstOrDefault(x =>
                        String.Equals(x.ClaimType, claimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        !x.UseMainClaimTypeOfDirectoryObject &&
                        x.DirectoryObjectProperty != AzureADObjectProperty.NotSet);

                    if (claimTypeConfig == null)
                    {
                        continue;
                    }
                    claimTypeConfig.ClaimTypeDisplayName = claimTypeInformation.DisplayName;
                    claimTypesSetInTrust.Add(claimTypeConfig);
                    if (String.Equals(SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                    {
                        // Identity claim type found, set IdentityClaimTypeConfig property
                        identityClaimTypeFound = true;
                        IdentityClaimTypeConfig = IdentityClaimTypeConfig.ConvertClaimTypeConfig(claimTypeConfig);
                    }
                    else if (!groupClaimTypeFound && claimTypeConfig.EntityType == DirectoryObjectType.Group)
                    {
                        groupClaimTypeFound = true;
                        MainGroupClaimTypeConfig = claimTypeConfig;
                    }
                }

                if (!identityClaimTypeFound)
                {
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Cannot continue because identity claim type '{SPTrust.IdentityClaimTypeInformation.MappedClaimType}' set in the SPTrustedIdentityTokenIssuer '{SPTrust.Name}' is missing in the ClaimTypeConfig list.", TraceSeverity.Unexpected, EventSeverity.ErrorCritical, TraceCategory.Core);
                    return false;
                }

                // Check if there are additional properties to use in queries (UseMainClaimTypeOfDirectoryObject set to true)
                List<ClaimTypeConfig> additionalClaimTypeConfigList = new List<ClaimTypeConfig>();
                foreach (ClaimTypeConfig claimTypeConfig in nonProcessedClaimTypes.Where(x => x.UseMainClaimTypeOfDirectoryObject))
                {
                    if (claimTypeConfig.EntityType == DirectoryObjectType.User)
                    {
                        claimTypeConfig.ClaimType = IdentityClaimTypeConfig.ClaimType;
                        claimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText = IdentityClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText;
                    }
                    else
                    {
                        // If not a user, it must be a group
                        if (MainGroupClaimTypeConfig == null)
                        {
                            continue;
                        }
                        claimTypeConfig.ClaimType = MainGroupClaimTypeConfig.ClaimType;
                        claimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText = MainGroupClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText;
                        claimTypeConfig.ClaimTypeDisplayName = MainGroupClaimTypeConfig.ClaimTypeDisplayName;
                    }
                    additionalClaimTypeConfigList.Add(claimTypeConfig);
                }

                this.ProcessedClaimTypesList = new List<ClaimTypeConfig>(claimTypesSetInTrust.Count + additionalClaimTypeConfigList.Count);
                this.ProcessedClaimTypesList.AddRange(claimTypesSetInTrust);
                this.ProcessedClaimTypesList.AddRange(additionalClaimTypeConfigList);

                // Get all PickerEntity metadata with a DirectoryObjectProperty set
                this.MetadataConfig = nonProcessedClaimTypes.Where(x =>
                    !String.IsNullOrEmpty(x.EntityDataKey) &&
                    x.DirectoryObjectProperty != AzureADObjectProperty.NotSet);
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in InitializeClaimTypeConfigList", TraceCategory.Core, ex);
                success = false;
            }
            return success;
        }

        /// <summary>
        /// Override this method to return a custom configuration of AzureCP.
        /// DO NOT Override this method if you use a custom persisted object to store configuration in config DB.
        /// To use a custom persisted object, override property PersistedObjectName and set its name
        /// </summary>
        /// <returns></returns>
        protected virtual IAzureCPConfiguration GetConfiguration(Uri context, string[] entityTypes, string persistedObjectName, string spTrustName)
        {
            return AzureCPConfig.GetConfiguration(persistedObjectName, spTrustName);
        }

        /// <summary>
        /// [Deprecated] Override this method to customize the configuration of AzureCP. Please override GetConfiguration instead.
        /// </summary> 
        /// <param name="context">The context, as a URI</param>
        /// <param name="entityTypes">The EntityType entity types set to scope the search to</param>
        [Obsolete("SetCustomConfiguration is deprecated, please override GetConfiguration instead.")]
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
            if (context == null) { return true; }
            var webApp = SPWebApplication.Lookup(context);
            if (webApp == null) { return false; }
            if (webApp.IsAdministrationWebApplication) { return true; }

            // Not central admin web app, enable AzureCP only if current web app uses it
            // It is not possible to exclude zones where AzureCP is not used because:
            // Consider following scenario: default zone is WinClaims, intranet zone is Federated:
            // In intranet zone, when creating permission, AzureCP will be called 2 times. The 2nd time (in FillResolve (SPClaim)), the context will always be the URL of the default zone
            foreach (var zone in Enum.GetValues(typeof(SPUrlZone)))
            {
                SPIisSettings iisSettings = webApp.GetIisSettingsWithFallback((SPUrlZone)zone);
                if (!iisSettings.UseTrustedClaimsAuthenticationProvider)
                {
                    continue;
                }

                // Get the list of authentication providers associated with the zone
                foreach (SPAuthenticationProvider prov in iisSettings.ClaimsAuthenticationProviders)
                {
                    if (prov.GetType() == typeof(Microsoft.SharePoint.Administration.SPTrustedAuthenticationProvider))
                    {
                        // Check if the current SPTrustedAuthenticationProvider is associated with the claim provider
                        if (String.Equals(prov.ClaimProviderName, ProviderInternalName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
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
            {
                return lp.First();
            }

            if (lp != null && lp.Count() > 1)
            {
                ClaimsProviderLogging.Log($"[{providerInternalName}] Cannot continue because '{providerInternalName}' is set with multiple SPTrustedIdentityTokenIssuer", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
            }
            ClaimsProviderLogging.Log($"[{providerInternalName}] Cannot continue because '{providerInternalName}' is not set with any SPTrustedIdentityTokenIssuer.\r\nVisit {ClaimsProviderConstants.PUBLICSITEURL} for more information.", TraceSeverity.High, EventSeverity.Warning, TraceCategory.Core);
            return null;
        }

        /// <summary>
        /// Uses reflection to return the value of a public property for the given object
        /// </summary>
        /// <param name="directoryObject"></param>
        /// <param name="propertyName"></param>
        /// <returns>Null if property doesn't exist, String.Empty if property exists but has no value, actual value otherwise</returns>
        public static string GetPropertyValue(object directoryObject, string propertyName)
        {
            if (directoryObject == null) { return null; }
            PropertyInfo pi = directoryObject.GetType().GetProperty(propertyName);
            if (pi == null) { return null; }   // Property doesn't exist
            object propertyValue = pi.GetValue(directoryObject, null);
            return propertyValue == null ? String.Empty : propertyValue.ToString();
        }

        /// <summary>
        /// Create a SPClaim with property OriginalIssuer correctly set
        /// </summary>
        /// <param name="type">Claim type</param>
        /// <param name="value">Claim value</param>
        /// <param name="valueType">Claim value type</param>
        /// <returns>SPClaim object</returns>
        protected virtual new SPClaim CreateClaim(string type, string value, string valueType)
        {
            // SPClaimProvider.CreateClaim sets property OriginalIssuer to SPOriginalIssuerType.ClaimProvider, which is not correct
            //return CreateClaim(type, value, valueType);
            return new SPClaim(type, value, valueType, IssuerName);
        }

        protected virtual PickerEntity CreatePickerEntityHelper(AzureCPResult result)
        {
            PickerEntity entity = CreatePickerEntity();
            SPClaim claim;
            string permissionValue = result.PermissionValue;
            string permissionClaimType = result.ClaimTypeConfig.ClaimType;
            bool isMappedClaimTypeConfig = false;

            if (String.Equals(result.ClaimTypeConfig.ClaimType, IdentityClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase)
                || result.ClaimTypeConfig.UseMainClaimTypeOfDirectoryObject)
            {
                isMappedClaimTypeConfig = true;
            }

            if (result.ClaimTypeConfig.UseMainClaimTypeOfDirectoryObject)
            {
                string claimValueType;
                if (result.ClaimTypeConfig.EntityType == DirectoryObjectType.User)
                {
                    permissionClaimType = IdentityClaimTypeConfig.ClaimType;
                    entity.EntityType = SPClaimEntityTypes.User;
                    claimValueType = IdentityClaimTypeConfig.ClaimValueType;
                }
                else
                {
                    permissionClaimType = MainGroupClaimTypeConfig.ClaimType;
                    entity.EntityType = ClaimsProviderConstants.GroupClaimEntityType;
                    claimValueType = MainGroupClaimTypeConfig.ClaimValueType;
                }
                permissionValue = FormatPermissionValue(permissionClaimType, permissionValue, isMappedClaimTypeConfig, result);
                claim = CreateClaim(
                    permissionClaimType,
                    permissionValue,
                    claimValueType);
            }
            else
            {
                permissionValue = FormatPermissionValue(permissionClaimType, permissionValue, isMappedClaimTypeConfig, result);
                claim = CreateClaim(
                    permissionClaimType,
                    permissionValue,
                    result.ClaimTypeConfig.ClaimValueType);
                entity.EntityType = result.ClaimTypeConfig.EntityType == DirectoryObjectType.User ? SPClaimEntityTypes.User : ClaimsProviderConstants.GroupClaimEntityType;
            }

            entity.Claim = claim;
            entity.IsResolved = true;
            //entity.EntityGroupName = "";
            entity.Description = String.Format(
                PickerEntityOnMouseOver,
                result.ClaimTypeConfig.DirectoryObjectProperty.ToString(),
                result.QueryMatchValue);

            int nbMetadata = 0;
            // Populate metadata of new PickerEntity
            foreach (ClaimTypeConfig ctConfig in MetadataConfig.Where(x => x.EntityType == result.ClaimTypeConfig.EntityType))
            {
                // if there is actally a value in the GraphObject, then it can be set
                string entityAttribValue = GetPropertyValue(result.UserOrGroupResult, ctConfig.DirectoryObjectProperty.ToString());
                if (!String.IsNullOrEmpty(entityAttribValue))
                {
                    entity.EntityData[ctConfig.EntityDataKey] = entityAttribValue;
                    nbMetadata++;
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Set metadata '{ctConfig.EntityDataKey}' of new entity to '{entityAttribValue}'", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
            }
            entity.DisplayText = FormatPermissionDisplayText(entity, isMappedClaimTypeConfig, result);
            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Created entity: display text: '{entity.DisplayText}', value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}', and filled with {nbMetadata.ToString()} metadata.", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
            return entity;
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
        /// <param name="entity"></param>
        /// <param name="isMappedClaimTypeConfig"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        protected virtual string FormatPermissionDisplayText(PickerEntity entity, bool isMappedClaimTypeConfig, AzureCPResult result)
        {
            string entityDisplayText = this.CurrentConfiguration.EntityDisplayTextPrefix;
            if (result.ClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText != AzureADObjectProperty.NotSet)
            {
                if (!isMappedClaimTypeConfig || result.ClaimTypeConfig.EntityType == DirectoryObjectType.Group)
                {
                    entityDisplayText += "(" + result.ClaimTypeConfig.ClaimTypeDisplayName + ") ";
                }

                string graphPropertyToDisplayValue = GetPropertyValue(result.UserOrGroupResult, result.ClaimTypeConfig.DirectoryObjectPropertyToShowAsDisplayText.ToString());
                if (!String.IsNullOrEmpty(graphPropertyToDisplayValue))
                {
                    entityDisplayText += graphPropertyToDisplayValue;
                }
                else
                {
                    entityDisplayText += result.PermissionValue;
                }
            }
            else
            {
                if (isMappedClaimTypeConfig)
                {
                    entityDisplayText += result.QueryMatchValue;
                }
                else
                {
                    entityDisplayText += String.Format(
                        PickerEntityDisplayText,
                        result.ClaimTypeConfig.ClaimTypeDisplayName,
                        result.PermissionValue);
                }
            }
            return entityDisplayText;
        }

        protected virtual PickerEntity CreatePickerEntityForSpecificClaimType(string input, ClaimTypeConfig ctConfig, bool inputHasKeyword)
        {
            List<PickerEntity> entities = CreatePickerEntityForSpecificClaimTypes(
                input,
                new List<ClaimTypeConfig>()
                    {
                        ctConfig,
                    },
                inputHasKeyword);
            return entities == null ? null : entities.First();
        }

        protected virtual List<PickerEntity> CreatePickerEntityForSpecificClaimTypes(string input, List<ClaimTypeConfig> ctConfigs, bool inputHasKeyword)
        {
            List<PickerEntity> entities = new List<PickerEntity>();
            foreach (var ctConfig in ctConfigs)
            {
                SPClaim claim = CreateClaim(ctConfig.ClaimType, input, ctConfig.ClaimValueType);
                PickerEntity entity = CreatePickerEntity();
                entity.Claim = claim;
                entity.IsResolved = true;
                entity.EntityType = ctConfig.EntityType == DirectoryObjectType.User ? SPClaimEntityTypes.User : ClaimsProviderConstants.GroupClaimEntityType;
                //entity.EntityGroupName = "";
                entity.Description = String.Format(PickerEntityOnMouseOver, ctConfig.DirectoryObjectProperty.ToString(), input);

                if (!String.IsNullOrEmpty(ctConfig.EntityDataKey))
                {
                    entity.EntityData[ctConfig.EntityDataKey] = entity.Claim.Value;
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Added metadata '{ctConfig.EntityDataKey}' with value '{entity.EntityData[ctConfig.EntityDataKey]}' to new entity", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                }

                AzureCPResult result = new AzureCPResult(null);
                result.ClaimTypeConfig = ctConfig;
                result.PermissionValue = input;
                result.QueryMatchValue = input;
                bool isIdentityClaimType = String.Equals(claim.ClaimType, IdentityClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase);
                entity.DisplayText = FormatPermissionDisplayText(entity, isIdentityClaimType, result);

                entities.Add(entity);
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Created entity: display text: '{entity.DisplayText}', value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'.", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
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
            if (claimTypes == null) { return; }
            try
            {
                this.Lock_Config.EnterReadLock();
                if (ProcessedClaimTypesList == null) { return; }
                foreach (var claimTypeSettings in ProcessedClaimTypesList)
                {
                    claimTypes.Add(claimTypeSettings.ClaimType);
                }
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in FillClaimTypes", TraceCategory.Core, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        protected override void FillClaimValueTypes(List<string> claimValueTypes)
        {
            claimValueTypes.Add(WIF4_5.ClaimValueTypes.String);
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
        protected void AugmentEntity(Uri context, SPClaim entity, SPClaimProviderContext claimProviderContext, List<SPClaim> claims)
        {
            Stopwatch timer = new Stopwatch();
            timer.Start();
            SPClaim decodedEntity;
            if (SPClaimProviderManager.IsUserIdentifierClaim(entity))
            {
                decodedEntity = SPClaimProviderManager.DecodeUserIdentifierClaim(entity);
            }
            else
            {
                if (SPClaimProviderManager.IsEncodedClaim(entity.Value))
                {
                    decodedEntity = SPClaimProviderManager.Local.DecodeClaim(entity.Value);
                }
                else
                {
                    decodedEntity = entity;
                }
            }

            SPOriginalIssuerType loginType = SPOriginalIssuers.GetIssuerType(decodedEntity.OriginalIssuer);
            if (loginType != SPOriginalIssuerType.TrustedProvider && loginType != SPOriginalIssuerType.ClaimProvider)
            {
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Not trying to augment '{decodedEntity.Value}' because his OriginalIssuer is '{decodedEntity.OriginalIssuer}'.",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Augmentation);
                return;
            }

            if (!Initialize(context, null)) { return; }

            this.Lock_Config.EnterReadLock();
            try
            {
                // There can be multiple TrustedProvider on the farm, but AzureCP should only do augmentation if current entity is from TrustedProvider it is associated with
                if (!String.Equals(decodedEntity.OriginalIssuer, IssuerName, StringComparison.InvariantCultureIgnoreCase)) { return; }

                if (!this.CurrentConfiguration.EnableAugmentation) { return; }

                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Starting augmentation for user '{decodedEntity.Value}'.", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                ClaimTypeConfig groupClaimTypeSettings = this.ProcessedClaimTypesList.FirstOrDefault(x => x.EntityType == DirectoryObjectType.Group);
                if (groupClaimTypeSettings == null)
                {
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] No claim type with EntityType 'Group' was found, please check claims mapping table.",
                        TraceSeverity.High, EventSeverity.Error, TraceCategory.Augmentation);
                    return;
                }

                OperationContext currentContext = new OperationContext(CurrentConfiguration, OperationType.Augmentation, ProcessedClaimTypesList, null, decodedEntity, context, null, null, Int32.MaxValue);
                Task<List<SPClaim>> resultsTask = GetGroupMembershipAsync(currentContext, groupClaimTypeSettings);
                resultsTask.Wait();
                List<SPClaim> groups = resultsTask.Result;
                timer.Stop();
                if (groups?.Count > 0)
                {
                    foreach (SPClaim group in groups)
                    {
                        claims.Add(group);
                        ClaimsProviderLogging.Log($"[{ProviderInternalName}] Added group '{group.Value}' to user '{currentContext.IncomingEntity.Value}'",
                            TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                    }
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] User '{currentContext.IncomingEntity.Value}' was augmented with {groups.Count.ToString()} groups in {timer.ElapsedMilliseconds.ToString()} ms",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Augmentation);
                }
                else
                {
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] No group found for user '{currentContext.IncomingEntity.Value}', search took {timer.ElapsedMilliseconds.ToString()} ms",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Augmentation);
                }
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in AugmentEntity", TraceCategory.Augmentation, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        protected async Task<List<SPClaim>> GetGroupMembershipAsync(OperationContext currentContext, ClaimTypeConfig groupClaimTypeSettings)
        {
            List<SPClaim> groups = new List<SPClaim>();

            // Create a task for each tenant to query
            // Using list CurrentConfiguration.AzureTenants directly doesn't cause thread safety issues because no property from this list is written during augmentation
            var tenantQueryTasks = this.CurrentConfiguration.AzureTenants.Select(async tenant =>
            {
                return await GetGroupMembershipFromAzureADAsync(currentContext, groupClaimTypeSettings, tenant).ConfigureAwait(false);
            });

            // Wait for all tasks to complete
            List<SPClaim>[] tenantResults = await Task.WhenAll(tenantQueryTasks).ConfigureAwait(false);

            // Process result returned by each tenant
            foreach (List<SPClaim> tenantResult in tenantResults)
            {
                if (tenantResult?.Count > 0)
                {
                    // The logic is that there will always be only 1 tenant returning groups, so as soon as 1 returned groups, foreach can stop
                    groups = tenantResult;
                    break;
                }
            }
            return groups;
        }

        protected async Task<List<SPClaim>> GetGroupMembershipFromAzureADAsync(OperationContext currentContext, ClaimTypeConfig groupClaimTypeConfig, AzureTenant tenant)
        {
            List<SPClaim> claims = new List<SPClaim>();
            // URL encode the filter to prevent that it gets truncated like this: "UserPrincipalName eq 'guest_contoso.com" instead of "UserPrincipalName eq 'guest_contoso.com#EXT#@TENANT.onmicrosoft.com'"
            string filter = HttpUtility.UrlEncode($"{currentContext.IncomingEntityClaimTypeConfig.DirectoryObjectProperty} eq '{currentContext.IncomingEntity.Value}'");

            // Do this operation in a try/catch, so if current tenant throws an exception (e.g. secret is expired), execution can still continue for other tenants
            IGraphServiceUsersCollectionPage userResult = null;
            try
            {
                // https://github.com/Yvand/AzureCP/issues/78
                // In this method, awaiting on the async task hangs in some scenario (reproduced only in multi-server 2019 farm in the w3wp of a site while using "check permissions" feature)
                // Workaround: Instead of awaiting on the async task directly, run it in a parent task, and await on the parent task.
                // userResult = await tenant.GraphService.Users.Request().Filter(filter).GetAsync().ConfigureAwait(false);
                userResult = await Task.Run(() => tenant.GraphService.Users.Request().Filter(filter).GetAsync()).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, $"on tenant '{tenant.Name}' while running query '{filter}'", TraceCategory.Lookup, ex);
                return claims;
            }

            User user = userResult.FirstOrDefault();
            if (user == null)
            {
                // If user was not found, he might be a Guest user. Query to check this: /users?$filter=userType eq 'Guest' and mail eq 'guest@live.com'&$select=userPrincipalName, Id
                string guestFilter = HttpUtility.UrlEncode($"userType eq 'Guest' and {IdentityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers} eq '{currentContext.IncomingEntity.Value}'");
                //userResult = await tenant.GraphService.Users.Request().Filter(guestFilter).Select(HttpUtility.UrlEncode("userPrincipalName, Id")).GetAsync().ConfigureAwait(false);
                userResult = await Task.Run(() => tenant.GraphService.Users.Request().Filter(guestFilter).Select(HttpUtility.UrlEncode("userPrincipalName, Id")).GetAsync()).ConfigureAwait(false);
                user = userResult.FirstOrDefault();
                if (user == null) { return claims; }
            }

            if (groupClaimTypeConfig.DirectoryObjectProperty == AzureADObjectProperty.Id)
            {
                // POST to /v1.0/users/user@TENANT.onmicrosoft.com/microsoft.graph.getMemberGroups is the preferred way to return security groups as it includes nested groups
                // But it returns only the group IDs so it can be used only if groupClaimTypeConfig.DirectoryObjectProperty == AzureADObjectProperty.Id
                // For Guest users, it must be the id: POST to /v1.0/users/18ff6ae9-dd01-4008-a786-aabf71f1492a/microsoft.graph.getMemberGroups
                //IDirectoryObjectGetMemberGroupsCollectionPage groupIDs = await tenant.GraphService.Users[user.Id].GetMemberGroups(CurrentConfiguration.FilterSecurityEnabledGroupsOnly).Request().PostAsync().ConfigureAwait(false);
                IDirectoryObjectGetMemberGroupsCollectionPage groupIDs = await Task.Run(() => tenant.GraphService.Users[user.Id].GetMemberGroups(CurrentConfiguration.FilterSecurityEnabledGroupsOnly).Request().PostAsync()).ConfigureAwait(false);
                if (groupIDs != null)
                {
                    bool morePages = groupIDs.Count > 0;
                    while (morePages)
                    {
                        foreach (string groupID in groupIDs)
                        {
                            claims.Add(CreateClaim(groupClaimTypeConfig.ClaimType, groupID, groupClaimTypeConfig.ClaimValueType));
                        }

                        if (groupIDs.NextPageRequest != null)
                        {
                            //groupIDs = await groupIDs.NextPageRequest.PostAsync().ConfigureAwait(false);
                            groupIDs = await Task.Run(() => groupIDs.NextPageRequest.PostAsync()).ConfigureAwait(false);
                        }
                        else
                        {
                            morePages = false;
                        }
                    }
                }
            }
            else
            {
                // Fallback to GET to /v1.0/users/user@TENANT.onmicrosoft.com/memberOf, which returns all group properties but does not return nested groups
                //IUserMemberOfCollectionWithReferencesPage groups = await tenant.GraphService.Users[user.Id].MemberOf.Request().GetAsync().ConfigureAwait(false);
                IUserMemberOfCollectionWithReferencesPage groups = await Task.Run(() => tenant.GraphService.Users[user.Id].MemberOf.Request().GetAsync()).ConfigureAwait(false);
                if (groups != null)
                {
                    bool morePages = groups.Count > 0;
                    while (morePages)
                    {
                        foreach (Group group in groups.OfType<Group>())
                        {
                            string groupClaimValue = GetPropertyValue(group, groupClaimTypeConfig.DirectoryObjectProperty.ToString());
                            claims.Add(CreateClaim(groupClaimTypeConfig.ClaimType, groupClaimValue, groupClaimTypeConfig.ClaimValueType));
                        }
                        if (groups.NextPageRequest != null)
                        {
                            //groups = await groups.NextPageRequest.GetAsync().ConfigureAwait(false);
                            groups = await Task.Run(() => groups.NextPageRequest.GetAsync()).ConfigureAwait(false);
                        }
                        else
                        {
                            morePages = false;
                        }
                    }
                }
            }
            return claims;
        }

        protected override void FillEntityTypes(List<string> entityTypes)
        {
            entityTypes.Add(SPClaimEntityTypes.User);
            entityTypes.Add(ClaimsProviderConstants.GroupClaimEntityType);
        }

        protected override void FillHierarchy(Uri context, string[] entityTypes, string hierarchyNodeID, int numberOfLevels, Microsoft.SharePoint.WebControls.SPProviderHierarchyTree hierarchy)
        {
            List<DirectoryObjectType> aadEntityTypes = new List<DirectoryObjectType>();
            if (entityTypes.Contains(SPClaimEntityTypes.User)) { aadEntityTypes.Add(DirectoryObjectType.User); }
            if (entityTypes.Contains(ClaimsProviderConstants.GroupClaimEntityType)) { aadEntityTypes.Add(DirectoryObjectType.Group); }

            if (!Initialize(context, entityTypes)) { return; }

            this.Lock_Config.EnterReadLock();
            try
            {
                if (hierarchyNodeID == null)
                {
                    // Root level
                    foreach (var azureObject in this.ProcessedClaimTypesList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject && aadEntityTypes.Contains(x.EntityType)))
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
                ClaimsProviderLogging.LogException(ProviderInternalName, "in FillHierarchy", TraceCategory.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        /// <summary>
        /// Override this method to change / remove entities created by AzureCP, or add new ones
        /// </summary>
        /// <param name="currentContext"></param>
        /// <param name="entityTypes"></param>
        /// <param name="input"></param>
        /// <param name="resolved">List of entities created by LDAPCP</param>
        protected virtual void FillEntities(OperationContext currentContext, ref List<PickerEntity> resolved)
        {
        }

        protected override void FillResolve(Uri context, string[] entityTypes, SPClaim resolveInput, List<Microsoft.SharePoint.WebControls.PickerEntity> resolved)
        {
            //ClaimsProviderLogging.LogDebug($"context passed to FillResolve (SPClaim): {context.ToString()}");
            if (!Initialize(context, entityTypes)) { return; }

            // Ensure incoming claim should be validated by AzureCP
            // Must be made after call to Initialize because SPTrustedLoginProvider name must be known
            if (!String.Equals(resolveInput.OriginalIssuer, IssuerName, StringComparison.InvariantCultureIgnoreCase)) { return; }

            this.Lock_Config.EnterReadLock();
            try
            {
                OperationContext currentContext = new OperationContext(CurrentConfiguration, OperationType.Validation, ProcessedClaimTypesList, resolveInput.Value, resolveInput, context, entityTypes, null, 1);
                List<PickerEntity> entities = SearchOrValidate(currentContext);
                if (entities?.Count == 1)
                {
                    resolved.Add(entities[0]);
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Validated entity: display text: '{entities[0].DisplayText}', claim value: '{entities[0].Claim.Value}', claim type: '{entities[0].Claim.ClaimType}'",
                        TraceSeverity.High, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
                else
                {
                    int entityCount = entities == null ? 0 : entities.Count;
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Validation failed: found {entityCount.ToString()} entities instead of 1 for incoming claim with value '{currentContext.IncomingEntity.Value}' and type '{currentContext.IncomingEntity.ClaimType}'", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Claims_Picking);
                }
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in FillResolve(SPClaim)", TraceCategory.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        protected override void FillResolve(Uri context, string[] entityTypes, string resolveInput, List<Microsoft.SharePoint.WebControls.PickerEntity> resolved)
        {
            if (!Initialize(context, entityTypes)) { return; }

            this.Lock_Config.EnterReadLock();
            try
            {
                OperationContext currentContext = new OperationContext(CurrentConfiguration, OperationType.Search, ProcessedClaimTypesList, resolveInput, null, context, entityTypes, null, CurrentConfiguration.MaxSearchResultsCount);
                List<PickerEntity> entities = SearchOrValidate(currentContext);
                FillEntities(currentContext, ref entities);
                if (entities == null || entities.Count == 0) { return; }
                foreach (PickerEntity entity in entities)
                {
                    resolved.Add(entity);
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Added entity: display text: '{entity.DisplayText}', claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                        TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Returned {entities.Count} entities with input '{currentContext.Input}'",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in FillResolve(string)", TraceCategory.Claims_Picking, ex);
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
            if (!Initialize(context, entityTypes)) { return; }

            this.Lock_Config.EnterReadLock();
            try
            {
                OperationContext currentContext = new OperationContext(CurrentConfiguration, OperationType.Search, ProcessedClaimTypesList, searchPattern, null, context, entityTypes, hierarchyNodeID, CurrentConfiguration.MaxSearchResultsCount);
                List<PickerEntity> entities = SearchOrValidate(currentContext);
                FillEntities(currentContext, ref entities);
                if (entities == null || entities.Count == 0) { return; }
                SPProviderHierarchyNode matchNode = null;
                foreach (PickerEntity entity in entities)
                {
                    // Add current PickerEntity to the corresponding ClaimType in the hierarchy
                    if (searchTree.HasChild(entity.Claim.ClaimType))
                    {
                        matchNode = searchTree.Children.First(x => x.HierarchyNodeID == entity.Claim.ClaimType);
                    }
                    else
                    {
                        ClaimTypeConfig ctConfig = ProcessedClaimTypesList.FirstOrDefault(x =>
                            !x.UseMainClaimTypeOfDirectoryObject &&
                            String.Equals(x.ClaimType, entity.Claim.ClaimType, StringComparison.InvariantCultureIgnoreCase));

                        string nodeName = ctConfig != null ? ctConfig.ClaimTypeDisplayName : entity.Claim.ClaimType;
                        matchNode = new SPProviderHierarchyNode(_ProviderInternalName, nodeName, entity.Claim.ClaimType, true);
                        searchTree.AddChild(matchNode);
                    }
                    matchNode.AddEntity(entity);
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Added entity: display text: '{entity.DisplayText}', claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                        TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Returned {entities.Count} entities from input '{currentContext.Input}'",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in FillSearch", TraceCategory.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        /// <summary>
        /// Search or validate incoming input or entity
        /// </summary>
        /// <param name="currentContext">Information about current context and operation</param>
        /// <returns>Entities generated by AzureCP</returns>
        protected List<PickerEntity> SearchOrValidate(OperationContext currentContext)
        {
            List<PickerEntity> entities = new List<PickerEntity>();
            try
            {
                if (this.CurrentConfiguration.AlwaysResolveUserInput)
                {
                    // Completely bypass query to Azure AD
                    entities = CreatePickerEntityForSpecificClaimTypes(
                        currentContext.Input,
                        currentContext.CurrentClaimTypeConfigList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject),
                        false);
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Created {entities.Count} entity(ies) without contacting Azure AD tenant(s) because AzureCP property AlwaysResolveUserInput is set to true.",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
                    return entities;
                }

                if (currentContext.OperationType == OperationType.Search)
                {
                    entities = SearchOrValidateInAzureAD(currentContext);

                    // Check if input starts with a prefix configured on a ClaimTypeConfig. If so an entity should be returned using ClaimTypeConfig found
                    // ClaimTypeConfigEnsureUniquePrefixToBypassLookup ensures that collection cannot contain duplicates
                    ClaimTypeConfig ctConfigWithInputPrefixMatch = currentContext.CurrentClaimTypeConfigList.FirstOrDefault(x =>
                        !String.IsNullOrEmpty(x.PrefixToBypassLookup) &&
                        currentContext.Input.StartsWith(x.PrefixToBypassLookup, StringComparison.InvariantCultureIgnoreCase));
                    if (ctConfigWithInputPrefixMatch != null)
                    {
                        string inputWithoutPrefix = currentContext.Input.Substring(ctConfigWithInputPrefixMatch.PrefixToBypassLookup.Length);
                        if (String.IsNullOrEmpty(inputWithoutPrefix))
                        {
                            // No value in the input after the prefix, return
                            return entities;
                        }
                        PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                            inputWithoutPrefix,
                            ctConfigWithInputPrefixMatch,
                            true);
                        if (entity != null)
                        {
                            if (entities == null) { entities = new List<PickerEntity>(); }
                            entities.Add(entity);
                            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Created entity without contacting Azure AD tenant(s) because input started with prefix '{ctConfigWithInputPrefixMatch.PrefixToBypassLookup}', which is configured for claim type '{ctConfigWithInputPrefixMatch.ClaimType}'. Claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                            //return entities;
                        }
                    }
                }
                else if (currentContext.OperationType == OperationType.Validation)
                {
                    entities = SearchOrValidateInAzureAD(currentContext);
                    if (entities?.Count == 1) { return entities; }

                    if (!String.IsNullOrEmpty(currentContext.IncomingEntityClaimTypeConfig.PrefixToBypassLookup))
                    {
                        // At this stage, it is impossible to know if entity was originally created with the keyword that bypass query to Azure AD
                        // But it should be always validated since property PrefixToBypassLookup is set for current ClaimTypeConfig, so create entity manually
                        PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                            currentContext.IncomingEntity.Value,
                            currentContext.IncomingEntityClaimTypeConfig,
                            currentContext.InputHasKeyword);
                        if (entity != null)
                        {
                            entities = new List<PickerEntity>(1) { entity };
                            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Validated entity without contacting Azure AD tenant(s) because its claim type ('{currentContext.IncomingEntityClaimTypeConfig.ClaimType}') has property 'PrefixToBypassLookup' set in AzureCPConfig.ClaimTypes. Claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in SearchOrValidate", TraceCategory.Claims_Picking, ex);
            }
            entities = this.ValidateEntities(currentContext, entities);
            return entities;
        }

        protected List<PickerEntity> SearchOrValidateInAzureAD(OperationContext currentContext)
        {
            string userFilter = String.Empty;
            string groupFilter = String.Empty;
            string userSelect = String.Empty;
            string groupSelect = String.Empty;

            // BUG: Filters must be set in an object created in this method (to be bound to current thread), otherwise filter may be updated by multiple threads
            List<AzureTenant> azureTenants = new List<AzureTenant>(this.CurrentConfiguration.AzureTenants.Count);
            foreach (AzureTenant tenant in this.CurrentConfiguration.AzureTenants)
            {
                azureTenants.Add(tenant.CopyConfiguration());
            }

            BuildFilter(currentContext, azureTenants);

            List<AzureADResult> aadResults = null;
            using (new SPMonitoredScope($"[{ProviderInternalName}] Total time spent to query Azure AD tenant(s)", 1000))
            {
                // Call async method in a task to avoid error "Asynchronous operations are not allowed in this context" error when permission is validated (POST from people picker)
                // More info on the error: https://stackoverflow.com/questions/672237/running-an-asynchronous-operation-triggered-by-an-asp-net-web-page-request
                Task azureADQueryTask = Task.Run(async () =>
                {
                    aadResults = await QueryAzureADTenantsAsync(currentContext, azureTenants).ConfigureAwait(false);
                });
                azureADQueryTask.Wait();
            }

            if (aadResults == null || aadResults.Count <= 0) { return null; }
            List<AzureCPResult> results = ProcessAzureADResults(currentContext, aadResults);
            if (results == null || results.Count <= 0) { return null; }
            List<PickerEntity> entities = new List<PickerEntity>();
            foreach (var result in results)
            {
                entities.Add(result.PickerEntity);
                //ClaimsProviderLogging.Log($"[{ProviderInternalName}] Added entity returned by Azure AD: claim value: '{result.PickerEntity.Claim.Value}', claim type: '{result.PickerEntity.Claim.ClaimType}'",
                //    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
            }
            return entities;
        }

        /// <summary>
        /// Override this method to inspect the entities generated by AzureCP, and remove some before they are returned to SharePoint.
        /// </summary>
        /// <param name="entities">Entities generated by AzureCP</param>
        /// <returns>List of entities that AzureCP will return to SharePoint</returns>
        protected virtual List<PickerEntity> ValidateEntities(OperationContext currentContext, List<PickerEntity> entities)
        {
            return entities;
        }

        /// <summary>
        /// Build filter and select statements used in queries sent to Azure AD
        /// $filter and $select must be URL encoded as documented in https://developer.microsoft.com/en-us/graph/docs/concepts/query_parameters#encoding-query-parameters
        /// </summary>
        /// <param name="currentContext"></param>
        protected virtual void BuildFilter(OperationContext currentContext, List<AzureTenant> azureTenants)
        {
            string searchPatternEquals = "{0} eq '{1}'";
            string searchPatternStartsWith = "startswith({0}, '{1}')";
            string identityConfigSearchPatternEquals = "({0} eq '{1}' and UserType eq '{2}')";
            string identityConfigSearchPatternStartsWith = "(startswith({0}, '{1}') and UserType eq '{2}')";

            StringBuilder userFilterBuilder = new StringBuilder();
            StringBuilder groupFilterBuilder = new StringBuilder();
            StringBuilder userSelectBuilder = new StringBuilder("UserType, Mail, ");    // UserType and Mail are always needed to deal with Guest users
            StringBuilder groupSelectBuilder = new StringBuilder("Id, securityEnabled, ");               // Id is always required for groups

            //// Microsoft Graph doesn't support operator not equals (ne) on attribute UserType, it can only be queried using equals (eq)
            //string memberOnlyUserTypeFilter = " and UserType eq 'Member'";
            //string guestOnlyUserTypeFilter = " and UserType eq 'Guest'";

            string preferredFilterPattern;
            string input = currentContext.Input;

            // https://github.com/Yvand/AzureCP/issues/88: Escape single quotes as documented in https://docs.microsoft.com/en-us/graph/query-parameters#escaping-single-quotes
            input = input.Replace("'", "''");

            if (currentContext.ExactSearch)
            {
                preferredFilterPattern = String.Format(searchPatternEquals, "{0}", input);
            }
            else
            {
                preferredFilterPattern = String.Format(searchPatternStartsWith, "{0}", input);
            }

            bool firstUserObjectProcessed = false;
            bool firstGroupObjectProcessed = false;
            foreach (ClaimTypeConfig ctConfig in currentContext.CurrentClaimTypeConfigList)
            {
                string currentPropertyString = ctConfig.DirectoryObjectProperty.ToString();
                string currentFilter;
                if (!ctConfig.SupportsWildcard)
                {
                    currentFilter = String.Format(searchPatternEquals, currentPropertyString, input);
                }
                else
                {
                    // Use String.Replace instead of String.Format because String.Format trows an exception if input contains a '{'
                    //currentFilter = String.Format(preferredFilterPattern, currentPropertyString);
                    currentFilter = preferredFilterPattern.Replace("{0}", currentPropertyString);
                }

                // Id needs a specific check: input must be a valid GUID AND equals filter must be used, otherwise Azure AD will throw an error
                if (ctConfig.DirectoryObjectProperty == AzureADObjectProperty.Id)
                {
                    Guid idGuid = new Guid();
                    if (!Guid.TryParse(input, out idGuid))
                    {
                        continue;
                    }
                    else
                    {
                        currentFilter = String.Format(searchPatternEquals, currentPropertyString, idGuid.ToString());
                    }
                }

                if (ctConfig.EntityType == DirectoryObjectType.User)
                {
                    if (ctConfig is IdentityClaimTypeConfig)
                    {
                        IdentityClaimTypeConfig identityClaimTypeConfig = ctConfig as IdentityClaimTypeConfig;
                        if (!ctConfig.SupportsWildcard)
                        {
                            currentFilter = "( " + String.Format(identityConfigSearchPatternEquals, currentPropertyString, input, AzureADUserTypeHelper.MemberUserType) + " or " + String.Format(identityConfigSearchPatternEquals, identityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers, input, AzureADUserTypeHelper.GuestUserType) + " )";
                        }
                        else
                        {
                            if (currentContext.ExactSearch)
                            {
                                currentFilter = "( " + String.Format(identityConfigSearchPatternEquals, currentPropertyString, input, AzureADUserTypeHelper.MemberUserType) + " or " + String.Format(identityConfigSearchPatternEquals, identityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers, input, AzureADUserTypeHelper.GuestUserType) + " )";
                            }
                            else
                            {
                                currentFilter = "( " + String.Format(identityConfigSearchPatternStartsWith, currentPropertyString, input, AzureADUserTypeHelper.MemberUserType) + " or " + String.Format(identityConfigSearchPatternStartsWith, identityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers, input, AzureADUserTypeHelper.GuestUserType) + " )";
                            }
                        }
                    }

                    if (!firstUserObjectProcessed)
                    {
                        firstUserObjectProcessed = true;
                    }
                    else
                    {
                        currentFilter = " or " + currentFilter;
                        currentPropertyString = ", " + currentPropertyString;
                    }
                    userFilterBuilder.Append(currentFilter);
                    userSelectBuilder.Append(currentPropertyString);
                }
                else
                {
                    // else assume it's a Group
                    if (!firstGroupObjectProcessed)
                    {
                        firstGroupObjectProcessed = true;
                    }
                    else
                    {
                        currentFilter = " or " + currentFilter;
                        currentPropertyString = ", " + currentPropertyString;
                    }
                    groupFilterBuilder.Append(currentFilter);
                    groupSelectBuilder.Append(currentPropertyString);
                }
            }

            // Also add metadata properties to $select of corresponding object type
            if (firstUserObjectProcessed)
            {
                foreach (ClaimTypeConfig ctConfig in MetadataConfig.Where(x => x.EntityType == DirectoryObjectType.User))
                {
                    userSelectBuilder.Append($", {ctConfig.DirectoryObjectProperty.ToString()}");
                }
            }
            if (firstGroupObjectProcessed)
            {
                foreach (ClaimTypeConfig ctConfig in MetadataConfig.Where(x => x.EntityType == DirectoryObjectType.Group))
                {
                    groupSelectBuilder.Append($", {ctConfig.DirectoryObjectProperty.ToString()}");
                }
            }

            //userFilterBuilder.Append(" ) and accountEnabled eq true");  // Graph throws this error if used: "Search filter expression has excessive height: 4. Max allowed: 3."
            string encodedUserFilter = HttpUtility.UrlEncode(userFilterBuilder.ToString());
            string encodedGroupFilter = HttpUtility.UrlEncode(groupFilterBuilder.ToString());
            string encodedUserSelect = HttpUtility.UrlEncode(userSelectBuilder.ToString());
            string encodedgroupSelect = HttpUtility.UrlEncode(groupSelectBuilder.ToString());
            //string encodedMemberOnlyUserTypeFilter = HttpUtility.UrlEncode(memberOnlyUserTypeFilter);
            //string encodedGuestOnlyUserTypeFilter = HttpUtility.UrlEncode(guestOnlyUserTypeFilter);

            foreach (AzureTenant tenant in azureTenants)
            {
                if (firstUserObjectProcessed)
                {
                    tenant.UserFilter = encodedUserFilter;
                    //if (tenant.MemberUserTypeOnly)
                    //    tenant.UserFilter += encodedMemberOnlyUserTypeFilter;
                    //else if (tenant.ExcludeGuestUsers)
                    //    tenant.UserFilter += encodedGuestOnlyUserTypeFilter;
                }
                else
                {
                    // Reset filters if no corresponding object was found in requestInfo.ClaimTypeConfigList, to detect that tenant should not be queried
                    tenant.UserFilter = String.Empty;
                }

                if (firstGroupObjectProcessed)
                {
                    tenant.GroupFilter = encodedGroupFilter;
                }
                else
                {
                    tenant.GroupFilter = String.Empty;
                }

                tenant.UserSelect = encodedUserSelect;
                tenant.GroupSelect = encodedgroupSelect;
            }
        }

        protected async Task<List<AzureADResult>> QueryAzureADTenantsAsync(OperationContext currentContext, List<AzureTenant> azureTenants)
        {
            // Create a task for each tenant to query
            var tenantQueryTasks = azureTenants.Select(async tenant =>
            {
                Stopwatch timer = new Stopwatch();
                AzureADResult tenantResult = null;
                try
                {
                    timer.Start();
                    tenantResult = await QueryAzureADTenantAsync(currentContext, tenant, true).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    ClaimsProviderLogging.LogException(ProviderInternalName, $"in QueryAzureADTenantsAsync while querying tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
                }
                finally
                {
                    timer.Stop();
                }
                if (tenantResult != null)
                {
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Got {tenantResult.UsersAndGroups.Count().ToString()} users/groups in {timer.ElapsedMilliseconds.ToString()} ms from '{tenant.Name}' with input '{currentContext.Input}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                }
                else
                {
                    ClaimsProviderLogging.Log($"[{ProviderInternalName}] Got no result from '{tenant.Name}' with input '{currentContext.Input}', search took {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                }
                return tenantResult;
            });

            // Wait for all tasks to complete and return result as a List<AzureADResult>
            AzureADResult[] tenantResults = await Task.WhenAll(tenantQueryTasks).ConfigureAwait(false);
            return tenantResults.ToList();
        }

        protected virtual async Task<AzureADResult> QueryAzureADTenantAsync(OperationContext currentContext, AzureTenant tenant, bool firstAttempt)
        {
            AzureADResult tenantResults = new AzureADResult();
            if (String.IsNullOrWhiteSpace(tenant.UserFilter) && String.IsNullOrWhiteSpace(tenant.GroupFilter))
            {
                return tenantResults;
            }

            if (tenant.GraphService == null)
            {
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Cannot query Azure AD tenant '{tenant.Name}' because it was not initialized", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
                return tenantResults;
            }

            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Querying Azure AD tenant '{tenant.Name}' for users and groups, with input '{currentContext.Input}'", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Lookup);
            bool tryAgain = false;
            object lockAddResultToCollection = new object();
            int timeout = this.CurrentConfiguration.Timeout;
#if DEBUG
            timeout = 60 * 1000;
#endif
            CancellationTokenSource cts = new CancellationTokenSource(timeout);
            try
            {
                using (new SPMonitoredScope($"[{ProviderInternalName}] Querying Azure AD tenant '{tenant.Name}' for users and groups, with input '{currentContext.Input}'", 1000))
                {
                    // Run in a task to timeout it if it takes too long
                    Task batchQueryTask = Task.Run(async () =>
                    {
                        // Initialize requests and variables that will receive the result
                        IGraphServiceUsersCollectionPage usersFound = null;
                        IGraphServiceGroupsCollectionPage groupsFound = null;
                        IGraphServiceUsersCollectionRequest userRequest = tenant.GraphService.Users.Request().Select(tenant.UserSelect).Filter(tenant.UserFilter).Top(currentContext.MaxCount);
                        IGraphServiceGroupsCollectionRequest groupRequest = tenant.GraphService.Groups.Request().Select(tenant.GroupSelect).Filter(tenant.GroupFilter).Top(currentContext.MaxCount);

                        // Do a batch query only if necessary
                        if (!String.IsNullOrWhiteSpace(tenant.UserFilter) && !String.IsNullOrWhiteSpace(tenant.GroupFilter))
                        {
                            // https://docs.microsoft.com/en-us/graph/sdks/batch-requests?tabs=csharp
                            BatchRequestContent batchRequestContent = new BatchRequestContent();
                            var userRequestId = batchRequestContent.AddBatchRequestStep(userRequest);
                            var groupRequestId = batchRequestContent.AddBatchRequestStep(groupRequest);
                            var batchResponse = await tenant.GraphService.Batch.Request().PostAsync(batchRequestContent).ConfigureAwait(false);

                            // De - serialize batch response based on known return type
                            GraphServiceUsersCollectionResponse usersBatchResponse = await batchResponse.GetResponseByIdAsync<GraphServiceUsersCollectionResponse>(userRequestId).ConfigureAwait(false);
                            usersFound = usersBatchResponse.Value;
                            GraphServiceGroupsCollectionResponse groupsBatchResponse = await batchResponse.GetResponseByIdAsync<GraphServiceGroupsCollectionResponse>(groupRequestId).ConfigureAwait(false);
                            groupsFound = groupsBatchResponse.Value;
                        }
                        else
                        {
                            // The request only asks for either users or groups, not both
                            if (!String.IsNullOrWhiteSpace(tenant.UserFilter))
                            {
                                usersFound = await userRequest.GetAsync().ConfigureAwait(false);
                            }
                            else
                            {
                                groupsFound = await groupRequest.GetAsync().ConfigureAwait(false);
                            }
                        }

                        Task userQueryTask = Task.Run(async () =>
                        {
                            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Query to tenant '{tenant.Name}' returned {(usersFound == null ? 0 : usersFound.Count)} user(s) with filter \"{HttpUtility.UrlDecode(tenant.UserFilter)}\"", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Lookup);
                            if (usersFound != null && usersFound.Count > 0)
                            {
                                do
                                {
                                    lock (lockAddResultToCollection)
                                    {
                                        IList<User> usersInCurrentPage = usersFound.CurrentPage;
                                        if (tenant.ExcludeMembers)
                                        {
                                            usersInCurrentPage = usersFound.CurrentPage.Where(x => !String.Equals(x.UserType, ClaimsProviderConstants.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase)).ToList<User>();
                                        }
                                        else if (tenant.ExcludeGuests)
                                        {
                                            usersInCurrentPage = usersFound.CurrentPage.Where(x => !String.Equals(x.UserType, ClaimsProviderConstants.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase)).ToList<User>();
                                        }
                                        tenantResults.UsersAndGroups.AddRange(usersInCurrentPage);
                                    }
                                    if (usersFound.NextPageRequest != null)
                                    {
                                        usersFound = await usersFound.NextPageRequest.GetAsync().ConfigureAwait(false);
                                    }
                                }
                                while (usersFound.Count > 0 && usersFound.NextPageRequest != null);
                            }
                        }, cts.Token);

                        Task groupQueryTask = Task.Run(async () =>
                        {
                            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Query to tenant '{tenant.Name}' returned {(groupsFound == null ? 0 : groupsFound.Count)} group(s) with filter \"{HttpUtility.UrlDecode(tenant.GroupFilter)}\"", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Lookup);
                            if (groupsFound != null && groupsFound.Count > 0)
                            {
                                do
                                {
                                    lock (lockAddResultToCollection)
                                    {
                                        tenantResults.UsersAndGroups.AddRange(groupsFound.CurrentPage);
                                    }
                                    if (groupsFound.NextPageRequest != null)
                                    {
                                        groupsFound = await groupsFound.NextPageRequest.GetAsync().ConfigureAwait(false);
                                    }
                                }
                                while (groupsFound.Count > 0 && groupsFound.NextPageRequest != null);
                            }
                        }, cts.Token);
                        Task.WaitAll(new Task[2] { userQueryTask, groupQueryTask }, timeout, cts.Token);
                    }, cts.Token);

                    // Waits for all tasks to complete execution within a specified number of milliseconds
                    ClaimsProviderLogging.LogDebug($"Waiting on Task.WaitAll for {tenant.Name} starting");
                    // Cannot use Task.WaitAll() because it's actually blocking the threads, preventing parallel queries on others AAD tenants.
                    // Use await Task.WhenAll() as it does not block other threads, so all AAD tenants are actually queried in parallel.
                    // More info: https://stackoverflow.com/questions/12337671/using-async-await-for-multiple-tasks
                    await Task.WhenAll(new Task[1] { batchQueryTask }).ConfigureAwait(false);
                    ClaimsProviderLogging.LogDebug($"Waiting on Task.WaitAll for {tenant.Name} finished");
                }
            }
            catch (OperationCanceledException)
            {
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Queries on Azure AD tenant '{tenant.Name}' exceeded timeout of {timeout} ms and were cancelled.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
                tryAgain = true;
            }
            catch (ServiceException ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, $"Microsoft.Graph could not query tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
                tryAgain = true;
            }
            catch (AggregateException ex)
            {
                // Task.WaitAll throws an AggregateException, which contains all exceptions thrown by tasks it waited on
                ClaimsProviderLogging.LogException(ProviderInternalName, $"while querying Azure AD tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
                tryAgain = true;
            }
            finally
            {
                cts.Dispose();
            }

            if (tryAgain && !CurrentConfiguration.EnableRetry) { tryAgain = false; }

            if (firstAttempt && tryAgain)
            {
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Doing new attempt to query Azure AD tenant '{tenant.Name}'...",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                tenantResults = await QueryAzureADTenantAsync(currentContext, tenant, false).ConfigureAwait(false);
            }
            return tenantResults;
        }

        protected virtual List<AzureCPResult> ProcessAzureADResults(OperationContext currentContext, List<AzureADResult> azureADResults)
        {
            // Split results between users/groups and list of registered domains in the tenant
            List<DirectoryObject> usersAndGroups = new List<DirectoryObject>();
            // For each Azure AD tenant where list of result (UsersAndGroups) is not null
            // singleTenantResults in azureADResults can be null if AzureCP failed to get a valid access token for it
            foreach (AzureADResult singleTenantResults in azureADResults.Where(singleTenantResults => singleTenantResults != null && singleTenantResults.UsersAndGroups != null))
            {
                usersAndGroups.AddRange(singleTenantResults.UsersAndGroups);
                //domains.AddRange(tenantResults.DomainsRegisteredInAzureADTenant);
            }

            // Return if no user / groups is found, or if no registered domain is found
            if (usersAndGroups == null || !usersAndGroups.Any() /*|| domains == null || !domains.Any()*/)
            {
                return null;
            };

            List<ClaimTypeConfig> ctConfigs = currentContext.CurrentClaimTypeConfigList;
            if (currentContext.ExactSearch)
            {
                ctConfigs = currentContext.CurrentClaimTypeConfigList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject);
            }

            List<AzureCPResult> processedResults = new List<AzureCPResult>();
            foreach (DirectoryObject userOrGroup in usersAndGroups)
            {
                DirectoryObject currentObject = null;
                DirectoryObjectType objectType;
                if (userOrGroup is User)
                {
                    // This section has become irrelevant since the specific handling of guest users is done lower in the filtering, introduced in v13
                    //// Always exclude shadow users: UserType is Guest and his mail matches a verified domain in any Azure AD tenant
                    //string userType = ((User)userOrGroup).UserType;
                    //if (String.Equals(userType, AzureADUserTypeHelper.GuestUserType, StringComparison.InvariantCultureIgnoreCase))
                    //{
                    //    string userMail = ((User)userOrGroup).Mail;
                    //    if (String.IsNullOrEmpty(userMail))
                    //    {
                    //        ClaimsProviderLogging.Log($"[{ProviderInternalName}] Guest user '{((User)userOrGroup).UserPrincipalName}' filtered out because his mail is empty.",
                    //            TraceSeverity.Unexpected, EventSeverity.Warning, TraceCategory.Lookup);
                    //        continue;
                    //    }
                    //    if (!userMail.Contains('@')) continue;
                    //    string maildomain = userMail.Split('@')[1];
                    //    if (domains.Any(x => String.Equals(x, maildomain, StringComparison.InvariantCultureIgnoreCase)))
                    //    {
                    //        ClaimsProviderLogging.Log($"[{ProviderInternalName}] Guest user '{((User)userOrGroup).UserPrincipalName}' filtered out because his email '{userMail}' matches a domain registered in a Azure AD tenant.",
                    //            TraceSeverity.Verbose, EventSeverity.Verbose, TraceCategory.Lookup);
                    //        continue;
                    //    }
                    //}
                    currentObject = userOrGroup;
                    objectType = DirectoryObjectType.User;
                }
                else
                {
                    currentObject = userOrGroup;
                    objectType = DirectoryObjectType.Group;

                    if (CurrentConfiguration.FilterSecurityEnabledGroupsOnly)
                    {
                        Group group = (Group)userOrGroup;
                        // If Group.SecurityEnabled is not set, assume the group is not SecurityEnabled - verified per tests, it is not documentated in https://docs.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0
                        bool isSecurityEnabled = group.SecurityEnabled ?? false;
                        if (!isSecurityEnabled)
                        {
                            continue;
                        }
                    }
                }

                foreach (ClaimTypeConfig ctConfig in ctConfigs.Where(x => x.EntityType == objectType))
                {
                    // Get value with of current GraphProperty
                    string directoryObjectPropertyValue = GetPropertyValue(currentObject, ctConfig.DirectoryObjectProperty.ToString());

                    if (ctConfig is IdentityClaimTypeConfig)
                    {
                        if (String.Equals(((User)currentObject).UserType, AzureADUserTypeHelper.GuestUserType, StringComparison.InvariantCultureIgnoreCase))
                        {
                            // For Guest users, use the value set in property DirectoryObjectPropertyForGuestUsers
                            directoryObjectPropertyValue = GetPropertyValue(currentObject, ((IdentityClaimTypeConfig)ctConfig).DirectoryObjectPropertyForGuestUsers.ToString());
                        }
                    }

                    // Check if property exists (not null) and has a value (not String.Empty)
                    if (String.IsNullOrEmpty(directoryObjectPropertyValue)) { continue; }

                    // Check if current value mathes input, otherwise go to next GraphProperty to check
                    if (currentContext.ExactSearch)
                    {
                        if (!String.Equals(directoryObjectPropertyValue, currentContext.Input, StringComparison.InvariantCultureIgnoreCase)) { continue; }
                    }
                    else
                    {
                        if (!directoryObjectPropertyValue.StartsWith(currentContext.Input, StringComparison.InvariantCultureIgnoreCase)) { continue; }
                    }

                    // Current DirectoryObjectProperty value matches user input. Add current result to search results if it is not already present
                    string entityClaimValue = directoryObjectPropertyValue;
                    ClaimTypeConfig claimTypeConfigToCompare;
                    if (ctConfig.UseMainClaimTypeOfDirectoryObject)
                    {
                        if (objectType == DirectoryObjectType.User)
                        {
                            claimTypeConfigToCompare = IdentityClaimTypeConfig;
                            if (String.Equals(((User)currentObject).UserType, AzureADUserTypeHelper.GuestUserType, StringComparison.InvariantCultureIgnoreCase))
                            {
                                // For Guest users, use the value set in property DirectoryObjectPropertyForGuestUsers
                                entityClaimValue = GetPropertyValue(currentObject, IdentityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers.ToString());
                            }
                            else
                            {
                                // Get the value of the DirectoryObjectProperty linked to current directory object
                                entityClaimValue = GetPropertyValue(currentObject, IdentityClaimTypeConfig.DirectoryObjectProperty.ToString());
                            }
                        }
                        else
                        {
                            claimTypeConfigToCompare = MainGroupClaimTypeConfig;
                            // Get the value of the DirectoryObjectProperty linked to current directory object
                            entityClaimValue = GetPropertyValue(currentObject, claimTypeConfigToCompare.DirectoryObjectProperty.ToString());
                        }

                        if (String.IsNullOrEmpty(entityClaimValue)) { continue; }
                    }
                    else
                    {
                        claimTypeConfigToCompare = ctConfig;
                    }

                    // if claim type and claim value already exists, skip
                    bool resultAlreadyExists = processedResults.Exists(x =>
                        String.Equals(x.ClaimTypeConfig.ClaimType, claimTypeConfigToCompare.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        String.Equals(x.PermissionValue, entityClaimValue, StringComparison.InvariantCultureIgnoreCase));
                    if (resultAlreadyExists) { continue; }

                    // Passed the checks, add it to the processedResults list
                    processedResults.Add(
                        new AzureCPResult(currentObject)
                        {
                            ClaimTypeConfig = ctConfig,
                            PermissionValue = entityClaimValue,
                            QueryMatchValue = directoryObjectPropertyValue,
                        });
                }
            }

            ClaimsProviderLogging.Log($"[{ProviderInternalName}] {processedResults.Count} entity(ies) to create after filtering", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Lookup);
            foreach (AzureCPResult result in processedResults)
            {
                PickerEntity pe = CreatePickerEntityHelper(result);
                result.PickerEntity = pe;
            }
            return processedResults;
        }

        public override string Name => ProviderInternalName;
        public override bool SupportsEntityInformation => true;
        public override bool SupportsHierarchy => true;
        public override bool SupportsResolve => true;
        public override bool SupportsSearch => true;
        public override bool SupportsUserKey => true;

        /// <summary>
        /// Return the identity claim type
        /// </summary>
        /// <returns></returns>
        public override string GetClaimTypeForUserKey()
        {
            // Initialization may fail because there is no yet configuration (fresh install)
            // In this case, AzureCP should not return null because it causes null exceptions in SharePoint when users sign-in
            Initialize(null, null);

            this.Lock_Config.EnterReadLock();
            try
            {
                if (SPTrust == null)
                {
                    return String.Empty;
                }

                return SPTrust.IdentityClaimTypeInformation.MappedClaimType;
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in GetClaimTypeForUserKey", TraceCategory.Rehydration, ex);
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
            // Initialization may fail because there is no yet configuration (fresh install)
            // In this case, AzureCP should not return null because it causes null exceptions in SharePoint when users sign-in
            bool initSucceeded = Initialize(null, null);

            this.Lock_Config.EnterReadLock();
            try
            {
                // If initialization failed but SPTrust is not null, rest of the method can be executed normally
                // Otherwise return the entity
                if (!initSucceeded && SPTrust == null)
                {
                    return entity;
                }

                // There are 2 scenarios:
                // 1: OriginalIssuer is "SecurityTokenService": Value looks like "05.t|yvanhost|yvand@yvanhost.local", claim type is "http://schemas.microsoft.com/sharepoint/2009/08/claims/userid" and it must be decoded properly
                // 2: OriginalIssuer is AzureCP: in this case incoming entity is valid and returned as is
                if (String.Equals(entity.OriginalIssuer, IssuerName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return entity;
                }

                SPClaimProviderManager cpm = SPClaimProviderManager.Local;
                SPClaim curUser = SPClaimProviderManager.DecodeUserIdentifierClaim(entity);

                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Returning user key for '{entity.Value}'",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Rehydration);
                return CreateClaim(SPTrust.IdentityClaimTypeInformation.MappedClaimType, curUser.Value, curUser.ValueType);
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "in GetUserKeyForEntity", TraceCategory.Rehydration, ex);
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
        public List<DirectoryObject> UsersAndGroups
        {
            get => _UsersAndGroups;
            set => _UsersAndGroups = value;
        }
        private List<DirectoryObject> _UsersAndGroups;

        public List<string> DomainsRegisteredInAzureADTenant
        {
            get => _DomainsRegisteredInAzureADTenant;
            set => _DomainsRegisteredInAzureADTenant = value;
        }
        private List<string> _DomainsRegisteredInAzureADTenant;
        //public string TenantName;

        public AzureADResult()
        {
            UsersAndGroups = new List<DirectoryObject>();
            DomainsRegisteredInAzureADTenant = new List<string>();
            //this.TenantName = tenantName;
        }
    }

    /// <summary>
    /// User / group found in Azure AD, with additional information
    /// </summary>
    public class AzureCPResult
    {
        public readonly DirectoryObject UserOrGroupResult;
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
