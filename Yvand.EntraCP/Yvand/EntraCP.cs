using Microsoft.Graph.Models;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Yvand.EntraID;
using Yvand.Config;
using WIF4_5 = System.Security.Claims;

namespace Yvand
{
    public interface IEntraCPSettings : IEntraSettings
    {
        List<ClaimTypeConfig> RuntimeClaimTypesList { get; }
        IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; }
        IdentityClaimTypeConfig IdentityClaimTypeConfig { get; }
        ClaimTypeConfig MainGroupClaimTypeConfig { get; }
    }

    public class EntraCPSettings : EntraProviderSettings, IEntraCPSettings
    {
        public List<ClaimTypeConfig> RuntimeClaimTypesList { get; set; }

        public IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; set; }

        public IdentityClaimTypeConfig IdentityClaimTypeConfig { get; set; }

        public ClaimTypeConfig MainGroupClaimTypeConfig { get; set; }
    }

    public class EntraCP : SPClaimProvider
    {
        public static string ClaimsProviderName => "EntraCP";
        public override string Name => ClaimsProviderName;
        public override bool SupportsEntityInformation => true;
        public override bool SupportsHierarchy => true;
        public override bool SupportsResolve => true;
        public override bool SupportsSearch => true;
        public override bool SupportsUserKey => true;
        public EntraIDEntityProvider EntityProvider { get; private set; }
        private ReaderWriterLockSlim Lock_LocalConfigurationRefresh = new ReaderWriterLockSlim();
        protected virtual string PickerEntityDisplayText => "({0}) {1}";
        protected virtual string PickerEntityOnMouseOver => "{0}={1}";
        public IEntraCPSettings Settings { get; protected set; }
        public long SettingsVersion { get; private set; } = -1;
        #region "Runtime settings"
        //protected List<ClaimTypeConfig> RuntimeClaimTypesList { get; private set; }
        //protected IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; private set; }
        //protected IdentityClaimTypeConfig IdentityClaimTypeConfig { get; private set; }
        //protected ClaimTypeConfig MainGroupClaimTypeConfig { get; private set; }
        private SPTrustedLoginProvider _SPTrust;
        /// <summary>
        /// Gets the SharePoint trust that has its property ClaimProviderName equals to <see cref="Name"/>
        /// </summary>
        private SPTrustedLoginProvider SPTrust
        {
            get
            {
                if (this._SPTrust == null)
                {
                    this._SPTrust = Utils.GetSPTrustAssociatedWithClaimsProvider(this.Name);
                }
                return this._SPTrust;
            }
        }
        #endregion

        /// <summary>
        /// Gets the issuer formatted to be like the property SPClaim.OriginalIssuer: "TrustedProvider:TrustedProviderName"
        /// </summary>
        public string OriginalIssuerName => this.SPTrust != null ? SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, this.SPTrust.Name) : String.Empty;

        public EntraCP(string displayName) : base(displayName)
        {
            this.EntityProvider = new EntraIDEntityProvider(Name);
        }

        public static EntraProviderConfig<IEntraSettings> GetConfiguration(bool initializeLocalConfiguration = false)
        {
            EntraProviderConfig<IEntraSettings> configuration = (EntraProviderConfig<IEntraSettings>)EntraProviderConfig<IEntraSettings>.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID), initializeLocalConfiguration);
            return configuration;
        }

        /// <summary>
        /// Creates a configuration for EntraCP. This will delete any existing configuration which may already exist
        /// </summary>
        /// <returns></returns>
        public static EntraProviderConfig<IEntraSettings> CreateConfiguration()
        {
            EntraProviderConfig<IEntraSettings> configuration = (EntraProviderConfig<IEntraSettings>)EntraProviderConfig<IEntraSettings>.CreateGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID), ClaimsProviderConstants.CONFIGURATION_NAME, EntraCP.ClaimsProviderName, typeof(EntraProviderConfig<IEntraSettings>));
            return configuration;
        }

        /// <summary>
        /// Deletes the configuration for EntraCP
        /// </summary>
        public static void DeleteConfiguration()
        {
            EntraProviderConfig<IEntraSettings> configuration = (EntraProviderConfig<IEntraSettings>)EntraProviderConfig<IEntraSettings>.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID));
            if (configuration != null)
            {
                configuration.Delete();
            }
        }

        /// <summary>
        /// Verifies if claims provider can run in the specified <paramref name="context"/>, and if it has valid and up to date <see cref="Settings"/>.
        /// </summary>
        /// <param name="context">The URI of the current site, or null</param>
        /// <returns>true if claims provider can run, false if it cannot continue</returns>
        public bool ValidateSettings(Uri context)
        {
            if (!Utils.IsClaimsProviderUsedInCurrentContext(context, Name))
            {
                return false;
            }

            if (this.SPTrust == null)
            {
                return false;
            }

            bool success = true;
            this.Lock_LocalConfigurationRefresh.EnterWriteLock();
            try
            {
                IEntraSettings settings = this.GetSettings();
                if (settings == null)
                {
                    return false;
                }

                if (settings.Version == this.SettingsVersion)
                {
                    Logger.Log($"[{this.Name}] Local copy of settings is up to date with version {this.SettingsVersion}.",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Core);
                    return true;
                }

                IEntraCPSettings claimsProviderSettings = new EntraCPSettings
                {
                    AlwaysResolveUserInput = settings.AlwaysResolveUserInput,
                    EntraIDTenantList = settings.EntraIDTenantList,
                    ClaimTypes = settings.ClaimTypes,
                    CustomData = settings.CustomData,
                    EnableAugmentation = settings.EnableAugmentation,
                    EntityDisplayTextPrefix = settings.EntityDisplayTextPrefix,
                    FilterExactMatchOnly = settings.FilterExactMatchOnly,
                    FilterSecurityEnabledGroupsOnly = settings.FilterSecurityEnabledGroupsOnly,
                    ProxyAddress = settings.ProxyAddress,
                    Timeout = settings.Timeout,
                    Version = settings.Version,
                };
                this.Settings = (IEntraCPSettings)claimsProviderSettings;

                Logger.Log($"[{this.Name}] Settings have new version {this.Settings.Version}, refreshing local copy",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
                success = this.InitializeInternalRuntimeSettings();
                if (success)
                {
#if !DEBUGx
                    this.SettingsVersion = this.Settings.Version;
#endif
                }
            }
            catch (Exception ex)
            {
                success = false;
                Logger.LogException(Name, "while refreshing configuration", TraceCategory.Core, ex);
            }
            finally
            {
                this.Lock_LocalConfigurationRefresh.ExitWriteLock();
            }
            return success;
        }

        /// <summary>
        /// Override this methor to return the settings to use
        /// </summary>
        /// <returns></returns>
        protected virtual IEntraSettings GetSettings()
        {
            IEntraSettings settings = null;
            EntraProviderConfig<IEntraSettings> PersistedConfiguration = (EntraProviderConfig<IEntraSettings>)EntraProviderConfig<IEntraSettings>.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID));
            if (PersistedConfiguration != null)
            {
                settings = PersistedConfiguration.Settings;
            }
            return settings;
        }

        /// <summary>
        /// Sets the internal runtime settings properties
        /// </summary>
        /// <returns>True if successful, false if not</returns>
        private bool InitializeInternalRuntimeSettings()
        {
            EntraCPSettings settings = (EntraCPSettings)this.Settings;
            if (settings.ClaimTypes?.Count <= 0)
            {
                Logger.Log($"[{this.Name}] Cannot continue because configuration has 0 claim configured.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                return false;
            }

            bool identityClaimTypeFound = false;
            bool groupClaimTypeFound = false;
            List<ClaimTypeConfig> claimTypesSetInTrust = new List<ClaimTypeConfig>();
            // Parse the ClaimTypeInformation collection set in the SPTrustedLoginProvider
            foreach (SPTrustedClaimTypeInformation claimTypeInformation in this.SPTrust.ClaimTypeInformation)
            {
                // Search if current claim type in trust exists in ClaimTypeConfigCollection
                ClaimTypeConfig claimTypeConfig = settings.ClaimTypes.FirstOrDefault(x =>
                    String.Equals(x.ClaimType, claimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                    !x.UseMainClaimTypeOfDirectoryObject &&
                    x.EntityProperty != DirectoryObjectProperty.NotSet);

                if (claimTypeConfig == null)
                {
                    continue;
                }
                ClaimTypeConfig localClaimTypeConfig = claimTypeConfig.CopyConfiguration();
                localClaimTypeConfig.ClaimTypeDisplayName = claimTypeInformation.DisplayName;
                claimTypesSetInTrust.Add(localClaimTypeConfig);
                if (String.Equals(this.SPTrust.IdentityClaimTypeInformation.MappedClaimType, localClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    // Identity claim type found, set IdentityClaimTypeConfig property
                    identityClaimTypeFound = true;
                    settings.IdentityClaimTypeConfig = IdentityClaimTypeConfig.ConvertClaimTypeConfig(localClaimTypeConfig);
                }
                else if (!groupClaimTypeFound && localClaimTypeConfig.EntityType == DirectoryObjectType.Group)
                {
                    groupClaimTypeFound = true;
                    settings.MainGroupClaimTypeConfig = localClaimTypeConfig;
                }
            }

            if (!identityClaimTypeFound)
            {
                Logger.Log($"[{this.Name}] Cannot continue because identity claim type '{this.SPTrust.IdentityClaimTypeInformation.MappedClaimType}' set in the SPTrustedIdentityTokenIssuer '{SPTrust.Name}' is missing in the ClaimTypeConfig list.", TraceSeverity.Unexpected, EventSeverity.ErrorCritical, TraceCategory.Core);
                return false;
            }

            // Check if there are additional properties to use in queries (UseMainClaimTypeOfDirectoryObject set to true)
            List<ClaimTypeConfig> additionalClaimTypeConfigList = new List<ClaimTypeConfig>();
            foreach (ClaimTypeConfig claimTypeConfig in settings.ClaimTypes.Where(x => x.UseMainClaimTypeOfDirectoryObject))
            {
                ClaimTypeConfig localClaimTypeConfig = claimTypeConfig.CopyConfiguration();
                if (localClaimTypeConfig.EntityType == DirectoryObjectType.User)
                {
                    localClaimTypeConfig.ClaimType = settings.IdentityClaimTypeConfig.ClaimType;
                    localClaimTypeConfig.EntityPropertyToUseAsDisplayText = settings.IdentityClaimTypeConfig.EntityPropertyToUseAsDisplayText;
                }
                else
                {
                    // If not a user, it must be a group
                    if (settings.MainGroupClaimTypeConfig == null)
                    {
                        continue;
                    }
                    localClaimTypeConfig.ClaimType = settings.MainGroupClaimTypeConfig.ClaimType;
                    localClaimTypeConfig.EntityPropertyToUseAsDisplayText = settings.MainGroupClaimTypeConfig.EntityPropertyToUseAsDisplayText;
                    localClaimTypeConfig.ClaimTypeDisplayName = settings.MainGroupClaimTypeConfig.ClaimTypeDisplayName;
                }
                additionalClaimTypeConfigList.Add(localClaimTypeConfig);
            }

            settings.RuntimeClaimTypesList = new List<ClaimTypeConfig>(claimTypesSetInTrust.Count + additionalClaimTypeConfigList.Count);
            settings.RuntimeClaimTypesList.AddRange(claimTypesSetInTrust);
            settings.RuntimeClaimTypesList.AddRange(additionalClaimTypeConfigList);

            // Get all PickerEntity metadata with a DirectoryObjectProperty set
            settings.RuntimeMetadataConfig = settings.ClaimTypes.Where(x =>
                !String.IsNullOrEmpty(x.EntityDataKey) &&
                x.EntityProperty != DirectoryObjectProperty.NotSet);

            if (settings.EntraIDTenantList == null || settings.EntraIDTenantList.Count < 1)
            {
                return false;
            }
            // Initialize Graph client on each tenant
            foreach (var tenant in settings.EntraIDTenantList)
            {
                tenant.InitializeAuthentication(settings.Timeout, settings.ProxyAddress);
            }
            this.Settings = settings;
            return true;
        }

        /// <summary>
        /// Search entities, or validate 1 entity, depending on <paramref name="currentContext"/>
        /// </summary>
        /// <param name="currentContext">Information about current context and operation</param>
        /// <returns>Entities generated by EntraCP</returns>
        protected List<PickerEntity> SearchOrValidate(OperationContext currentContext)
        {
            List<DirectoryObject> azureADEntityList = null;
            List<PickerEntity> pickerEntityList = new List<PickerEntity>();
            try
            {
                if (this.Settings.AlwaysResolveUserInput)
                {
                    // Completely bypass query to Microsoft Entra ID
                    pickerEntityList = CreatePickerEntityForSpecificClaimTypes(
                        currentContext.Input,
                        currentContext.CurrentClaimTypeConfigList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject),
                        false);
                    Logger.Log($"[{Name}] Created {pickerEntityList.Count} entity(ies) without contacting Microsoft Entra ID tenant(s) because EntraCP property AlwaysResolveUserInput is set to true.",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
                    return pickerEntityList;
                }

                if (currentContext.OperationType == OperationType.Search)
                {
                    // Call async method in a task to avoid error "Asynchronous operations are not allowed in this context" error when permission is validated (POST from people picker)
                    // More info on the error: https://stackoverflow.com/questions/672237/running-an-asynchronous-operation-triggered-by-an-asp-net-web-page-request
                    Task azureADQueryTask = Task.Run(async () =>
                    {
                        azureADEntityList = await SearchOrValidateInAzureADAsync(currentContext).ConfigureAwait(false);
                    });
                    azureADQueryTask.Wait();
                    pickerEntityList = this.ProcessAzureADResults(currentContext, azureADEntityList);

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
                            return pickerEntityList;
                        }
                        PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                            inputWithoutPrefix,
                            ctConfigWithInputPrefixMatch,
                            true);
                        if (entity != null)
                        {
                            if (pickerEntityList == null) { pickerEntityList = new List<PickerEntity>(); }
                            pickerEntityList.Add(entity);
                            Logger.Log($"[{Name}] Created entity without contacting Microsoft Entra ID tenant(s) because input started with prefix '{ctConfigWithInputPrefixMatch.PrefixToBypassLookup}', which is configured for claim type '{ctConfigWithInputPrefixMatch.ClaimType}'. Claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                            //return entities;
                        }
                    }
                }
                else if (currentContext.OperationType == OperationType.Validation)
                {
                    // Call async method in a task to avoid error "Asynchronous operations are not allowed in this context" error when permission is validated (POST from people picker)
                    // More info on the error: https://stackoverflow.com/questions/672237/running-an-asynchronous-operation-triggered-by-an-asp-net-web-page-request
                    Task azureADQueryTask = Task.Run(async () =>
                    {
                        azureADEntityList = await SearchOrValidateInAzureADAsync(currentContext).ConfigureAwait(false);
                    });
                    azureADQueryTask.Wait();
                    if (azureADEntityList?.Count == 1)
                    {
                        // Got the expected count (1 DirectoryObject)
                        pickerEntityList = this.ProcessAzureADResults(currentContext, azureADEntityList);
                    }
                    //if (entities?.Count == 1) { return entities; }

                    if (!String.IsNullOrEmpty(currentContext.IncomingEntityClaimTypeConfig.PrefixToBypassLookup))
                    {
                        // At this stage, it is impossible to know if entity was originally created with the keyword that bypass query to Microsoft Entra ID
                        // But it should be always validated since property PrefixToBypassLookup is set for current ClaimTypeConfig, so create entity manually
                        PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                            currentContext.IncomingEntity.Value,
                            currentContext.IncomingEntityClaimTypeConfig,
                            currentContext.InputHasKeyword);
                        if (entity != null)
                        {
                            pickerEntityList = new List<PickerEntity>(1) { entity };
                            Logger.Log($"[{Name}] Validated entity without contacting Microsoft Entra ID tenant(s) because its claim type ('{currentContext.IncomingEntityClaimTypeConfig.ClaimType}') has property 'PrefixToBypassLookup' set in EntraCPConfig.ClaimTypes. Claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in SearchOrValidate", TraceCategory.Claims_Picking, ex);
            }
            pickerEntityList = this.InspectEntitiesFound(currentContext, pickerEntityList);
            return pickerEntityList;
        }

        /// <summary>
        /// Override this method to inspect the entities generated by EntraCP during a search or a validation operation, and add or remove entities
        /// </summary>
        /// <param name="currentContext">Information about current context and operation</param>
        /// <param name="entities">Entities generated by EntraCP</param>
        /// <returns>Final list of entities that EntraCP will return to SharePoint</returns>
        protected virtual List<PickerEntity> InspectEntitiesFound(OperationContext currentContext, List<PickerEntity> entities)
        {
            return entities;
        }

        /// <summary>
        /// Search entities, or validate 1 entity, depending on <paramref name="currentContext"/>
        /// </summary>
        /// <param name="currentContext">Information about current context and operation</param>
        /// <returns></returns>
        protected async Task<List<DirectoryObject>> SearchOrValidateInAzureADAsync(OperationContext currentContext)
        {
            using (new SPMonitoredScope($"[{Name}] Total time spent to query Microsoft Entra ID tenant(s)", 1000))
            {
                List<DirectoryObject> results = await this.EntityProvider.SearchOrValidateEntitiesAsync(currentContext).ConfigureAwait(false);
                return results;
            }
        }

        private List<PickerEntity> ProcessAzureADResults(OperationContext currentContext, List<DirectoryObject> usersAndGroups)
        {
            if (usersAndGroups == null || !usersAndGroups.Any())
            {
                return null;
            };

            List<ClaimTypeConfig> ctConfigs = currentContext.CurrentClaimTypeConfigList;
            //Really?
            //if (currentContext.ExactSearch)
            //{
            //    ctConfigs = currentContext.CurrentClaimTypeConfigList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject);
            //}

            List<ClaimsProviderEntityResult> processedResults = new List<ClaimsProviderEntityResult>();
            foreach (DirectoryObject userOrGroup in usersAndGroups)
            {
                DirectoryObject currentObject = null;
                DirectoryObjectType objectType;
                if (userOrGroup is User)
                {
                    currentObject = userOrGroup;
                    objectType = DirectoryObjectType.User;
                }
                else
                {
                    currentObject = userOrGroup;
                    objectType = DirectoryObjectType.Group;

                    if (this.Settings.FilterSecurityEnabledGroupsOnly)
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
                    string directoryObjectPropertyValue = GetPropertyValue(currentObject, ctConfig.EntityProperty.ToString());

                    if (ctConfig is IdentityClaimTypeConfig)
                    {
                        if (String.Equals(((User)currentObject).UserType, ClaimsProviderConstants.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
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
                            claimTypeConfigToCompare = this.Settings.IdentityClaimTypeConfig;
                            if (String.Equals(((User)currentObject).UserType, ClaimsProviderConstants.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
                            {
                                // For Guest users, use the value set in property DirectoryObjectPropertyForGuestUsers
                                entityClaimValue = GetPropertyValue(currentObject, this.Settings.IdentityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers.ToString());
                            }
                            else
                            {
                                // Get the value of the DirectoryObjectProperty linked to current directory object
                                entityClaimValue = GetPropertyValue(currentObject, claimTypeConfigToCompare.EntityProperty.ToString());
                            }
                        }
                        else
                        {
                            claimTypeConfigToCompare = this.Settings.MainGroupClaimTypeConfig;
                            // Get the value of the DirectoryObjectProperty linked to current directory object
                            entityClaimValue = GetPropertyValue(currentObject, claimTypeConfigToCompare.EntityProperty.ToString());
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
                    processedResults.Add(new ClaimsProviderEntityResult(currentObject, ctConfig, entityClaimValue, directoryObjectPropertyValue));

                }
            }

            List<PickerEntity> entities = new List<PickerEntity>();
            Logger.Log($"[{Name}] {processedResults.Count} entity(ies) to create after filtering", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Lookup);
            foreach (ClaimsProviderEntityResult result in processedResults)
            {
                entities.Add(CreatePickerEntityHelper(result));
            }
            return entities;
        }

        private PickerEntity CreatePickerEntityHelper(ClaimsProviderEntityResult result)
        {
            PickerEntity entity = CreatePickerEntity();
            SPClaim claim;
            string permissionValue = result.PermissionValue;
            string permissionClaimType = result.ClaimTypeConfig.ClaimType;
            bool isMappedClaimTypeConfig = false;

            if (String.Equals(result.ClaimTypeConfig.ClaimType, this.Settings.IdentityClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase)
                || result.ClaimTypeConfig.UseMainClaimTypeOfDirectoryObject)
            {
                isMappedClaimTypeConfig = true;
            }

            entity.EntityType = result.ClaimTypeConfig.SharePointEntityType;
            if (result.ClaimTypeConfig.UseMainClaimTypeOfDirectoryObject)
            {
                string claimValueType;
                if (result.ClaimTypeConfig.EntityType == DirectoryObjectType.User)
                {
                    permissionClaimType = this.Settings.IdentityClaimTypeConfig.ClaimType;
                    claimValueType = this.Settings.IdentityClaimTypeConfig.ClaimValueType;
                    if (String.IsNullOrEmpty(entity.EntityType))
                    {
                        entity.EntityType = SPClaimEntityTypes.User;
                    }
                }
                else
                {
                    permissionClaimType = this.Settings.MainGroupClaimTypeConfig.ClaimType;
                    claimValueType = this.Settings.MainGroupClaimTypeConfig.ClaimValueType;
                    if (String.IsNullOrEmpty(entity.EntityType))
                    {
                        entity.EntityType = ClaimsProviderConstants.GroupClaimEntityType;
                    }
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
                if (String.IsNullOrEmpty(entity.EntityType))
                {
                    entity.EntityType = result.ClaimTypeConfig.EntityType == DirectoryObjectType.User ? SPClaimEntityTypes.User : ClaimsProviderConstants.GroupClaimEntityType;
                }
            }

            entity.Claim = claim;
            entity.IsResolved = true;
            //entity.EntityGroupName = "";
            entity.Description = String.Format(
                PickerEntityOnMouseOver,
                result.ClaimTypeConfig.EntityProperty.ToString(),
                result.DirectoryObjectPropertyValue);

            int nbMetadata = 0;
            // If current result is a SharePoint group but was found on an AAD User object, then 1 to many User objects could match so no metadata from the current match should be set
            if (!String.Equals(result.ClaimTypeConfig.SharePointEntityType, ClaimsProviderConstants.GroupClaimEntityType, StringComparison.InvariantCultureIgnoreCase) ||
                result.ClaimTypeConfig.EntityType != DirectoryObjectType.User)
            {
                // Populate metadata of new PickerEntity
                foreach (ClaimTypeConfig ctConfig in this.Settings.RuntimeMetadataConfig.Where(x => x.EntityType == result.ClaimTypeConfig.EntityType))
                {
                    // if there is actally a value in the GraphObject, then it can be set
                    string entityAttribValue = GetPropertyValue(result.DirectoryEntity, ctConfig.EntityProperty.ToString());
                    if (!String.IsNullOrEmpty(entityAttribValue))
                    {
                        entity.EntityData[ctConfig.EntityDataKey] = entityAttribValue;
                        nbMetadata++;
                        Logger.Log($"[{Name}] Set metadata '{ctConfig.EntityDataKey}' of new entity to '{entityAttribValue}'", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                    }
                }
            }
            entity.DisplayText = FormatPermissionDisplayText(entity, isMappedClaimTypeConfig, result);
            Logger.Log($"[{Name}] Created entity: display text: '{entity.DisplayText}', value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}', and filled with {nbMetadata.ToString()} metadata.", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
            return entity;
        }

        private PickerEntity CreatePickerEntityForSpecificClaimType(string input, ClaimTypeConfig ctConfig, bool inputHasKeyword)
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

        private List<PickerEntity> CreatePickerEntityForSpecificClaimTypes(string input, List<ClaimTypeConfig> ctConfigs, bool inputHasKeyword)
        {
            List<PickerEntity> entities = new List<PickerEntity>();
            foreach (var ctConfig in ctConfigs)
            {
                SPClaim claim = CreateClaim(ctConfig.ClaimType, input, ctConfig.ClaimValueType);
                PickerEntity entity = CreatePickerEntity();
                entity.Claim = claim;
                entity.IsResolved = true;
                entity.EntityType = ctConfig.SharePointEntityType;
                if (String.IsNullOrEmpty(entity.EntityType))
                {
                    entity.EntityType = ctConfig.EntityType == DirectoryObjectType.User ? SPClaimEntityTypes.User : ClaimsProviderConstants.GroupClaimEntityType;
                }
                //entity.EntityGroupName = "";
                entity.Description = String.Format(PickerEntityOnMouseOver, ctConfig.EntityProperty.ToString(), input);

                if (!String.IsNullOrEmpty(ctConfig.EntityDataKey))
                {
                    entity.EntityData[ctConfig.EntityDataKey] = entity.Claim.Value;
                    Logger.Log($"[{Name}] Added metadata '{ctConfig.EntityDataKey}' with value '{entity.EntityData[ctConfig.EntityDataKey]}' to new entity", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                }

                ClaimsProviderEntityResult result = new ClaimsProviderEntityResult(null, ctConfig, input, input);
                bool isIdentityClaimType = String.Equals(claim.ClaimType, this.Settings.IdentityClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase);
                entity.DisplayText = FormatPermissionDisplayText(entity, isIdentityClaimType, result);

                entities.Add(entity);
                Logger.Log($"[{Name}] Created entity: display text: '{entity.DisplayText}', value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'.", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
            }
            return entities.Count > 0 ? entities : null;
        }

        /// <summary>
        /// Override this method to customize value of permission being created
        /// </summary>
        /// <param name="claimType"></param>
        /// <param name="claimValue"></param>
        /// <param name="isIdentityClaimType"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        protected virtual string FormatPermissionValue(string claimType, string claimValue, bool isIdentityClaimType, ClaimsProviderEntityResult result)
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
        protected virtual string FormatPermissionDisplayText(PickerEntity entity, bool isMappedClaimTypeConfig, ClaimsProviderEntityResult result)
        {
            string entityDisplayText = this.Settings.EntityDisplayTextPrefix;
            if (result.ClaimTypeConfig.EntityPropertyToUseAsDisplayText != DirectoryObjectProperty.NotSet)
            {
                if (!isMappedClaimTypeConfig || result.ClaimTypeConfig.EntityType == DirectoryObjectType.Group)
                {
                    entityDisplayText += "(" + result.ClaimTypeConfig.ClaimTypeDisplayName + ") ";
                }

                string graphPropertyToDisplayValue = GetPropertyValue(result.DirectoryEntity, result.ClaimTypeConfig.EntityPropertyToUseAsDisplayText.ToString());
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
                    entityDisplayText += result.DirectoryObjectPropertyValue;
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

        /// <summary>
        /// Uses reflection to return the value of a public property for the given object
        /// </summary>
        /// <param name="directoryObject"></param>
        /// <param name="propertyName"></param>
        /// <returns>Null if property does not exist, String.Empty if property exists but it has no value, actual value otherwise</returns>
        public static string GetPropertyValue(object directoryObject, string propertyName)
        {
            if (directoryObject == null)
            {
                return null;
            }

            if (propertyName.StartsWith("extensionAttribute"))
            {
                try
                {
                    var returnString = string.Empty;
                    if (directoryObject is User)
                    {
                        var userobject = (User)directoryObject;
                        if (userobject.AdditionalData != null)
                        {
                            var obj = userobject.AdditionalData.FirstOrDefault(s => s.Key.EndsWith(propertyName));
                            if (obj.Value != null)
                            {
                                returnString = obj.Value.ToString();
                            }
                        }
                    }
                    else if (directoryObject is Group)
                    {
                        var groupobject = (Group)directoryObject;
                        if (groupobject.AdditionalData != null)
                        {
                            var obj = groupobject.AdditionalData.FirstOrDefault(s => s.Key.EndsWith(propertyName));
                            if (obj.Value != null)
                            {
                                returnString = obj.Value.ToString();
                            }
                        }
                    }
                    // Never return null for an extensionAttribute since we know it exists for both User and Group
                    return returnString == null ? String.Empty : returnString;
                }
                catch
                {
                    return String.Empty;
                }
            }

            PropertyInfo pi = directoryObject.GetType().GetProperty(propertyName);
            if (pi == null)
            {
                return null; // Property does not exist, return null
            }
            object propertyValue = pi.GetValue(directoryObject, null);
            return propertyValue == null ? String.Empty : propertyValue.ToString();
        }

        protected override void FillSchema(SPProviderSchema schema)
        {
            schema.AddSchemaElement(new SPSchemaElement(PeopleEditorEntityDataKeys.DisplayName, "Display Name", SPSchemaElementType.Both));
        }

        protected override void FillClaimTypes(List<string> claimTypes)
        {
            if (claimTypes == null) { return; }
            bool configIsValid = ValidateSettings(null);
            if (configIsValid)
            {
                this.Lock_LocalConfigurationRefresh.EnterReadLock();
                try
                {

                    foreach (var claimTypeSettings in this.Settings.RuntimeClaimTypesList)
                    {
                        claimTypes.Add(claimTypeSettings.ClaimType);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogException(Name, "in FillClaimTypes", TraceCategory.Core, ex);
                }
                finally
                {
                    this.Lock_LocalConfigurationRefresh.ExitReadLock();
                }
            }
        }

        protected override void FillClaimValueTypes(List<string> claimValueTypes)
        {
            claimValueTypes.Add(WIF4_5.ClaimValueTypes.String);
        }

        protected override void FillEntityTypes(List<string> entityTypes)
        {
            entityTypes.Add(SPClaimEntityTypes.User);
            entityTypes.Add(ClaimsProviderConstants.GroupClaimEntityType);
        }

        protected override void FillClaimsForEntity(Uri context, SPClaim entity, List<SPClaim> claims)
        {
            AugmentEntity(context, entity, null, claims);
        }
        protected override void FillClaimsForEntity(Uri context, SPClaim entity, SPClaimProviderContext claimProviderContext, List<SPClaim> claims)
        {
            AugmentEntity(context, entity, claimProviderContext, claims);
        }

        /// <summary>
        /// Gets the group membership of the <paramref name="entity"/> and add it to the list of <paramref name="claims"/>
        /// </summary>
        /// <param name="context"></param>
        /// <param name="entity">entity to augment</param>
        /// <param name="claimProviderContext">Can be null</param>
        /// <param name="claims"></param>
        protected void AugmentEntity(Uri context, SPClaim entity, SPClaimProviderContext claimProviderContext, List<SPClaim> claims)
        {
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
                Logger.Log($"[{Name}] Not trying to augment '{decodedEntity.Value}' because his OriginalIssuer is '{decodedEntity.OriginalIssuer}'.",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Augmentation);
                return;
            }

            if (!ValidateSettings(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                // There can be multiple TrustedProvider on the farm, but EntraCP should only do augmentation if current entity is from TrustedProvider it is associated with
                if (!String.Equals(decodedEntity.OriginalIssuer, this.OriginalIssuerName, StringComparison.InvariantCultureIgnoreCase)) { return; }

                if (!this.Settings.EnableAugmentation) { return; }

                Logger.Log($"[{Name}] Starting augmentation for user '{decodedEntity.Value}'.", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                ClaimTypeConfig groupClaimTypeSettings = this.Settings.RuntimeClaimTypesList.FirstOrDefault(x => x.EntityType == DirectoryObjectType.Group);
                if (groupClaimTypeSettings == null)
                {
                    Logger.Log($"[{Name}] No claim type with EntityType 'Group' was found, please check claims mapping table.",
                        TraceSeverity.High, EventSeverity.Error, TraceCategory.Augmentation);
                    return;
                }

                OperationContext currentContext = new OperationContext(this.Settings, OperationType.Augmentation, null, decodedEntity, context, null, null, Int32.MaxValue);
                Stopwatch timer = new Stopwatch();
                timer.Start();
                Task<List<string>> groupsTask = this.EntityProvider.GetEntityGroupsAsync(currentContext, groupClaimTypeSettings.EntityProperty);
                groupsTask.Wait();
                List<string> groups = groupsTask.Result;
                timer.Stop();
                if (groups?.Count > 0)
                {
                    foreach (string group in groups)
                    {
                        claims.Add(CreateClaim(groupClaimTypeSettings.ClaimType, group, groupClaimTypeSettings.ClaimValueType));
                        Logger.Log($"[{Name}] Added group '{group}' to user '{currentContext.IncomingEntity.Value}'",
                            TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                    }
                    Logger.Log($"[{Name}] Augmented user '{currentContext.IncomingEntity.Value}' with {groups.Count} groups in {timer.ElapsedMilliseconds} ms",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Augmentation);
                }
                else
                {
                    Logger.Log($"[{Name}] Got no group in {timer.ElapsedMilliseconds} ms for user '{currentContext.IncomingEntity.Value}'",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Augmentation);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in AugmentEntity", TraceCategory.Augmentation, ex);
            }
            finally
            {
                this.Lock_LocalConfigurationRefresh.ExitReadLock();
            }
        }

        protected virtual new SPClaim CreateClaim(string type, string value, string valueType)
        {
            // SPClaimProvider.CreateClaim sets property OriginalIssuer to SPOriginalIssuerType.ClaimProvider, which is not correct
            //return CreateClaim(type, value, valueType);
            return new SPClaim(type, value, valueType, this.OriginalIssuerName);
        }

        protected override void FillHierarchy(Uri context, string[] entityTypes, string hierarchyNodeID, int numberOfLevels, SPProviderHierarchyTree hierarchy)
        {
            List<DirectoryObjectType> aadEntityTypes = new List<DirectoryObjectType>();
            if (entityTypes.Contains(SPClaimEntityTypes.User)) { aadEntityTypes.Add(DirectoryObjectType.User); }
            if (entityTypes.Contains(ClaimsProviderConstants.GroupClaimEntityType)) { aadEntityTypes.Add(DirectoryObjectType.Group); }

            if (!ValidateSettings(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                if (hierarchyNodeID == null)
                {
                    // Root level
                    foreach (var azureObject in this.Settings.RuntimeClaimTypesList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject && aadEntityTypes.Contains(x.EntityType)))
                    {
                        hierarchy.AddChild(
                            new Microsoft.SharePoint.WebControls.SPProviderHierarchyNode(
                                Name,
                                azureObject.ClaimTypeDisplayName,
                                azureObject.ClaimType,
                                true));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in FillHierarchy", TraceCategory.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_LocalConfigurationRefresh.ExitReadLock();
            }
        }

        protected override void FillResolve(Uri context, string[] entityTypes, string resolveInput, List<PickerEntity> resolved)
        {
            if (!ValidateSettings(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                OperationContext currentContext = new OperationContext(this.Settings, OperationType.Search, resolveInput, null, context, entityTypes, null, 30);
                List<PickerEntity> entities = SearchOrValidate(currentContext);
                if (entities == null || entities.Count == 0) { return; }
                foreach (PickerEntity entity in entities)
                {
                    resolved.Add(entity);
                    Logger.Log($"[{Name}] Added entity: display text: '{entity.DisplayText}', claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                        TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
                Logger.Log($"[{Name}] Returned {entities.Count} entities with input '{currentContext.Input}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in FillResolve(string)", TraceCategory.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_LocalConfigurationRefresh.ExitReadLock();
            }
        }

        protected override void FillResolve(Uri context, string[] entityTypes, SPClaim resolveInput, List<PickerEntity> resolved)
        {
            if (!ValidateSettings(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                // Ensure incoming claim should be validated by EntraCP
                // Must be made after call to Initialize because SPTrustedLoginProvider name must be known
                if (!String.Equals(resolveInput.OriginalIssuer, this.OriginalIssuerName, StringComparison.InvariantCultureIgnoreCase)) { return; }

                OperationContext currentContext = new OperationContext(this.Settings, OperationType.Validation, resolveInput.Value, resolveInput, context, entityTypes, null, 1);
                List<PickerEntity> entities = this.SearchOrValidate(currentContext);
                if (entities?.Count == 1)
                {
                    resolved.Add(entities[0]);
                    Logger.Log($"[{Name}] Validated entity: display text: '{entities[0].DisplayText}', claim value: '{entities[0].Claim.Value}', claim type: '{entities[0].Claim.ClaimType}'",
                        TraceSeverity.High, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
                else
                {
                    int entityCount = entities == null ? 0 : entities.Count;
                    Logger.Log($"[{Name}] Validation failed: found {entityCount.ToString()} entities instead of 1 for incoming claim with value '{currentContext.IncomingEntity.Value}' and type '{currentContext.IncomingEntity.ClaimType}'", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Claims_Picking);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in FillResolve(SPClaim)", TraceCategory.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_LocalConfigurationRefresh.ExitReadLock();
            }
        }

        protected override void FillSearch(Uri context, string[] entityTypes, string searchPattern, string hierarchyNodeID, int maxCount, SPProviderHierarchyTree searchTree)
        {
            if (!ValidateSettings(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                OperationContext currentContext = new OperationContext(this.Settings, OperationType.Search, searchPattern, null, context, entityTypes, hierarchyNodeID, maxCount);
                List<PickerEntity> entities = this.SearchOrValidate(currentContext);
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
                        ClaimTypeConfig ctConfig = this.Settings.RuntimeClaimTypesList.FirstOrDefault(x =>
                            !x.UseMainClaimTypeOfDirectoryObject &&
                            String.Equals(x.ClaimType, entity.Claim.ClaimType, StringComparison.InvariantCultureIgnoreCase));

                        string nodeName = ctConfig != null ? ctConfig.ClaimTypeDisplayName : entity.Claim.ClaimType;
                        matchNode = new SPProviderHierarchyNode(Name, nodeName, entity.Claim.ClaimType, true);
                        searchTree.AddChild(matchNode);
                    }
                    matchNode.AddEntity(entity);
                    Logger.Log($"[{Name}] Added entity: display text: '{entity.DisplayText}', claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                        TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
                Logger.Log($"[{Name}] Returned {entities.Count} entities from input '{currentContext.Input}'",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                this.Lock_LocalConfigurationRefresh.ExitReadLock();
            }
        }

        /// <summary>
        /// Return the identity claim type
        /// </summary>
        /// <returns></returns>
        public override string GetClaimTypeForUserKey()
        {
            try
            {
                return this.SPTrust != null ? this.SPTrust.IdentityClaimTypeInformation.MappedClaimType : String.Empty;
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in GetClaimTypeForUserKey", TraceCategory.Rehydration, ex);
            }
            return String.Empty;
        }

        /// <summary>
        /// Return the user key (SPClaim with identity claim type) from the incoming entity
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        protected override SPClaim GetUserKeyForEntity(SPClaim entity)
        {
            try
            {
                if (this.SPTrust == null)
                {
                    return entity;
                }

                // There are 2 scenarios:
                // 1: OriginalIssuer is "SecurityTokenService": Value looks like "05.t|contoso.local|yvand@contoso.local", claim type is "http://schemas.microsoft.com/sharepoint/2009/08/claims/userid" and it must be decoded properly
                // 2: OriginalIssuer is "TrustedProvider:contoso.local": The incoming entity is fine and returned as is
                if (String.Equals(entity.OriginalIssuer, this.OriginalIssuerName, StringComparison.InvariantCultureIgnoreCase))
                {
                    return entity;
                }

                // SPClaimProviderManager.IsUserIdentifierClaim tests if:
                // ClaimType == SPClaimTypes.UserIdentifier ("http://schemas.microsoft.com/sharepoint/2009/08/claims/userid")
                // OriginalIssuer type == SPOriginalIssuerType.SecurityTokenService
                if (!SPClaimProviderManager.IsUserIdentifierClaim(entity))
                {
                    // return entity if not true, otherwise SPClaimProviderManager.DecodeUserIdentifierClaim(entity) throws an ArgumentException
                    return entity;
                }

                // Since SPClaimProviderManager.IsUserIdentifierClaim() returned true, SPClaimProviderManager.DecodeUserIdentifierClaim() will work
                SPClaim curUser = SPClaimProviderManager.DecodeUserIdentifierClaim(entity);
                Logger.Log($"[{Name}] Returning user key for '{entity.Value}'",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Rehydration);
                return CreateClaim(this.SPTrust.IdentityClaimTypeInformation.MappedClaimType, curUser.Value, curUser.ValueType);
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in GetUserKeyForEntity", TraceCategory.Rehydration, ex);
            }
            return null;
        }
    }

    /// <summary>
    /// User / group found in Microsoft Entra ID, with additional information
    /// </summary>
    public class ClaimsProviderEntityResult
    {
        /// <summary>
        /// Gets the entity returned by Microsoft Entra ID
        /// </summary>
        public DirectoryObject DirectoryEntity { get; private set; }

        /// <summary>
        /// Gets the relevant ClaimTypeConfig object to use for the property PickerEntity.Claim
        /// </summary>
        public ClaimTypeConfig ClaimTypeConfig { get; private set; }

        /// <summary>
        /// Gets the DirectoryObject's attribute value to use as the actual permission value
        /// </summary>
        public string PermissionValue { get; private set; }

        /// <summary>
        /// Gets the DirectoryObject's attribute value which matched the query
        /// </summary>
        public string DirectoryObjectPropertyValue { get; private set; }

        public ClaimsProviderEntityResult(DirectoryObject directoryEntity, ClaimTypeConfig claimTypeConfig, string permissionValue, string directoryObjectPropertyValue)
        {
            DirectoryEntity = directoryEntity;
            ClaimTypeConfig = claimTypeConfig;
            PermissionValue = permissionValue;
            DirectoryObjectPropertyValue = directoryObjectPropertyValue;
        }
    }
}
