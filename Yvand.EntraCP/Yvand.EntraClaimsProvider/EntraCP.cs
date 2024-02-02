using Azure.Core.Diagnostics;
using Microsoft.Graph.Models;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Tracing;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Yvand.EntraClaimsProvider.Configuration;
using WIF4_5 = System.Security.Claims;

namespace Yvand.EntraClaimsProvider
{
    public interface IEntraCPSettings : IEntraIDProviderSettings
    {
        List<ClaimTypeConfig> RuntimeClaimTypesList { get; }
        IEnumerable<ClaimTypeConfig> RuntimeMetadataConfig { get; }
        IdentityClaimTypeConfig IdentityClaimTypeConfig { get; }
        ClaimTypeConfig MainGroupClaimTypeConfig { get; }
    }

    public class EntraCPSettings : EntraIDProviderSettings, IEntraCPSettings
    {
        public static new EntraCPSettings GetDefaultSettings(string claimsProviderName)
        {
            EntraIDProviderSettings entraIDProviderSettings = EntraIDProviderSettings.GetDefaultSettings(claimsProviderName);
            return GenerateFromEntraIDProviderSettings(entraIDProviderSettings);
        }

        public static EntraCPSettings GenerateFromEntraIDProviderSettings(IEntraIDProviderSettings settings)
        {
            EntraCPSettings copy = new EntraCPSettings();
            Utils.CopyPublicProperties(typeof(EntraIDProviderSettings), settings, copy);
            return copy;
        }

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

        /// <summary>
        /// Gets the settings that contain the configuration for EntraCP
        /// </summary>
        public IEntraCPSettings Settings { get; protected set; }

        /// <summary>
        /// Gets custom settings that will be used instead of the settings from the persisted object
        /// </summary>
        private IEntraCPSettings CustomSettings { get; }

        /// <summary>
        /// Gets the version of the settings, used to refresh the settings if the persisted object is updated
        /// </summary>
        public long SettingsVersion { get; private set; } = -1;
        AzureEventSourceListener GraphEventsListener;

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

        /// <summary>
        /// Gets the issuer formatted to be like the property SPClaim.OriginalIssuer: "TrustedProvider:TrustedProviderName"
        /// </summary>
        public string OriginalIssuerName => this.SPTrust != null ? SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, this.SPTrust.Name) : String.Empty;

        public EntraCP(string displayName) : base(displayName)
        {
            this.GraphEventsListener = new AzureEventSourceListener((args, message) =>
            {
                if (args.EventSource.Name == "Azure-Identity")
                {
                    Logger.Log($"[{EntraCP.ClaimsProviderName}] {args.EventName} {message}", Utils.EventLogToTraceSeverity(args.Level), EventSeverity.Error, TraceCategory.AzureIdentity);
                }
            }, EventLevel.Informational);
        }

        public EntraCP(string displayName, IEntraCPSettings customSettings) : base(displayName)
        {
            this.CustomSettings = customSettings;
        }

        #region ManageConfiguration
        public static EntraIDProviderConfiguration GetConfiguration(bool initializeLocalConfiguration = false)
        {
            EntraIDProviderConfiguration configuration = EntraIDProviderConfiguration.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID), initializeLocalConfiguration);
            return configuration;
        }

        /// <summary>
        /// Creates a configuration for EntraCP. This will delete any existing configuration which may already exist
        /// </summary>
        /// <returns></returns>
        public static EntraIDProviderConfiguration CreateConfiguration()
        {
            EntraIDProviderConfiguration configuration = EntraIDProviderConfiguration.CreateGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID), ClaimsProviderConstants.CONFIGURATION_NAME, EntraCP.ClaimsProviderName);
            return configuration;
        }

        /// <summary>
        /// Deletes the configuration for EntraCP
        /// </summary>
        public static void DeleteConfiguration()
        {
            EntraIDProviderConfiguration configuration = EntraIDProviderConfiguration.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID));
            if (configuration != null)
            {
                configuration.Delete();
            }
        }
        #endregion

        #region Initialization
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
                IEntraIDProviderSettings settings = this.GetSettings();
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

                this.Settings = EntraCPSettings.GenerateFromEntraIDProviderSettings(settings);
                Logger.Log($"[{this.Name}] Settings have new version {this.Settings.Version}, refreshing local copy",
                    TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);
                success = this.InitializeInternalRuntimeSettings();
                if (success)
                {
#if !DEBUGx
                    this.SettingsVersion = this.Settings.Version;
#endif
                    this.EntityProvider = new EntraIDEntityProvider(Name, this.Settings);
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
        /// Returns the settings to use
        /// </summary>
        /// <returns></returns>
        public virtual IEntraIDProviderSettings GetSettings()
        {
            if (this.CustomSettings != null)
            {
                return this.CustomSettings;
            }

            IEntraIDProviderSettings persistedSettings = null;
            EntraIDProviderConfiguration PersistedConfiguration = EntraIDProviderConfiguration.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID));
            if (PersistedConfiguration != null)
            {
                persistedSettings = PersistedConfiguration.Settings;
            }
            return persistedSettings;
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
                !String.IsNullOrWhiteSpace(x.EntityDataKey) &&
                x.EntityProperty != DirectoryObjectProperty.NotSet);

            if (settings.EntraIDTenants == null || settings.EntraIDTenants.Count < 1)
            {
                return false;
            }
            // Initialize Graph client on each tenant
            foreach (var tenant in settings.EntraIDTenants)
            {
                tenant.InitializeAuthentication(settings.Timeout, settings.ProxyAddress);
            }
            this.Settings = settings;
            return true;
        }
        #endregion

        #region Augmentation
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
                groupsTask.ConfigureAwait(false);
                groupsTask.Wait(this.Settings.Timeout);
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
        #endregion

        #region Search
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
                Logger.Log($"[{Name}] Returned {entities.Count} entities with value '{currentContext.Input}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
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
                Logger.Log($"[{Name}] Returned {entities.Count} entities from value '{currentContext.Input}'",
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
        #endregion

        #region Validation
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
        #endregion

        #region ProcessSearchOrValidation
        /// <summary>
        /// Search spEntities, or validate 1 entity, depending on <paramref name="currentContext"/>
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
                        currentContext.CurrentClaimTypeConfigList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject));
                    Logger.Log($"[{Name}] Created {pickerEntityList.Count} entity(ies) without contacting Microsoft Entra ID tenant(s) because EntraCP property AlwaysResolveUserInput is set to true.",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Claims_Picking);
                    return pickerEntityList;
                }

                // Create a delegate to query Entra ID, so it is called only if needed
                Func<Task> SearchOrValidateInEntraID = delegate ()
                {
                    return Task.Run(async () =>
                    {
                        using (new SPMonitoredScope($"[{Name}] Total time spent to query Microsoft Entra ID tenant(s)", 1000))
                        {
                            azureADEntityList = await this.EntityProvider.SearchOrValidateEntitiesAsync(currentContext).ConfigureAwait(false);
                        }
                    });
                };

                if (currentContext.OperationType == OperationType.Search)
                {
                    // Between 0 to many PickerEntity is expected by SharePoint

                    // Check if value starts with a prefix configured on a ClaimTypeConfig. If so an entity should be returned using ClaimTypeConfig found
                    // ClaimTypeConfigEnsureUniquePrefixToBypassLookup ensures that collection cannot contain duplicates
                    ClaimTypeConfig ctConfigWithInputPrefixMatch = currentContext.CurrentClaimTypeConfigList.FirstOrDefault(x =>
                        !String.IsNullOrEmpty(x.PrefixToBypassLookup) &&
                        currentContext.Input.StartsWith(x.PrefixToBypassLookup, StringComparison.InvariantCultureIgnoreCase));
                    if (ctConfigWithInputPrefixMatch != null)
                    {
                        string inputWithoutPrefix = currentContext.Input.Substring(ctConfigWithInputPrefixMatch.PrefixToBypassLookup.Length);
                        if (String.IsNullOrEmpty(inputWithoutPrefix))
                        {
                            // No value in the value after the prefix, return
                            return pickerEntityList;
                        }
                        pickerEntityList = CreatePickerEntityForSpecificClaimTypes(
                            inputWithoutPrefix,
                            new List<ClaimTypeConfig>() { ctConfigWithInputPrefixMatch });
                        if (pickerEntityList?.Count == 1)
                        {
                            PickerEntity entity = pickerEntityList.FirstOrDefault();
                            Logger.Log($"[{Name}] Created entity without contacting Microsoft Entra ID tenant(s) because value started with prefix '{ctConfigWithInputPrefixMatch.PrefixToBypassLookup}', which is configured for claim type '{ctConfigWithInputPrefixMatch.ClaimType}'. Claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                        }
                    }
                    else
                    {
                        // Call async method in a task to avoid error "Asynchronous operations are not allowed in this context" error when permission is validated (POST from people picker)
                        // More info on the error: https://stackoverflow.com/questions/672237/running-an-asynchronous-operation-triggered-by-an-asp-net-web-page-request
                        Task.Run(async () => await SearchOrValidateInEntraID()).Wait();
                        pickerEntityList = this.ProcessAzureADResults(currentContext, azureADEntityList);
                    }
                }
                else if (currentContext.OperationType == OperationType.Validation)
                {
                    // Exactly 1 PickerEntity is expected by SharePoint

                    // Check if config corresponding to current claim type has a config to bypass Entra ID
                    if (!String.IsNullOrWhiteSpace(currentContext.IncomingEntityClaimTypeConfig.PrefixToBypassLookup))
                    {
                        // At this stage, it is impossible to know if entity was originally created with the keyword that bypass query to Microsoft Entra ID
                        // But it should be always validated since property PrefixToBypassLookup is set for current ClaimTypeConfig, so create entity manually
                        pickerEntityList = CreatePickerEntityForSpecificClaimTypes(
                            currentContext.IncomingEntity.Value,
                            new List<ClaimTypeConfig>() { currentContext.IncomingEntityClaimTypeConfig });
                        if (pickerEntityList?.Count == 1)
                        {
                            PickerEntity entity = pickerEntityList.FirstOrDefault();
                            Logger.Log($"[{Name}] Validated entity without contacting Microsoft Entra ID tenant(s) because its claim type ('{currentContext.IncomingEntityClaimTypeConfig.ClaimType}') has property 'PrefixToBypassLookup' set in EntraCPConfig.ClaimTypes. Claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                        }
                    }
                    else
                    {
                        // Call async method in a task to avoid error "Asynchronous operations are not allowed in this context" error when permission is validated (POST from people picker)
                        // More info on the error: https://stackoverflow.com/questions/672237/running-an-asynchronous-operation-triggered-by-an-asp-net-web-page-request
                        Task.Run(async () => await SearchOrValidateInEntraID()).Wait();
                        if (azureADEntityList?.Count == 1)
                        {
                            // Got the expected count (1 DirectoryObject)
                            pickerEntityList = this.ProcessAzureADResults(currentContext, azureADEntityList);
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
        /// Override this method to inspect the spEntities generated by EntraCP during a search or a validation operation, and add or remove spEntities
        /// </summary>
        /// <param name="currentContext">Information about current context and operation</param>
        /// <param name="entities">Entities generated by EntraCP</param>
        /// <returns>Final list of spEntities that EntraCP will return to SharePoint</returns>
        protected virtual List<PickerEntity> InspectEntitiesFound(OperationContext currentContext, List<PickerEntity> entities)
        {
            return entities;
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

            List<PickerEntity> spEntities = new List<PickerEntity>();
            List<ClaimsProviderEntity> uniqueDirectoryResults = new List<ClaimsProviderEntity>();
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
                    string directoryObjectPropertyValue = Utils.GetDirectoryObjectPropertyValue(currentObject, ctConfig.EntityProperty.ToString());

                    if (ctConfig is IdentityClaimTypeConfig)
                    {
                        if (String.Equals(((User)currentObject).UserType, ClaimsProviderConstants.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
                        {
                            // For Guest users, use the value set in property DirectoryObjectPropertyForGuestUsers
                            directoryObjectPropertyValue = Utils.GetDirectoryObjectPropertyValue(currentObject, ((IdentityClaimTypeConfig)ctConfig).DirectoryObjectPropertyForGuestUsers.ToString());
                        }
                    }

                    // Check if property exists (not null) and has a value (not String.Empty)
                    if (String.IsNullOrEmpty(directoryObjectPropertyValue)) { continue; }

                    // Check if current value mathes value, otherwise go to next GraphProperty to check
                    if (currentContext.ExactSearch)
                    {
                        if (!String.Equals(directoryObjectPropertyValue, currentContext.Input, StringComparison.InvariantCultureIgnoreCase)) { continue; }
                    }
                    else
                    {
                        if (!directoryObjectPropertyValue.StartsWith(currentContext.Input, StringComparison.InvariantCultureIgnoreCase)) { continue; }
                    }

                    // Current DirectoryObjectProperty value matches user value. Add current result to search results if it is not already present
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
                                entityClaimValue = Utils.GetDirectoryObjectPropertyValue(currentObject, this.Settings.IdentityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers.ToString());
                            }
                            else
                            {
                                // Get the value of the DirectoryObjectProperty linked to current directory object
                                entityClaimValue = Utils.GetDirectoryObjectPropertyValue(currentObject, claimTypeConfigToCompare.EntityProperty.ToString());
                            }
                        }
                        else
                        {
                            claimTypeConfigToCompare = this.Settings.MainGroupClaimTypeConfig;
                            // Get the value of the DirectoryObjectProperty linked to current directory object
                            entityClaimValue = Utils.GetDirectoryObjectPropertyValue(currentObject, claimTypeConfigToCompare.EntityProperty.ToString());
                        }

                        if (String.IsNullOrEmpty(entityClaimValue)) { continue; }
                    }
                    else
                    {
                        claimTypeConfigToCompare = ctConfig;
                    }

                    // if claim type and claim value already exists, skip
                    bool resultAlreadyExists = uniqueDirectoryResults.Exists(x =>
                        String.Equals(x.ClaimTypeConfigMatch.ClaimType, claimTypeConfigToCompare.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        String.Equals(x.PermissionValue, entityClaimValue, StringComparison.InvariantCultureIgnoreCase));
                    if (resultAlreadyExists) { continue; }

                    // Passed the checks, add it to the uniqueDirectoryResults list
                    ClaimsProviderEntity claimsProviderEntity = new ClaimsProviderEntity(currentObject, ctConfig, entityClaimValue, directoryObjectPropertyValue);
                    spEntities.Add(CreatePickerEntityHelper(currentContext, claimsProviderEntity));
                    uniqueDirectoryResults.Add(claimsProviderEntity);

                }
            }
            Logger.Log($"[{Name}] Created {spEntities.Count} entity(ies) after filtering directory results", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Lookup);
            return spEntities;
        }
        #endregion

        #region Helpers
        protected PickerEntity CreatePickerEntityHelper(OperationContext currentContext, ClaimsProviderEntity result)
        {
            ClaimTypeConfig ctConfigToUseForClaimValue = result.ClaimTypeConfigMatch;
            if (result.ClaimTypeConfigMatch.UseMainClaimTypeOfDirectoryObject)
            {
                // Get the config to use to create the actual entity (claim type and its DirectoryObjectAttribute) from current result
                ctConfigToUseForClaimValue = result.ClaimTypeConfigMatch.EntityType == DirectoryObjectType.User ? this.Settings.IdentityClaimTypeConfig : this.Settings.MainGroupClaimTypeConfig;
            }

            string permissionValue = FormatPermissionValue(result.PermissionValue);
            SPClaim claim = CreateClaim(ctConfigToUseForClaimValue.ClaimType, permissionValue, ctConfigToUseForClaimValue.ClaimValueType);
            PickerEntity entity = CreatePickerEntity();
            entity.Claim = claim;
            entity.EntityType = ctConfigToUseForClaimValue.EntityType == DirectoryObjectType.User ? SPClaimEntityTypes.User : ClaimsProviderConstants.GroupClaimEntityType;
            entity.IsResolved = true;
            entity.EntityGroupName = this.Name;
            entity.Description = String.Format(
                PickerEntityOnMouseOver,
                result.ClaimTypeConfigMatch.EntityProperty.ToString(),
                result.DirectoryObjectPropertyValueMatch);
            entity.DisplayText = FormatPermissionDisplayText(entity, result.ClaimTypeConfigMatch.UseMainClaimTypeOfDirectoryObject, result);

            int nbMetadata = 0;
            // Populate the metadata for this PickerEntity
            // Populate metadata of new PickerEntity
            foreach (ClaimTypeConfig ctConfig in this.Settings.RuntimeMetadataConfig.Where(x => x.EntityType == result.ClaimTypeConfigMatch.EntityType))
            {
                // if there is actally a value in the GraphObject, then it can be set
                string entityAttribValue = Utils.GetDirectoryObjectPropertyValue(result.DirectoryEntity, ctConfig.EntityProperty.ToString());
                if (!String.IsNullOrEmpty(entityAttribValue))
                {
                    entity.EntityData[ctConfig.EntityDataKey] = entityAttribValue;
                    nbMetadata++;
                    Logger.Log($"[{Name}] Set metadata '{ctConfig.EntityDataKey}' of new entity to '{entityAttribValue}'", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                }
            }

            Logger.Log($"[{Name}] Created entity: display text: '{entity.DisplayText}', claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}', and filled with {nbMetadata} metadata.", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
            return entity;
        }

        private List<PickerEntity> CreatePickerEntityForSpecificClaimTypes(string value, List<ClaimTypeConfig> ctConfigs)
        {
            List<PickerEntity> entities = new List<PickerEntity>();
            foreach (var ctConfig in ctConfigs)
            {
                SPClaim claim = CreateClaim(ctConfig.ClaimType, value, ctConfig.ClaimValueType);
                PickerEntity entity = CreatePickerEntity();
                entity.Claim = claim;
                entity.IsResolved = true;
                entity.EntityType = ctConfig.SharePointEntityType;
                if (String.IsNullOrEmpty(entity.EntityType))
                {
                    entity.EntityType = ctConfig.EntityType == DirectoryObjectType.User ? SPClaimEntityTypes.User : ClaimsProviderConstants.GroupClaimEntityType;
                }
                //entity.EntityGroupName = "";
                entity.Description = String.Format(PickerEntityOnMouseOver, ctConfig.EntityProperty.ToString(), value);

                if (!String.IsNullOrEmpty(ctConfig.EntityDataKey))
                {
                    entity.EntityData[ctConfig.EntityDataKey] = entity.Claim.Value;
                    Logger.Log($"[{Name}] Added metadata '{ctConfig.EntityDataKey}' with value '{entity.EntityData[ctConfig.EntityDataKey]}' to new entity", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                }

                ClaimsProviderEntity result = new ClaimsProviderEntity(ctConfig, value);
                bool isIdentityClaimType = String.Equals(claim.ClaimType, this.Settings.IdentityClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase);
                entity.DisplayText = FormatPermissionDisplayText(entity, isIdentityClaimType, result);

                entities.Add(entity);
                Logger.Log($"[{Name}] Created entity: display text: '{entity.DisplayText}', value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'.", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
            }
            return entities.Count > 0 ? entities : null;
        }

        protected virtual string FormatPermissionValue(string claimValue)
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
        protected virtual string FormatPermissionDisplayText(PickerEntity entity, bool isMappedClaimTypeConfig, ClaimsProviderEntity result)
        {
            string entityDisplayText = this.Settings.EntityDisplayTextPrefix;
            if (result.ClaimTypeConfigMatch.EntityPropertyToUseAsDisplayText != DirectoryObjectProperty.NotSet)
            {
                if (!isMappedClaimTypeConfig || result.ClaimTypeConfigMatch.EntityType == DirectoryObjectType.Group)
                {
                    entityDisplayText += "(" + result.ClaimTypeConfigMatch.ClaimTypeDisplayName + ") ";
                }

                string graphPropertyToDisplayValue = Utils.GetDirectoryObjectPropertyValue(result.DirectoryEntity, result.ClaimTypeConfigMatch.EntityPropertyToUseAsDisplayText.ToString());
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
                    entityDisplayText += result.DirectoryObjectPropertyValueMatch;
                }
                else
                {
                    entityDisplayText += String.Format(
                        PickerEntityDisplayText,
                        result.ClaimTypeConfigMatch.ClaimTypeDisplayName,
                        result.PermissionValue);
                }
            }
            return entityDisplayText;
        }

        protected virtual new SPClaim CreateClaim(string type, string value, string valueType)
        {
            // SPClaimProvider.CreateClaim sets property OriginalIssuer to SPOriginalIssuerType.ClaimProvider, which is not correct
            //return CreateClaim(type, value, valueType);
            return new SPClaim(type, value, valueType, this.OriginalIssuerName);
        }
        #endregion

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
    public class ClaimsProviderEntity
    {
        /// <summary>
        /// Gets the entity returned by Microsoft Entra ID
        /// </summary>
        public DirectoryObject DirectoryEntity { get; private set; }

        /// <summary>
        /// Gets the relevant ClaimTypeConfig object to use for the property PickerEntity.Claim
        /// </summary>
        public ClaimTypeConfig ClaimTypeConfigMatch { get; private set; }

        /// <summary>
        /// Gets the DirectoryObject's attribute value to use as the actual permission value
        /// </summary>
        public string PermissionValue { get; private set; }

        /// <summary>
        /// Gets the DirectoryObject's attribute value which matched the query
        /// </summary>
        public string DirectoryObjectPropertyValueMatch { get; private set; }

        public ClaimsProviderEntity(DirectoryObject directoryEntity, ClaimTypeConfig claimTypeConfig, string permissionValue, string directoryObjectPropertyValue)
        {
            DirectoryEntity = directoryEntity;
            ClaimTypeConfigMatch = claimTypeConfig;
            PermissionValue = permissionValue;
            DirectoryObjectPropertyValueMatch = directoryObjectPropertyValue;
        }

        public ClaimsProviderEntity(ClaimTypeConfig claimTypeConfig, string permissionValue)
        {
            ClaimTypeConfigMatch = claimTypeConfig;
            PermissionValue = permissionValue;
        }
    }
}
