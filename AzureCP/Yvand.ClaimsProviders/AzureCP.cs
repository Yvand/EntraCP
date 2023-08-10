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
using Yvand.ClaimsProviders.AzureAD;
using Yvand.ClaimsProviders.Config;
using WIF4_5 = System.Security.Claims;

namespace Yvand.ClaimsProviders
{
    public class AzureCP : SPClaimProvider
    {
        public static string ClaimsProviderName => "AzureCPSE";
        public override string Name => ClaimsProviderName;
        public override bool SupportsEntityInformation => true;
        public override bool SupportsHierarchy => true;
        public override bool SupportsResolve => true;
        public override bool SupportsSearch => true;
        public override bool SupportsUserKey => true;
        public AzureADEntityProvider EntityProvider { get; private set; }
        private ReaderWriterLockSlim Lock_LocalConfigurationRefresh = new ReaderWriterLockSlim();
        protected virtual string PickerEntityDisplayText => "({0}) {1}";
        protected virtual string PickerEntityOnMouseOver => "{0}={1}";
        protected AADEntityProviderConfig<IAADSettings> PersistedConfiguration { get; private set; }
        public IAADSettings LocalConfiguration { get; private set; }

        public AzureCP(string displayName) : base(displayName)
        {
            this.EntityProvider = new AzureADEntityProvider(Name);
        }

        public static AADEntityProviderConfig<IAADSettings> GetConfiguration(bool initializeRuntimeSettings = false)
        {
            //AzureADEntityProviderConfiguration configuration = EntityProviderBase<AzureADEntityProviderConfiguration>.GetGlobalConfiguration(ClaimsProviderConstants.CONFIGURATION_NAME, initializeRuntimeSettings);
            AADEntityProviderConfig<IAADSettings> configuration = (AADEntityProviderConfig<IAADSettings>)AADEntityProviderConfig<IAADSettings>.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID), initializeRuntimeSettings);
            return configuration;
        }

        public static AADEntityProviderConfig<IAADSettings> CreateConfiguration()
        {
            //AzureADEntityProviderConfiguration configuration = EntityProviderBase<AzureADEntityProviderConfiguration>.CreateGlobalConfiguration(ClaimsProviderConstants.CONFIGURATION_ID, ClaimsProviderConstants.CONFIGURATION_NAME, AzureCP.ClaimsProviderName);
            AADEntityProviderConfig<IAADSettings> configuration = (AADEntityProviderConfig<IAADSettings>)AADEntityProviderConfig<IAADSettings>.CreateGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID), ClaimsProviderConstants.CONFIGURATION_NAME, AzureCP.ClaimsProviderName, typeof(AADEntityProviderConfig<IAADSettings>));
            return configuration;
        }

        public static void DeleteConfiguration()
        {
            AADEntityProviderConfig<IAADSettings> configuration = (AADEntityProviderConfig<IAADSettings>)AADEntityProviderConfig<IAADSettings>.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID));
            if (configuration != null)
            {
                configuration.Delete();
            }
        }

        public bool ValidateLocalConfiguration(Uri context)
        {
            if (!Utils.ShouldRun(context, Name))
            {
                return false;
            }

            bool success = true;
            this.Lock_LocalConfigurationRefresh.EnterWriteLock();
            try
            {
                if (this.PersistedConfiguration == null)
                {
                    this.PersistedConfiguration = (AADEntityProviderConfig<IAADSettings>)AADEntityProviderConfig<IAADSettings>.GetGlobalConfiguration(new Guid(ClaimsProviderConstants.CONFIGURATION_ID));
                }
                if (this.PersistedConfiguration != null)
                {
                    //LocalConfiguration = this.EntityProvider.RefreshLocalConfigurationIfNeeded(ClaimsProviderConstants.CONFIGURATION_NAME);
                    LocalConfiguration = this.PersistedConfiguration.RefreshLocalConfigurationIfNeeded();
                }
                else
                {
                    success = false;
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
        /// Search or validate incoming input or entity
        /// </summary>
        /// <param name="currentContext">Information about current context and operation</param>
        /// <returns>Entities generated by AzureCP</returns>
        protected List<PickerEntity> SearchOrValidate(OperationContext currentContext)
        {
            List<DirectoryObject> azureADEntityList = null;
            List<PickerEntity> pickerEntityList = new List<PickerEntity>();
            try
            {
                if (this.LocalConfiguration.AlwaysResolveUserInput)
                {
                    // Completely bypass query to Azure AD
                    pickerEntityList = CreatePickerEntityForSpecificClaimTypes(
                        currentContext.Input,
                        currentContext.CurrentClaimTypeConfigList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject),
                        false);
                    Logger.Log($"[{Name}] Created {pickerEntityList.Count} entity(ies) without contacting Azure AD tenant(s) because AzureCP property AlwaysResolveUserInput is set to true.",
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
                            Logger.Log($"[{Name}] Created entity without contacting Azure AD tenant(s) because input started with prefix '{ctConfigWithInputPrefixMatch.PrefixToBypassLookup}', which is configured for claim type '{ctConfigWithInputPrefixMatch.ClaimType}'. Claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
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
                        // At this stage, it is impossible to know if entity was originally created with the keyword that bypass query to Azure AD
                        // But it should be always validated since property PrefixToBypassLookup is set for current ClaimTypeConfig, so create entity manually
                        PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                            currentContext.IncomingEntity.Value,
                            currentContext.IncomingEntityClaimTypeConfig,
                            currentContext.InputHasKeyword);
                        if (entity != null)
                        {
                            pickerEntityList = new List<PickerEntity>(1) { entity };
                            Logger.Log($"[{Name}] Validated entity without contacting Azure AD tenant(s) because its claim type ('{currentContext.IncomingEntityClaimTypeConfig.ClaimType}') has property 'PrefixToBypassLookup' set in AzureCPConfig.ClaimTypes. Claim value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'",
                                TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in SearchOrValidate", TraceCategory.Claims_Picking, ex);
            }
            pickerEntityList = this.ValidateEntities(currentContext, pickerEntityList);
            return pickerEntityList;
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

        protected async Task<List<DirectoryObject>> SearchOrValidateInAzureADAsync(OperationContext currentContext)
        {
            using (new SPMonitoredScope($"[{Name}] Total time spent to query Azure AD tenant(s)", 1000))
            {
                List<DirectoryObject> results = await this.EntityProvider.SearchOrValidateEntitiesAsync(currentContext).ConfigureAwait(false);
                return results;
            }
        }

        protected virtual List<PickerEntity> ProcessAzureADResults(OperationContext currentContext, List<DirectoryObject> usersAndGroups)
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

                    if (this.LocalConfiguration.FilterSecurityEnabledGroupsOnly)
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
                            claimTypeConfigToCompare = this.LocalConfiguration.IdentityClaimTypeConfig;
                            if (String.Equals(((User)currentObject).UserType, ClaimsProviderConstants.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
                            {
                                // For Guest users, use the value set in property DirectoryObjectPropertyForGuestUsers
                                entityClaimValue = GetPropertyValue(currentObject, this.LocalConfiguration.IdentityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers.ToString());
                            }
                            else
                            {
                                // Get the value of the DirectoryObjectProperty linked to current directory object
                                entityClaimValue = GetPropertyValue(currentObject, claimTypeConfigToCompare.EntityProperty.ToString());
                            }
                        }
                        else
                        {
                            claimTypeConfigToCompare = this.LocalConfiguration.MainGroupClaimTypeConfig;
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

        protected virtual PickerEntity CreatePickerEntityHelper(ClaimsProviderEntityResult result)
        {
            PickerEntity entity = CreatePickerEntity();
            SPClaim claim;
            string permissionValue = result.PermissionValue;
            string permissionClaimType = result.ClaimTypeConfig.ClaimType;
            bool isMappedClaimTypeConfig = false;

            if (String.Equals(result.ClaimTypeConfig.ClaimType, this.LocalConfiguration.IdentityClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase)
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
                    permissionClaimType = this.LocalConfiguration.IdentityClaimTypeConfig.ClaimType;
                    claimValueType = this.LocalConfiguration.IdentityClaimTypeConfig.ClaimValueType;
                    if (String.IsNullOrEmpty(entity.EntityType))
                    {
                        entity.EntityType = SPClaimEntityTypes.User;
                    }
                }
                else
                {
                    permissionClaimType = this.LocalConfiguration.MainGroupClaimTypeConfig.ClaimType;
                    claimValueType = this.LocalConfiguration.MainGroupClaimTypeConfig.ClaimValueType;
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
                foreach (ClaimTypeConfig ctConfig in this.LocalConfiguration.RuntimeMetadataConfig.Where(x => x.EntityType == result.ClaimTypeConfig.EntityType))
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
                bool isIdentityClaimType = String.Equals(claim.ClaimType, this.LocalConfiguration.IdentityClaimTypeConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase);
                entity.DisplayText = FormatPermissionDisplayText(entity, isIdentityClaimType, result);

                entities.Add(entity);
                Logger.Log($"[{Name}] Created entity: display text: '{entity.DisplayText}', value: '{entity.Claim.Value}', claim type: '{entity.Claim.ClaimType}'.", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Claims_Picking);
            }
            return entities.Count > 0 ? entities : null;
        }

        /// <summary>
        /// Override this method to customize value of permission created
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
            string entityDisplayText = this.LocalConfiguration.EntityDisplayTextPrefix;
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
            bool configIsValid = ValidateLocalConfiguration(null);
            if (configIsValid)
            {
                this.Lock_LocalConfigurationRefresh.EnterReadLock();
                try
                {

                    foreach (var claimTypeSettings in this.LocalConfiguration.RuntimeClaimTypesList)
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
        /// Perform augmentation of entity supplied
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

            if (!ValidateLocalConfiguration(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                // There can be multiple TrustedProvider on the farm, but AzureCP should only do augmentation if current entity is from TrustedProvider it is associated with
                if (!String.Equals(decodedEntity.OriginalIssuer, this.PersistedConfiguration.OriginalIssuerName, StringComparison.InvariantCultureIgnoreCase)) { return; }

                if (!this.LocalConfiguration.EnableAugmentation) { return; }

                Logger.Log($"[{Name}] Starting augmentation for user '{decodedEntity.Value}'.", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                ClaimTypeConfig groupClaimTypeSettings = this.LocalConfiguration.RuntimeClaimTypesList.FirstOrDefault(x => x.EntityType == DirectoryObjectType.Group);
                if (groupClaimTypeSettings == null)
                {
                    Logger.Log($"[{Name}] No claim type with EntityType 'Group' was found, please check claims mapping table.",
                        TraceSeverity.High, EventSeverity.Error, TraceCategory.Augmentation);
                    return;
                }

                OperationContext currentContext = new OperationContext(this.LocalConfiguration, OperationType.Augmentation, null, decodedEntity, context, null, null, Int32.MaxValue);
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
                    Logger.Log($"[{Name}] User '{currentContext.IncomingEntity.Value}' was augmented with {groups.Count} groups in {timer.ElapsedMilliseconds} ms",
                        TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Augmentation);
                }
                else
                {
                    Logger.Log($"[{Name}] No group found for user '{currentContext.IncomingEntity.Value}', search took {timer.ElapsedMilliseconds.ToString()} ms",
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
            return new SPClaim(type, value, valueType, this.PersistedConfiguration.OriginalIssuerName);
        }

        protected override void FillHierarchy(Uri context, string[] entityTypes, string hierarchyNodeID, int numberOfLevels, SPProviderHierarchyTree hierarchy)
        {
            List<DirectoryObjectType> aadEntityTypes = new List<DirectoryObjectType>();
            if (entityTypes.Contains(SPClaimEntityTypes.User)) { aadEntityTypes.Add(DirectoryObjectType.User); }
            if (entityTypes.Contains(ClaimsProviderConstants.GroupClaimEntityType)) { aadEntityTypes.Add(DirectoryObjectType.Group); }

            if (!ValidateLocalConfiguration(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                if (hierarchyNodeID == null)
                {
                    // Root level
                    foreach (var azureObject in this.LocalConfiguration.RuntimeClaimTypesList.FindAll(x => !x.UseMainClaimTypeOfDirectoryObject && aadEntityTypes.Contains(x.EntityType)))
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
            if (!ValidateLocalConfiguration(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                OperationContext currentContext = new OperationContext(this.LocalConfiguration, OperationType.Search, resolveInput, null, context, entityTypes, null, 30);
                List<PickerEntity> entities = SearchOrValidate(currentContext);
                FillEntities(currentContext, ref entities);
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
            if (!ValidateLocalConfiguration(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                // Ensure incoming claim should be validated by AzureCP
                // Must be made after call to Initialize because SPTrustedLoginProvider name must be known
                if (!String.Equals(resolveInput.OriginalIssuer, this.PersistedConfiguration.OriginalIssuerName, StringComparison.InvariantCultureIgnoreCase)) { return; }

                OperationContext currentContext = new OperationContext(this.LocalConfiguration, OperationType.Validation, resolveInput.Value, resolveInput, context, entityTypes, null, 1);
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
            if (!ValidateLocalConfiguration(context)) { return; }

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                OperationContext currentContext = new OperationContext(this.LocalConfiguration, OperationType.Search, searchPattern, null, context, entityTypes, hierarchyNodeID, maxCount);
                List<PickerEntity> entities = this.SearchOrValidate(currentContext);
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
                        ClaimTypeConfig ctConfig = this.LocalConfiguration.RuntimeClaimTypesList.FirstOrDefault(x =>
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
        /// Override this method to change / remove entities created by AzureCP, or add new ones
        /// </summary>
        /// <param name="currentContext"></param>
        /// <param name="resolved">List of entities created by LDAPCP</param>
        protected virtual void FillEntities(OperationContext currentContext, ref List<PickerEntity> resolved)
        {
        }

        /// <summary>
        /// Return the identity claim type
        /// </summary>
        /// <returns></returns>
        public override string GetClaimTypeForUserKey()
        {
            // Initialization may fail because there is no yet configuration (fresh install)
            // In this case, AzureCP should not return null because it causes null exceptions in SharePoint when users sign-in
            bool configIsValid = ValidateLocalConfiguration(null);
            if (configIsValid)
            {
                this.Lock_LocalConfigurationRefresh.EnterReadLock();
                try
                {
                    return this.PersistedConfiguration.SPTrust.IdentityClaimTypeInformation.MappedClaimType;
                }
                catch (Exception ex)
                {
                    Logger.LogException(Name, "in GetClaimTypeForUserKey", TraceCategory.Rehydration, ex);
                }
                finally
                {
                    this.Lock_LocalConfigurationRefresh.ExitReadLock();
                }
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
            // Initialization may fail because there is no yet configuration (fresh install)
            // In this case, AzureCP should not return null because it causes null exceptions in SharePoint when users sign-in
            bool initSucceeded = ValidateLocalConfiguration(null);

            this.Lock_LocalConfigurationRefresh.EnterReadLock();
            try
            {
                // If initialization failed but SPTrust is not null, rest of the method can be executed normally
                // Otherwise return the entity
                if (!initSucceeded && this.PersistedConfiguration?.SPTrust == null)
                {
                    return entity;
                }

                // There are 2 scenarios:
                // 1: OriginalIssuer is "SecurityTokenService": Value looks like "05.t|contoso.local|yvand@contoso.local", claim type is "http://schemas.microsoft.com/sharepoint/2009/08/claims/userid" and it must be decoded properly
                // 2: OriginalIssuer is AzureCP: in this case incoming entity is valid and returned as is
                if (String.Equals(entity.OriginalIssuer, this.PersistedConfiguration.SPTrust.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    return entity;
                }

                SPClaimProviderManager cpm = SPClaimProviderManager.Local;
                SPClaim curUser = SPClaimProviderManager.DecodeUserIdentifierClaim(entity);

                Logger.Log($"[{Name}] Returning user key for '{entity.Value}'",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Rehydration);
                return CreateClaim(this.PersistedConfiguration.SPTrust.IdentityClaimTypeInformation.MappedClaimType, curUser.Value, curUser.ValueType);
            }
            catch (Exception ex)
            {
                Logger.LogException(Name, "in GetUserKeyForEntity", TraceCategory.Rehydration, ex);
            }
            finally
            {
                this.Lock_LocalConfigurationRefresh.ExitReadLock();
            }
            return null;
        }
    }

    /// <summary>
    /// User / group found in Azure AD, with additional information
    /// </summary>
    public class ClaimsProviderEntityResult
    {
        /// <summary>
        /// Gets the entity returned by Azure AD
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
