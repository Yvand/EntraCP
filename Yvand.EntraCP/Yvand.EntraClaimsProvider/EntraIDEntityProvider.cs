﻿using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Groups;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users;
using Microsoft.Graph.Users.Item.GetMemberGroups;
using Microsoft.Identity.Client;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Threading.Tasks;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider
{
    public class EntraIDEntityProvider : EntityProviderBase
    {
        public EntraIDEntityProvider(string claimsProviderName) : base(claimsProviderName) { }

        public async override Task<List<string>> GetEntityGroupsAsync(OperationContext currentContext, DirectoryObjectProperty groupProperty)
        {
            List<EntraIDTenant> azureTenants = currentContext.AzureTenants;
            // URL encode the filter to prevent that it gets truncated like this: "UserPrincipalName eq 'guest_contoso.com" instead of "UserPrincipalName eq 'guest_contoso.com#EXT#@TENANT.onmicrosoft.com'"
            string getMemberUserFilter = $"{currentContext.IncomingEntityClaimTypeConfig.EntityProperty} eq '{currentContext.IncomingEntity.Value}'";
            string getGuestUserFilter = $"userType eq 'Guest' and {currentContext.Settings.IdentityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers} eq '{currentContext.IncomingEntity.Value}'";

            // Create a task for each tenant to query
            IEnumerable<Task<List<string>>> tenantTasks = azureTenants.Select(async tenant =>
            {
                List<string> groupsInTenant = new List<string>();
                Stopwatch timer = new Stopwatch();
                timer.Start();
                try
                {
                    // Search the user as a member
                    var userCollectionResult = await tenant.GraphService.Users.GetAsync((config) =>
                    {
                        config.QueryParameters.Filter = getMemberUserFilter;
                        config.QueryParameters.Select = new[] { "Id" };
                        config.QueryParameters.Top = 1;
                    }).ConfigureAwait(false);

                    User user = userCollectionResult?.Value?.FirstOrDefault();
                    if (user == null)
                    {
                        // If user was not found, he might be a Guest user. Query to check this: /users?$filter=userType eq 'Guest' and mail eq 'guest@live.com'&$select=userPrincipalName, Id
                        //string guestFilter = HttpUtility.UrlEncode($"userType eq 'Guest' and {IdentityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers} eq '{currentContext.IncomingEntity.Value}'");
                        //userResult = await tenant.GraphService.Users.Request().Filter(guestFilter).Select(HttpUtility.UrlEncode("userPrincipalName, Id")).GetAsync().ConfigureAwait(false);
                        //userResult = await Task.Run(() => tenant.GraphService.Users.Request().Filter(guestFilter).Select(HttpUtility.UrlEncode("userPrincipalName, Id")).GetAsync()).ConfigureAwait(false);
                        userCollectionResult = await Task.Run(() => tenant.GraphService.Users.GetAsync((config) =>
                        {
                            config.QueryParameters.Filter = getGuestUserFilter;
                            config.QueryParameters.Select = new[] { "Id" };
                            config.QueryParameters.Top = 1;
                        })).ConfigureAwait(false);
                        user = userCollectionResult?.Value?.FirstOrDefault();
                    }
                    if (user == null) { return groupsInTenant; }

                    if (groupProperty == DirectoryObjectProperty.Id)
                    {
                        // POST to /v1.0/users/user@TENANT.onmicrosoft.com/microsoft.graph.getMemberGroups is the preferred way to return security groups as it includes nested groups
                        // But it returns only the group IDs so it can be used only if groupClaimTypeConfig.DirectoryObjectProperty == AzureADObjectProperty.Id
                        // For Guest users, it must be the id: POST to /v1.0/users/18ff6ae9-dd01-4008-a786-aabf71f1492a/microsoft.graph.getMemberGroups
                        GetMemberGroupsPostRequestBody getGroupsOptions = new GetMemberGroupsPostRequestBody { SecurityEnabledOnly = currentContext.Settings.FilterSecurityEnabledGroupsOnly };
                        GetMemberGroupsResponse memberGroupsResponse = await Task.Run(() => tenant.GraphService.Users[user.Id].GetMemberGroups.PostAsync(getGroupsOptions)).ConfigureAwait(false);
                        if (memberGroupsResponse?.Value != null)
                        {
                            PageIterator<string, GetMemberGroupsResponse> memberGroupsPageIterator = PageIterator<string, GetMemberGroupsResponse>.CreatePageIterator(
                            tenant.GraphService,
                            memberGroupsResponse,
                            (groupId) =>
                            {
                                groupsInTenant.Add(groupId);
                                return true; // return true to continue the iteration
                            });
                            await memberGroupsPageIterator.IterateAsync().ConfigureAwait(false);
                        }
                    }
                    else
                    {
                        // Fallback to GET to /v1.0/users/user@TENANT.onmicrosoft.com/memberOf, which returns all group properties but does not return nested groups
                        DirectoryObjectCollectionResponse memberOfResponse = await Task.Run(() => tenant.GraphService.Users[user.Id].MemberOf.GetAsync()).ConfigureAwait(false);
                        if (memberOfResponse?.Value != null)
                        {
                            PageIterator<Group, DirectoryObjectCollectionResponse> memberGroupsPageIterator = PageIterator<Group, DirectoryObjectCollectionResponse>.CreatePageIterator(
                            tenant.GraphService,
                            memberOfResponse,
                            (group) =>
                            {
                                string groupClaimValue = GetPropertyValue(group, groupProperty.ToString());
                                groupsInTenant.Add(groupClaimValue);
                                return true; // return true to continue the iteration
                            });
                            await memberGroupsPageIterator.IterateAsync().ConfigureAwait(false);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogException(ClaimsProviderName, $"while getting groups for user '{currentContext.IncomingEntity.Value}' from tenant '{tenant.Name}'", TraceCategory.Augmentation, ex);
                }
                finally
                {
                    timer.Stop();
                }
                if (groupsInTenant != null)
                {
                    Logger.Log($"[{ClaimsProviderName}] Got {groupsInTenant.Count} groups in {timer.ElapsedMilliseconds} ms for user '{currentContext.IncomingEntity.Value}' from tenant '{tenant.Name}'", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                }
                else
                {
                    Logger.Log($"[{ClaimsProviderName}] Got no group in {timer.ElapsedMilliseconds} ms for user '{currentContext.IncomingEntity.Value}' from tenant '{tenant.Name}'", TraceSeverity.Verbose, EventSeverity.Information, TraceCategory.Augmentation);
                }
                return groupsInTenant;
            });

            List<string> groups = new List<string>();
            // Wait for all tasks to complete
            List<string>[] groupsInAllTenants = await Task.WhenAll(tenantTasks).ConfigureAwait(false);
            for (int i = 0; i < groupsInAllTenants.Length; i++)
            {
                groups.AddRange(groupsInAllTenants[i]);
            }
            return groups;
        }

        public async override Task<List<DirectoryObject>> SearchOrValidateEntitiesAsync(OperationContext currentContext)
        {
            //// this.CurrentConfiguration.EntraIDTenants must be cloned locally to ensure its properties ($select / $filter) won't be updated by multiple threads
            //List<EntraIDTenant> azureTenants = new List<EntraIDTenant>(this.Configuration.EntraIDTenants.Count);
            //foreach (EntraIDTenant tenant in this.Configuration.EntraIDTenants)
            //{
            //    azureTenants.Add(tenant.CopyPublicProperties());
            //}
            this.BuildFilter(currentContext, currentContext.AzureTenants);
            List<DirectoryObject> results = await this.QueryEntraIDTenantsAsync(currentContext, currentContext.AzureTenants);
            return results;
        }

        protected virtual void BuildFilter(OperationContext currentContext, List<EntraIDTenant> azureTenants)
        {
            string searchPatternEquals = "{0} eq '{1}'";
            string searchPatternStartsWith = "startswith({0}, '{1}')";
            string identityConfigSearchPatternEquals = "({0} eq '{1}' and UserType eq '{2}')";
            string identityConfigSearchPatternStartsWith = "(startswith({0}, '{1}') and UserType eq '{2}')";

            List<string> userFilterBuilder = new List<string>();
            List<string> groupFilterBuilder = new List<string>();
            List<string> userSelectBuilder = new List<string> { "UserType", "Mail" };    // UserType and Mail are always needed to deal with Guest users
            List<string> groupSelectBuilder = new List<string> { "Id", "securityEnabled" };               // Id is always required for groups

            string filterPattern;
            string input = currentContext.Input;

            // https://github.com/Yvand/AzureCP/issues/88: Escape single quotes as documented in https://docs.microsoft.com/en-us/graph/query-parameters#escaping-single-quotes
            input = input.Replace("'", "''");

            if (currentContext.ExactSearch)
            {
                filterPattern = String.Format(searchPatternEquals, "{0}", input);
            }
            else
            {
                filterPattern = String.Format(searchPatternStartsWith, "{0}", input);
            }

            foreach (ClaimTypeConfig ctConfig in currentContext.CurrentClaimTypeConfigList)
            {
                string currentPropertyString = ctConfig.EntityProperty.ToString();
                if (currentPropertyString.StartsWith("extensionAttribute"))
                {
                    currentPropertyString = String.Format("{0}_{1}_{2}", "extension", "EXTENSIONATTRIBUTESAPPLICATIONID", currentPropertyString);
                }

                string currentFilter;
                if (!ctConfig.SupportsWildcard)
                {
                    currentFilter = String.Format(searchPatternEquals, currentPropertyString, input);
                }
                else
                {
                    // Use String.Replace instead of String.Format because String.Format trows an exception if input contains a '{'
                    currentFilter = filterPattern.Replace("{0}", currentPropertyString);
                }

                // Id needs a specific check: input must be a valid GUID AND equals filter must be used, otherwise Microsoft Entra ID will throw an error
                if (ctConfig.EntityProperty == DirectoryObjectProperty.Id)
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
                            currentFilter = "( " + String.Format(identityConfigSearchPatternEquals, currentPropertyString, input, ClaimsProviderConstants.MEMBER_USERTYPE) + " or " + String.Format(identityConfigSearchPatternEquals, identityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers, input, ClaimsProviderConstants.GUEST_USERTYPE) + " )";
                        }
                        else
                        {
                            if (currentContext.ExactSearch)
                            {
                                currentFilter = "( " + String.Format(identityConfigSearchPatternEquals, currentPropertyString, input, ClaimsProviderConstants.MEMBER_USERTYPE) + " or " + String.Format(identityConfigSearchPatternEquals, identityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers, input, ClaimsProviderConstants.GUEST_USERTYPE) + " )";
                            }
                            else
                            {
                                currentFilter = "( " + String.Format(identityConfigSearchPatternStartsWith, currentPropertyString, input, ClaimsProviderConstants.MEMBER_USERTYPE) + " or " + String.Format(identityConfigSearchPatternStartsWith, identityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers, input, ClaimsProviderConstants.GUEST_USERTYPE) + " )";
                            }
                        }
                    }

                    userFilterBuilder.Add(currentFilter);
                    userSelectBuilder.Add(currentPropertyString);
                }
                else
                {
                    // else assume it's a Group
                    groupFilterBuilder.Add(currentFilter);
                    groupSelectBuilder.Add(currentPropertyString);
                }
            }

            // Also add metadata properties to $select of corresponding object type
            if (userFilterBuilder.Count > 0)
            {
                foreach (ClaimTypeConfig ctConfig in currentContext.Settings.RuntimeMetadataConfig.Where(x => x.EntityType == DirectoryObjectType.User))
                {
                    userSelectBuilder.Add(ctConfig.EntityProperty.ToString());
                }
            }
            if (groupFilterBuilder.Count > 0)
            {
                foreach (ClaimTypeConfig ctConfig in currentContext.Settings.RuntimeMetadataConfig.Where(x => x.EntityType == DirectoryObjectType.Group))
                {
                    groupSelectBuilder.Add(ctConfig.EntityProperty.ToString());
                }
            }

            foreach (EntraIDTenant tenant in azureTenants)
            {
                List<string> userFilterBuilderForTenantList;
                List<string> groupFilterBuilderForTenantList;
                List<string> userSelectBuilderForTenantList;
                List<string> groupSelectBuilderForTenantList;

                // Add extension attribute on current tenant only if it is configured for it, otherwise request fails with this error:
                // message=Property 'extension_00000000000000000000000000000000_extensionAttribute1' does not exist as a declared property or extension property.
                if (tenant.ExtensionAttributesApplicationId == Guid.Empty)
                {
                    userFilterBuilderForTenantList = userFilterBuilder.FindAll(elem => !elem.Contains("EXTENSIONATTRIBUTESAPPLICATIONID"));
                    groupFilterBuilderForTenantList = groupFilterBuilder.FindAll(elem => !elem.Contains("EXTENSIONATTRIBUTESAPPLICATIONID"));
                    userSelectBuilderForTenantList = userSelectBuilder.FindAll(elem => !elem.Contains("EXTENSIONATTRIBUTESAPPLICATIONID"));
                    groupSelectBuilderForTenantList = groupSelectBuilder.FindAll(elem => !elem.Contains("EXTENSIONATTRIBUTESAPPLICATIONID"));
                }
                else
                {
                    userFilterBuilderForTenantList = userFilterBuilder.Select(elem => elem.Replace("EXTENSIONATTRIBUTESAPPLICATIONID", tenant.ExtensionAttributesApplicationId.ToString("N"))).ToList<string>();
                    groupFilterBuilderForTenantList = groupFilterBuilder.Select(elem => elem.Replace("EXTENSIONATTRIBUTESAPPLICATIONID", tenant.ExtensionAttributesApplicationId.ToString("N"))).ToList<string>();
                    userSelectBuilderForTenantList = userSelectBuilder.Select(elem => elem.Replace("EXTENSIONATTRIBUTESAPPLICATIONID", tenant.ExtensionAttributesApplicationId.ToString("N"))).ToList<string>();
                    groupSelectBuilderForTenantList = groupSelectBuilder.Select(elem => elem.Replace("EXTENSIONATTRIBUTESAPPLICATIONID", tenant.ExtensionAttributesApplicationId.ToString("N"))).ToList<string>();
                }

                if (userFilterBuilder.Count > 0)
                {
                    tenant.UserFilter = String.Join(" or ", userFilterBuilderForTenantList);
                }
                else
                {
                    // Reset filter if no corresponding object was found in requestInfo.ClaimTypeConfigList, to detect that tenant should not be queried
                    tenant.UserFilter = String.Empty;
                }

                if (groupFilterBuilder.Count > 0)
                {
                    tenant.GroupFilter = String.Join(" or ", groupFilterBuilderForTenantList);
                }
                else
                {
                    tenant.GroupFilter = String.Empty;
                }

                tenant.UserSelect = userSelectBuilderForTenantList.ToArray();
                tenant.GroupSelect = groupSelectBuilderForTenantList.ToArray();
            }
        }

        protected async Task<List<DirectoryObject>> QueryEntraIDTenantsAsync(OperationContext currentContext, List<EntraIDTenant> azureTenants)
        {
            // Create a task for each tenant to query
            var tenantQueryTasks = azureTenants.Select(async tenant =>
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                List<DirectoryObject> tenantResults = await QueryEntraIDTenantAsync(currentContext, tenant).ConfigureAwait(false);
                timer.Stop();
                if (tenantResults != null)
                {
                    Logger.Log($"[{ClaimsProviderName}] Got {tenantResults.Count} users/groups in {timer.ElapsedMilliseconds.ToString()} ms from '{tenant.Name}' with input '{currentContext.Input}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                }
                else
                {
                    Logger.Log($"[{ClaimsProviderName}] Got no result from '{tenant.Name}' with input '{currentContext.Input}', search took {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                }
                return tenantResults;
            });

            // Wait for all tasks to complete
            List<DirectoryObject> allResults = new List<DirectoryObject>();
            List<DirectoryObject>[] tenantsResults = await Task.WhenAll(tenantQueryTasks).ConfigureAwait(false);
            for (int i = 0; i < tenantsResults.Length; i++)
            {
                // If request to Graph failed, tenantsResults[i] is null and that would cause a ThrowArgumentNullException in List<T>.InsertRange()
                if (tenantsResults[i] != null && tenantsResults[i].Count > 0)
                {
                    allResults.AddRange(tenantsResults[i]);
                }
            }
            return allResults;
        }

        protected virtual async Task<List<DirectoryObject>> QueryEntraIDTenantAsync(OperationContext currentContext, EntraIDTenant tenant)
        {
            List<DirectoryObject> tenantResults = new List<DirectoryObject>();
            if (String.IsNullOrWhiteSpace(tenant.UserFilter) && String.IsNullOrWhiteSpace(tenant.GroupFilter))
            {
                return tenantResults;
            }

            if (tenant.GraphService == null)
            {
                Logger.Log($"[{ClaimsProviderName}] Cannot query Microsoft Entra ID tenant '{tenant.Name}' because it was not initialized", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
                return tenantResults;
            }

            Logger.Log($"[{ClaimsProviderName}] Querying Microsoft Entra ID tenant '{tenant.Name}' for users and groups, with input '{currentContext.Input}'", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Lookup);
            object lockAddResultToCollection = new object();
            int timeout = currentContext.Settings.Timeout;
            int maxRetry = currentContext.OperationType == OperationType.Validation ? 3 : 2;
            int tenantResultCount = 0;

            try
            {
                using (new SPMonitoredScope($"[{ClaimsProviderName}] Querying Microsoft Entra ID tenant '{tenant.Name}' for users and groups, with input '{currentContext.Input}'", 1000))
                {
                    RetryHandlerOption retryHandlerOption = new RetryHandlerOption()
                    {
                        Delay = 1,
                        RetriesTimeLimit = TimeSpan.FromMilliseconds(timeout),
                        MaxRetry = maxRetry,
                        ShouldRetry = (delay, attempt, httpResponse) =>
                        {
                            // Pointless to retry if this is Unauthorized
                            if (httpResponse.StatusCode == System.Net.HttpStatusCode.Unauthorized)
                            {
                                return false;
                            }
                            return true;
                        }
                    };

                    // Build the batch
                    BatchRequestContentCollection batchRequestContent = new BatchRequestContentCollection(tenant.GraphService);
                    string usersRequestId = String.Empty;
                    if (!String.IsNullOrWhiteSpace(tenant.UserFilter))
                    {
                        // https://stackoverflow.com/questions/56417435/when-i-set-an-object-using-an-action-the-object-assigned-is-always-null
                        RequestInformation userRequest = tenant.GraphService.Users.ToGetRequestInformation(conf =>
                        {
                            conf.QueryParameters = new UsersRequestBuilder.UsersRequestBuilderGetQueryParameters
                            {
                                Count = true,
                                Filter = tenant.UserFilter,
                                Select = tenant.UserSelect,
                                Top = currentContext.MaxCount,
                            };
                            conf.Headers = new RequestHeaders
                            {
                                // Allow Advanced query as documented in  https://learn.microsoft.com/en-us/graph/sdks/create-requests?tabs=csharp#retrieve-a-list-of-entities
                                //to fix $filter on CompanyName - https://github.com/Yvand/AzureCP/issues/166
                                { "ConsistencyLevel", "eventual" }
                            };
                            conf.Options = new List<IRequestOption>
                            {
                                retryHandlerOption,
                            };
                        });
                        // Using AddBatchRequestStepAsync adds each request as a step with no specified order of execution
                        usersRequestId = await batchRequestContent.AddBatchRequestStepAsync(userRequest).ConfigureAwait(false);
                    }

                    // Groups
                    string groupsRequestId = String.Empty;
                    if (!String.IsNullOrWhiteSpace(tenant.GroupFilter))
                    {
                        RequestInformation groupRequest = tenant.GraphService.Groups.ToGetRequestInformation(conf =>
                        {
                            conf.QueryParameters = new GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters
                            {
                                Count = true,
                                Filter = tenant.GroupFilter,
                                Select = tenant.GroupSelect,
                                Top = currentContext.MaxCount,
                            };
                            conf.Headers = new RequestHeaders
                            {
                                // Allow Advanced query as documented in  https://learn.microsoft.com/en-us/graph/sdks/create-requests?tabs=csharp#retrieve-a-list-of-entities
                                //to fix $filter on CompanyName - https://github.com/Yvand/AzureCP/issues/166
                                { "ConsistencyLevel", "eventual" }
                            };
                            conf.Options = new List<IRequestOption>
                            {
                                retryHandlerOption,
                            };
                        });
                        // Using AddBatchRequestStepAsync adds each request as a step with no specified order of execution
                        groupsRequestId = await batchRequestContent.AddBatchRequestStepAsync(groupRequest).ConfigureAwait(false);
                    }

                    // Run the batch request and get the HTTP status code of each request inside the batch
                    BatchResponseContentCollection batchResponse = await tenant.GraphService.Batch.PostAsync(batchRequestContent).ConfigureAwait(false);
                    Dictionary<string, HttpStatusCode> requestsStatusInBatchResponse = await batchResponse.GetResponsesStatusCodesAsync().ConfigureAwait(false);

                    // Check if the users' request in the batch request was successful. If so, get the users that were returned by Graph
                    HttpStatusCode usersRequestStatus;
                    UserCollectionResponse userCollectionResult = null;
                    if (requestsStatusInBatchResponse.TryGetValue(usersRequestId, out usersRequestStatus))
                    {
                        if (usersRequestStatus == HttpStatusCode.OK)
                        {
                            userCollectionResult = await batchResponse.GetResponseByIdAsync<UserCollectionResponse>(usersRequestId).ConfigureAwait(false);
                        }
                        else
                        {
                            Logger.Log($"[{ClaimsProviderName}] Query to tenant '{tenant.Name}' returned unexpected status '{usersRequestStatus}' for users request with filter \"{tenant.UserFilter}\"", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
                        }
                    }

                    // Check if the groups' request in the batch request was successful. If so, get the users that were returned by Graph
                    HttpStatusCode groupRequestStatus;
                    GroupCollectionResponse groupCollectionResult = null;
                    if (requestsStatusInBatchResponse.TryGetValue(groupsRequestId, out groupRequestStatus))
                    {
                        if (groupRequestStatus == HttpStatusCode.OK)
                        {
                            groupCollectionResult = await batchResponse.GetResponseByIdAsync<GroupCollectionResponse>(groupsRequestId).ConfigureAwait(false);
                        }
                        else
                        {
                            Logger.Log($"[{ClaimsProviderName}] Query to tenant '{tenant.Name}' returned unexpected status '{groupRequestStatus}' for groups request with filter \"{tenant.GroupFilter}\"", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
                        }
                    }

                    Logger.Log($"[{ClaimsProviderName}] Query to tenant '{tenant.Name}' returned {(userCollectionResult?.Value == null ? 0 : userCollectionResult.Value.Count)} user(s) with filter \"{tenant.UserFilter}\" and {(groupCollectionResult?.Value == null ? 0 : groupCollectionResult.Value.Count)} group(s) with filter \"{tenant.GroupFilter}\"", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Lookup);
                    // Process users result
                    if (userCollectionResult?.Value != null)
                    {
                        PageIterator<User, UserCollectionResponse> usersPageIterator = PageIterator<User, UserCollectionResponse>.CreatePageIterator(
                            tenant.GraphService,
                            userCollectionResult,
                            (user) =>
                            {
                                bool addUser = false;
                                if (tenant.ExcludeMemberUsers == false || tenant.ExcludeGuestUsers == false)
                                {
                                    bool userIsAMember = String.Equals(user.UserType, ClaimsProviderConstants.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase);
                                    bool userIsAGuest = !userIsAMember;

                                    if (tenant.ExcludeMemberUsers == false && tenant.ExcludeGuestUsers == false)
                                    {
                                        addUser = true;
                                    }
                                    else if (tenant.ExcludeMemberUsers == true && userIsAMember == false
                                     || tenant.ExcludeGuestUsers == true && userIsAGuest == false)
                                    {
                                        addUser = true;
                                    }
                                }

                                bool continueIteration = true;
                                if (addUser)
                                {
                                    lock (lockAddResultToCollection)
                                    {
                                        if (tenantResultCount < currentContext.MaxCount)
                                        {
                                            tenantResults.Add(user);
                                            tenantResultCount++;
                                        }
                                        else
                                        {
                                            continueIteration = false;
                                        }
                                    }
                                }
                                return continueIteration; // return true to continue the iteration
                            });
                        await usersPageIterator.IterateAsync().ConfigureAwait(false);
                    }

                    // Process groups result
                    if (groupCollectionResult?.Value != null)
                    {
                        PageIterator<Group, GroupCollectionResponse> groupsPageIterator = PageIterator<Group, GroupCollectionResponse>.CreatePageIterator(
                            tenant.GraphService,
                            groupCollectionResult,
                            (group) =>
                            {
                                bool continueIteration = true;
                                lock (lockAddResultToCollection)
                                {
                                    if (tenantResultCount < currentContext.MaxCount)
                                    {
                                        tenantResults.Add(group);
                                        tenantResultCount++;
                                    }
                                    else
                                    {
                                        continueIteration = false;
                                    }
                                }
                                return continueIteration; // return true to continue the iteration
                            });
                        await groupsPageIterator.IterateAsync().ConfigureAwait(false);
                    }

                    //// Cannot use Task.WaitAll() because it's actually blocking the threads, preventing parallel queries on others AAD tenants.
                    //// Use await Task.WhenAll() as it does not block other threads, so all AAD tenants are actually queried in parallel.
                    //// More info: https://stackoverflow.com/questions/12337671/using-async-await-for-multiple-tasks
                    //await Task.WhenAll(new Task[1] { batchQueryTask }).ConfigureAwait(false);
                    //ClaimsProviderLogging.LogDebug($"Waiting on Task.WaitAll for {tenant.Name} finished");
                }
            }
            catch (OperationCanceledException)
            {
                Logger.Log($"[{ClaimsProviderName}] Queries on Microsoft Entra ID tenant '{tenant.Name}' exceeded timeout of {timeout} ms and were cancelled.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
            }
            catch (AuthenticationFailedException ex)
            {
                Logger.LogException(ClaimsProviderName, $": Could not authenticate for tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
            }
            catch (MsalServiceException ex)
            {
                Logger.LogException(ClaimsProviderName, $": Msal could not query tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
            }
            catch (ServiceException ex)
            {
                Logger.LogException(ClaimsProviderName, $": Microsoft.Graph could not query tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
            }
            catch (AggregateException ex)
            {
                // Task.WaitAll throws an AggregateException, which contains all exceptions thrown by tasks it waited on
                Logger.LogException(ClaimsProviderName, $"while querying Microsoft Entra ID tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
            }
            catch (Exception ex)
            {
                Logger.LogException(ClaimsProviderName, $"in QueryEntraIDTenantAsync while querying tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
            }
            finally
            {
            }
            return tenantResults;
        }

        /// <summary>
        /// Uses reflection to return the value of a public property for the given object
        /// </summary>
        /// <param name="directoryObject"></param>
        /// <param name="propertyName"></param>
        /// <returns>Null if property doesn't exist, String.Empty if property exists but has no value, actual value otherwise</returns>
        public static string GetPropertyValue(DirectoryObject directoryObject, string propertyName)
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
                            else
                            {
                                return null;
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
                            else
                            {
                                return null;
                            }
                        }
                    }
                    return returnString == null ? propertyName : returnString;
                }
                catch
                {
                    return null;
                }
            }

            PropertyInfo pi = directoryObject.GetType().GetProperty(propertyName);
            if (pi == null)
            {
                return null;
            }   // Property doesn't exist
            object propertyValue = pi.GetValue(directoryObject, null);
            return propertyValue == null ? String.Empty : propertyValue.ToString();
        }
    }
}
