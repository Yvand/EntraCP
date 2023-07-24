using Microsoft.Graph;
using Microsoft.Graph.Groups;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yvand.ClaimsProviders.Configuration;
using Yvand.ClaimsProviders.Configuration.AzureAD;
using static Yvand.ClaimsProviders.ClaimsProviderLogging;

namespace Yvand.ClaimsProviders.AzureAD
{
    public class AzureADEntityProvider : EntityProviderBase<AzureADEntityProviderConfiguration>
    {
        public AzureADEntityProvider(string providerInternalName) : base(providerInternalName) { }        

        public async override Task<List<Group>> GetEntityGroupsAsync(OperationContext currentContext)
        {
            throw new NotImplementedException();
        }

        public async override Task<List<DirectoryObject>> SearchOrValidateEntitiesAsync(OperationContext currentContext)
        {
            // this.CurrentConfiguration.AzureTenants must be cloned locally var to ensure its properties ($select / $filter) won't be updated by multiple threads
            List<AzureTenant> azureTenants = new List<AzureTenant>(this.LocalConfiguration.AzureTenants.Count);
            foreach (AzureTenant tenant in this.LocalConfiguration.AzureTenants)
            {
                azureTenants.Add(tenant.CopyConfiguration());
            }
            this.BuildFilter(currentContext, azureTenants);
            List<DirectoryObject> results = await this.QueryAzureADTenantsAsync(currentContext, azureTenants);
            return results;
        }

        protected virtual void BuildFilter(OperationContext currentContext, List<AzureTenant> azureTenants)
        {
            string searchPatternEquals = "{0} eq '{1}'";
            string searchPatternStartsWith = "startswith({0}, '{1}')";
            string identityConfigSearchPatternEquals = "({0} eq '{1}' and UserType eq '{2}')";
            string identityConfigSearchPatternStartsWith = "(startswith({0}, '{1}') and UserType eq '{2}')";

            StringBuilder userFilterBuilder = new StringBuilder();
            StringBuilder groupFilterBuilder = new StringBuilder();
            List<string> userSelectBuilder = new List<string> { "UserType", "Mail" };    // UserType and Mail are always needed to deal with Guest users
            List<string> groupSelectBuilder = new List<string> { "Id", "securityEnabled" };               // Id is always required for groups

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

                    if (!firstUserObjectProcessed)
                    {
                        firstUserObjectProcessed = true;
                    }
                    else
                    {
                        currentFilter = " or " + currentFilter;
                    }
                    userFilterBuilder.Append(currentFilter);
                    userSelectBuilder.Add(currentPropertyString);
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
                    }
                    groupFilterBuilder.Append(currentFilter);
                    groupSelectBuilder.Add(currentPropertyString);
                }
            }

            // Also add metadata properties to $select of corresponding object type
            if (firstUserObjectProcessed)
            {
                foreach (ClaimTypeConfig ctConfig in LocalConfiguration.MetadataConfig.Where(x => x.EntityType == DirectoryObjectType.User))
                {
                    userSelectBuilder.Add(ctConfig.DirectoryObjectProperty.ToString());
                }
            }
            if (firstGroupObjectProcessed)
            {
                foreach (ClaimTypeConfig ctConfig in LocalConfiguration.MetadataConfig.Where(x => x.EntityType == DirectoryObjectType.Group))
                {
                    groupSelectBuilder.Add(ctConfig.DirectoryObjectProperty.ToString());
                }
            }

            foreach (AzureTenant tenant in azureTenants)
            {
                string userFilterForTenant = userFilterBuilder.ToString().Replace("EXTENSIONATTRIBUTESAPPLICATIONID", tenant.ExtensionAttributesApplicationId.ToString("N"));
                List<string> userSelectBuilderForTenant = userSelectBuilder.Select(elem => elem.Replace("EXTENSIONATTRIBUTESAPPLICATIONID", tenant.ExtensionAttributesApplicationId.ToString("N"))).ToList<string>();
                string groupFilterForTenant = groupFilterBuilder.ToString().Replace("EXTENSIONATTRIBUTESAPPLICATIONID", tenant.ExtensionAttributesApplicationId.ToString("N"));
                List<string> groupSelectBuilderForTenant = groupSelectBuilder.Select(elem => elem.Replace("EXTENSIONATTRIBUTESAPPLICATIONID", tenant.ExtensionAttributesApplicationId.ToString("N"))).ToList<string>();

                if (firstUserObjectProcessed)
                {
                    tenant.UserFilter = userFilterForTenant;
                }
                else
                {
                    // Reset filter if no corresponding object was found in requestInfo.ClaimTypeConfigList, to detect that tenant should not be queried
                    tenant.UserFilter = String.Empty;
                }

                if (firstGroupObjectProcessed)
                {
                    tenant.GroupFilter = groupFilterForTenant;
                }
                else
                {
                    tenant.GroupFilter = String.Empty;
                }

                tenant.UserSelect = userSelectBuilderForTenant.ToArray();
                tenant.GroupSelect = groupSelectBuilderForTenant.ToArray();
            }
        }

        protected async Task<List<DirectoryObject>> QueryAzureADTenantsAsync(OperationContext currentContext, List<AzureTenant> azureTenants)
        {
            // Create a task for each tenant to query
            var tenantQueryTasks = azureTenants.Select(async tenant =>
            {
                Stopwatch timer = new Stopwatch();
                List<DirectoryObject> tenantResults = null;
                try
                {
                    timer.Start();
                    tenantResults = await QueryAzureADTenantAsync(currentContext, tenant, true).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    ClaimsProviderLogging.LogException(ClaimsProviderName, $"in QueryAzureADTenantsAsync while querying tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
                }
                finally
                {
                    timer.Stop();
                }
                if (tenantResults != null)
                {
                    ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Got {tenantResults.Count} users/groups in {timer.ElapsedMilliseconds.ToString()} ms from '{tenant.Name}' with input '{currentContext.Input}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                }
                else
                {
                    ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Got no result from '{tenant.Name}' with input '{currentContext.Input}', search took {timer.ElapsedMilliseconds.ToString()} ms", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Lookup);
                }
                return tenantResults;
            });

            // Wait for all tasks to complete
            List<DirectoryObject> allResults = new List<DirectoryObject>();
            List<DirectoryObject>[] tenantsResults = await Task.WhenAll(tenantQueryTasks).ConfigureAwait(false);
            for (int i = 0; i < tenantsResults.Length; i++)
            {
                allResults.AddRange(tenantsResults[i]);
            }
            return allResults;
        }

        protected virtual async Task<List<DirectoryObject>> QueryAzureADTenantAsync(OperationContext currentContext, AzureTenant tenant, bool firstAttempt)
        {
            List<DirectoryObject> tenantResults = new List<DirectoryObject>();
            if (String.IsNullOrWhiteSpace(tenant.UserFilter) && String.IsNullOrWhiteSpace(tenant.GroupFilter))
            {
                return tenantResults;
            }

            if (tenant.GraphService == null)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Cannot query Azure AD tenant '{tenant.Name}' because it was not initialized", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
                return tenantResults;
            }

            ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Querying Azure AD tenant '{tenant.Name}' for users and groups, with input '{currentContext.Input}'", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Lookup);
            object lockAddResultToCollection = new object();
            int timeout = this.LocalConfiguration.Timeout;
            int maxRetry = currentContext.OperationType == OperationType.Validation ? 3 : 2;

            try
            {
                using (new SPMonitoredScope($"[{ClaimsProviderName}] Querying Azure AD tenant '{tenant.Name}' for users and groups, with input '{currentContext.Input}'", 1000))
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
                    BatchRequestContent batchRequestContent = new BatchRequestContent(tenant.GraphService);


                    // Allow Advanced query as documented in  https://learn.microsoft.com/en-us/graph/sdks/create-requests?tabs=csharp#retrieve-a-list-of-entities
                    // Add ConsistencyLevel header to eventual and $count=true to fix $filter on CompanyName - https://github.com/Yvand/AzureCP/issues/166
                    //// (Only work for non-batched requests)
                    ///
                    string usersRequestId = String.Empty;
                    if (!String.IsNullOrWhiteSpace(tenant.UserFilter))
                    {
                        // https://stackoverflow.com/questions/56417435/when-i-set-an-object-using-an-action-the-object-assigned-is-always-null
                        RequestInformation userRequest = tenant.GraphService.Users.ToGetRequestInformation(conf =>
                        {
                            conf.QueryParameters = new UsersRequestBuilder.UsersRequestBuilderGetQueryParameters
                            {
                                Count = true,
                                //Filter = tenant.UserFilter,
                                Select = tenant.UserSelect,
                                Top = 2,    // YVANDEBUG
                            };
                            conf.Headers = new RequestHeaders
                            {
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
                        GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration groupsRequestConfig = new GroupsRequestBuilder.GroupsRequestBuilderGetRequestConfiguration
                        {
                            QueryParameters = new GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters
                            {
                                Count = true,
                                Filter = tenant.GroupFilter,
                                Select = tenant.GroupSelect,
                            },
                            Headers = new RequestHeaders
                                {
                                    { "ConsistencyLevel", "eventual" }
                                },
                            Options = new List<IRequestOption>
                                {
                                    retryHandlerOption,
                                }
                        };
                        RequestInformation groupRequest = tenant.GraphService.Groups.ToGetRequestInformation(conf =>
                        {
                            conf.QueryParameters = new GroupsRequestBuilder.GroupsRequestBuilderGetQueryParameters
                            {
                                Count = true,
                                Filter = tenant.GroupFilter,
                                Select = tenant.GroupSelect,
                            };
                            conf.Headers = new RequestHeaders
                            {
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

                    BatchResponseContent returnedResponse = await tenant.GraphService.Batch.PostAsync(batchRequestContent).ConfigureAwait(false);
                    UserCollectionResponse userCollectionResult = await returnedResponse.GetResponseByIdAsync<UserCollectionResponse>(usersRequestId).ConfigureAwait(false);
                    GroupCollectionResponse groupCollectionResult = await returnedResponse.GetResponseByIdAsync<GroupCollectionResponse>(groupsRequestId).ConfigureAwait(false);

                    // Process users result
                    ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Query to tenant '{tenant.Name}' returned {(userCollectionResult?.Value == null ? 0 : userCollectionResult.Value.Count)} user(s) with filter \"{tenant.UserFilter}\"", TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Lookup);
                    if (userCollectionResult?.Value != null)
                    {
                        PageIterator<User, UserCollectionResponse> usersPageIterator = PageIterator<User, UserCollectionResponse>.CreatePageIterator(
                            tenant.GraphService,
                            userCollectionResult,
                            (user) =>
                            {
                                lock (lockAddResultToCollection)
                                {
                                    if (tenant.ExcludeMembers == true && !String.Equals(user.UserType, ClaimsProviderConstants.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        tenantResults.Add(user);
                                    }
                                    else if (tenant.ExcludeGuests == true && !String.Equals(user.UserType, ClaimsProviderConstants.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        tenantResults.Add(user);
                                    }
                                    else
                                    {
                                        tenantResults.Add(user);
                                    }
                                }
                                return true; // return true to continue the iteration
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
                                lock (lockAddResultToCollection)
                                {
                                    tenantResults.Add(group);
                                }
                                return true; // return true to continue the iteration
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
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Queries on Azure AD tenant '{tenant.Name}' exceeded timeout of {timeout} ms and were cancelled.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Lookup);
            }
            catch (ServiceException ex)
            {
                ClaimsProviderLogging.LogException(ClaimsProviderName, $"Microsoft.Graph could not query tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
            }
            catch (AggregateException ex)
            {
                // Task.WaitAll throws an AggregateException, which contains all exceptions thrown by tasks it waited on
                ClaimsProviderLogging.LogException(ClaimsProviderName, $"while querying Azure AD tenant '{tenant.Name}'", TraceCategory.Lookup, ex);
            }
            finally
            {
            }
            return tenantResults;
        }
    }
}
