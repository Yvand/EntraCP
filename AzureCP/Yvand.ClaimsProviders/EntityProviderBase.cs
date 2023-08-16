﻿using Microsoft.Graph.Models;
using System.Collections.Generic;
using System.Threading.Tasks;
using Yvand.ClaimsProviders.Config;

namespace Yvand.ClaimsProviders
{
    public abstract class EntityProviderBase
    {
        /// <summary>
        /// Gets the name of the claims provider using this class
        /// </summary>
        public string ClaimsProviderName { get; }

        /// <summary>
        /// Returns a list of users and groups
        /// </summary>
        /// <param name="currentContext"></param>
        /// <returns></returns>
        public abstract Task<List<DirectoryObject>> SearchOrValidateEntitiesAsync(OperationContext currentContext);

        /// <summary>
        /// Returns the groups the user is member of
        /// </summary>
        /// <param name="currentContext"></param>
        /// <param name="groupClaimTypeConfig"></param>
        /// <returns></returns>
        public abstract Task<List<string>> GetEntityGroupsAsync(OperationContext currentContext, DirectoryObjectProperty groupClaimTypeConfig);

        public EntityProviderBase(string claimsProviderName)
        {
            this.ClaimsProviderName = claimsProviderName;
        }
    }
}
