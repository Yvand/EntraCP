using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yvand.ClaimsProviders.Configuration;
using static Yvand.ClaimsProviders.ClaimsProviderLogging;
using Yvand.ClaimsProviders.Configuration.AzureAD;
using System.Threading;
using System.Reflection;
using Microsoft.IdentityModel.Protocols;
using Microsoft.Graph.Models;

namespace Yvand.ClaimsProviders
{
    public abstract class EntityProviderBase<TConfiguration>
    where TConfiguration : EntityProviderConfiguration // constrain the generic type to be a IEntityProviderConfiguration
    {
        public TConfiguration CurrentConfiguration { get; set; }
        public long CurrentConfigurationVersion = 0;
        public string ProviderInternalName { get; set; }
        protected ReaderWriterLockSlim Lock_Config;

        public abstract Task<List<DirectoryObject>> SearchOrValidateUsersAsync(OperationContext currentContext);
        public abstract Task<List<Group>> GetEntityGroupsAsync(OperationContext currentContext);

        public EntityProviderBase(string providerInternalName, ref ReaderWriterLockSlim Lock_Config)
        {
            this.ProviderInternalName = providerInternalName;
            this.Lock_Config = Lock_Config;
        }
        public bool ValidateLocalConfiguration(Uri context, string[] entityTypes, string persistedObjectName)
        {
            bool configIsVald = true;
            this.UpdateLocalCopyOfGlobalConfigurationIfNeeded(context, entityTypes, persistedObjectName);
            if (this.CurrentConfiguration == null)
            {
                configIsVald = false;
            }
            if (this.CurrentConfiguration.ClaimTypes == null || this.CurrentConfiguration.ClaimTypes.Count == 0)
            {
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Configuration '{persistedObjectName}' was found but collection ClaimTypes is null or empty. Visit AzureCP admin pages in central administration to create it.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                configIsVald = false;
            }
            return configIsVald;
        }

        public TConfiguration UpdateLocalCopyOfGlobalConfigurationIfNeeded(Uri context, string[] entityTypes, string persistedObjectName)
        {
            // Use reflection to call method GetConfiguration(string) of the generic type because TConfiguration.GetConfiguration(persistedObjectName) return Compiler Error CS0704
            //TConfiguration globalConfiguration = TConfiguration.GetConfiguration(persistedObjectName);
            TConfiguration globalConfiguration = (TConfiguration)typeof(TConfiguration).GetMethod("GetConfiguration", new[] { typeof(string) }).Invoke(null, new object[] { persistedObjectName });

            if (globalConfiguration == null)
            {
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Cannot continue because configuration '{persistedObjectName}' was not found in configuration database, visit AzureCP admin pages in central administration to create it.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                this.CurrentConfiguration = null;
                return null;
            }

            if (this.CurrentConfigurationVersion == ((SPPersistedObject)globalConfiguration).Version)
            {
                ClaimsProviderLogging.Log($"[{ProviderInternalName}] Configuration '{persistedObjectName}' was found, version {((SPPersistedObject)globalConfiguration).Version.ToString()}",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Core);

                return this.CurrentConfiguration;
            }

            ClaimsProviderLogging.Log($"[{ProviderInternalName}] Configuration '{persistedObjectName}' was found with new version {globalConfiguration.Version.ToString()}, refreshing local copy",
                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);

            // Configuration needs to be refreshed, lock current thread in write mode
            this.Lock_Config.EnterWriteLock();
            try
            {
                this.CurrentConfiguration = (TConfiguration)globalConfiguration.CopyConfiguration();
                this.CurrentConfigurationVersion = ((SPPersistedObject)globalConfiguration).Version;
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(ProviderInternalName, "while refreshing configuration", TraceCategory.Core, ex);
            }
            finally
            {
                this.Lock_Config.ExitWriteLock();
            }
            return this.CurrentConfiguration;
        }
    }
}
