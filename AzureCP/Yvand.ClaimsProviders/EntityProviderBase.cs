﻿using Microsoft.Graph.Models;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using Yvand.ClaimsProviders.Configuration;
using static Yvand.ClaimsProviders.ClaimsProviderLogging;

namespace Yvand.ClaimsProviders
{
    public abstract class EntityProviderBase<TConfiguration>
    where TConfiguration : EntityProviderConfiguration
    {
        public TConfiguration LocalConfiguration { get; private set; }
        public long LocalConfigurationVersion = 0;
        public string ClaimsProviderName { get; set; }
        public abstract Task<List<DirectoryObject>> SearchOrValidateEntitiesAsync(OperationContext currentContext);
        /// <summary>
        /// Returns the groups the user is member of
        /// </summary>
        /// <param name="currentContext"></param>
        /// <param name="groupClaimTypeConfig"></param>
        /// <returns></returns>
        public abstract Task<List<string>> GetEntityGroupsAsync(OperationContext currentContext, AzureADObjectProperty groupClaimTypeConfig);

        public EntityProviderBase(string claimsProviderName)
        {
            this.ClaimsProviderName = claimsProviderName;
        }

        /// <summary>
        /// Ensure that property LocalConfiguration is valid and up to date
        /// </summary>
        /// <param name="configurationName"></param>
        /// <returns>return true if local configuration is valid and up to date</returns>
        public bool RefreshLocalConfigurationIfNeeded(string configurationName)
        {
            bool configIsVald = true;
            // Use reflection to call method GetConfiguration(string) of the generic type because TConfiguration.GetConfiguration(persistedObjectName) return Compiler Error CS0704
            //TConfiguration globalConfiguration = TConfiguration.GetConfiguration(persistedObjectName);
            //TConfiguration globalConfiguration = (TConfiguration)typeof(TConfiguration).GetMethod("GetConfiguration", new[] { typeof(string) }).Invoke(null, new object[] { persistedObjectName });
            TConfiguration globalConfiguration = GetGlobalConfiguration(configurationName);

            if (globalConfiguration == null)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Cannot continue because configuration '{configurationName}' was not found in configuration database, visit AzureCP admin pages in central administration to create it.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                this.LocalConfiguration = null;
                return false;
            }

            if (this.LocalConfigurationVersion == globalConfiguration.Version)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Configuration '{configurationName}' is up to date with version {this.LocalConfigurationVersion}.",
                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Core);
                return true;
            }

            ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Configuration '{globalConfiguration.Name}' has new version {globalConfiguration.Version}, refreshing local copy",
                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);

            this.LocalConfiguration = (TConfiguration)globalConfiguration.CopyConfiguration();
#if !DEBUGx
            this.LocalConfigurationVersion = ((SPPersistedObject)globalConfiguration).Version;
#endif

            if (this.LocalConfiguration.ClaimTypes == null || this.LocalConfiguration.ClaimTypes.Count == 0)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] Configuration '{this.LocalConfiguration.Name}' was found but collection ClaimTypes is empty. Visit AzureCP admin pages in central administration to create it.",
                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
                configIsVald = false;
            }
            return configIsVald;
        }

        /// <summary>
        /// Returns the configuration of AzureCP, but does not initialize the runtime settings
        /// </summary>
        /// <param name="configurationName">Name of the configuration</param>
        /// <returns></returns>
        public static TConfiguration GetGlobalConfiguration(string configurationName, bool initializeRuntimeSettings = false)
        {
            SPFarm parent = SPFarm.Local;
            try
            {
                TConfiguration configuration = (TConfiguration)parent.GetObject(configurationName, parent.Id, typeof(TConfiguration));
                if (configuration != null && initializeRuntimeSettings == true)
                {
                    configuration.InitializeRuntimeSettings();
                }
                return configuration;
            }
            catch (Exception ex)
            {
                ClaimsProviderLogging.LogException(String.Empty, $"while retrieving configuration '{configurationName}'", TraceCategory.Configuration, ex);
            }
            return null;
        }

        /// <summary>
        /// Delete persisted object from configuration database
        /// </summary>
        /// <param name="configurationName">Name of persisted object to delete</param>
        public static void DeleteGlobalConfiguration(string configurationName)
        {
            TConfiguration configuration = GetGlobalConfiguration(configurationName);
            if (configuration == null)
            {
                ClaimsProviderLogging.Log($"Configuration '{configurationName}' was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
                return;
            }
            configuration.Delete();
            ClaimsProviderLogging.Log($"Configuration '{configurationName}' was successfully deleted from configuration database", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        }

        /// <summary>
        /// Create a persisted object with default configuration of AzureCP.
        /// </summary>
        /// <param name="configurationID">GUID of the configuration, stored as a persisted object into SharePoint configuration database</param>
        /// <param name="configurationName">Name of the configuration, stored as a persisted object into SharePoint configuration database</param>
        /// <param name="claimsProviderName">Name of the SPTrustedLoginProvider that claims provider is associated with</param>
        /// <returns></returns>
        public static TConfiguration CreateGlobalConfiguration(string configurationID, string configurationName, string claimsProviderName)
        {
            if (String.IsNullOrWhiteSpace(claimsProviderName))
            {
                throw new ArgumentNullException("claimsProviderName");
            }

            // Ensure it doesn't already exists and delete it if so
            TConfiguration existingConfig = GetGlobalConfiguration(configurationName);
            if (existingConfig != null)
            {
                DeleteGlobalConfiguration(configurationName);
            }

            ClaimsProviderLogging.Log($"Creating configuration '{configurationName}' with Id {configurationID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);

            // Calling constructor as below is not possible and generate Compiler Error CS0304, so use reflection to call the desired constructor instead
            //TConfiguration config = new TConfiguration(persistedObjectName, SPFarm.Local, claimsProviderName);
            ConstructorInfo ctorWithParameters = typeof(TConfiguration).GetConstructor(new[] { typeof(string), typeof(SPFarm), typeof(string) });
            TConfiguration config = (TConfiguration)ctorWithParameters.Invoke(new object[] { configurationName, SPFarm.Local, claimsProviderName });

            config.Id = new Guid(configurationID);
            // If parameter ensure is true, the call will not throw if the object already exists.
            config.Update(true);
            ClaimsProviderLogging.Log($"Created configuration '{configurationName}' with Id {config.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
            return config;
        }

        public static void SaveGlobalConfiguration(TConfiguration globalConfiguration)
        {
            // If parameter ensure is true, the call will not throw if the object already exists.
            globalConfiguration.Update(true);
        }
    }
}
