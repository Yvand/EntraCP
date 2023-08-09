using Microsoft.Graph.Models;
using Microsoft.SharePoint.Administration;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Threading.Tasks;
using Yvand.ClaimsProviders.Configuration;
using Yvand.ClaimsProviders.Configuration.AzureAD;

namespace Yvand.ClaimsProviders
{
    public abstract class EntityProviderBase//<TConfiguration>
        //where TConfiguration : IEntityProviderSettings
    {
        /// <summary>
        /// Gets or sets the local configuration, which is a copy of the global configuration stored in a persisted object
        /// </summary>
        //private TConfiguration LocalConfiguration { get; set; }

        ///// <summary>
        ///// Gets or sets the current version of the local configuration
        ///// </summary>
        //private long LocalConfigurationVersion { get; set; } = 0;

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

//        /// <summary>
//        /// Ensure that property LocalConfiguration is valid and up to date
//        /// </summary>
//        /// <param name="configurationName"></param>
//        /// <returns>return true if local configuration is valid and up to date</returns>
//        public TConfiguration RefreshLocalConfigurationIfNeeded(string configurationName)
//        {
//            TConfiguration globalConfiguration = GetGlobalConfiguration(configurationName);

//            if (globalConfiguration == null)
//            {
//                Logger.Log($"[{ClaimsProviderName}] Cannot continue because configuration '{configurationName}' was not found in configuration database, visit AzureCP admin pages in central administration to create it.",
//                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
//                this.LocalConfiguration = default(TConfiguration);
//                return default(TConfiguration);
//            }

//            if (this.LocalConfigurationVersion == globalConfiguration.Version)
//            {
//                Logger.Log($"[{ClaimsProviderName}] Configuration '{configurationName}' is up to date with version {this.LocalConfigurationVersion}.",
//                    TraceSeverity.VerboseEx, EventSeverity.Information, TraceCategory.Core);
//                return this.LocalConfiguration;
//            }

//            Logger.Log($"[{ClaimsProviderName}] Configuration '{globalConfiguration.Name}' has new version {globalConfiguration.Version}, refreshing local copy",
//                TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Core);

//            this.LocalConfiguration = (TConfiguration)globalConfiguration.CopyConfiguration();
//#if !DEBUGx
//            this.LocalConfigurationVersion = globalConfiguration.Version;
//#endif

//            if (this.LocalConfiguration.ClaimTypes == null || this.LocalConfiguration.ClaimTypes.Count == 0)
//            {
//                Logger.Log($"[{ClaimsProviderName}] Configuration '{this.LocalConfiguration.Name}' was found but collection ClaimTypes is empty. Visit AzureCP admin pages in central administration to create it.",
//                    TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
//            }
//            return this.LocalConfiguration;
//        }

        ///// <summary>
        ///// Returns the global configuration, stored as a persisted object in the SharePoint configuration database
        ///// </summary>
        ///// <param name="configurationName">The name of the configuration</param>
        ///// <param name="initializeRuntimeSettings">Set to true to initialize the runtime settings</param>
        ///// <returns></returns>
        //public static TConfiguration GetGlobalConfiguration(string configurationName, bool initializeRuntimeSettings = false)
        //{
        //    //TConfiguration configuration = (TConfiguration) EntityProviderConfiguration.GetGlobalConfiguration(configurationName, T, initializeRuntimeSettings);
        //    TConfiguration configuration = (TConfiguration)EntityProviderConfiguration.GetGlobalConfiguration(configurationName, typeof(TConfiguration), initializeRuntimeSettings);
        //    return configuration;
        //    //SPFarm parent = SPFarm.Local;
        //    //try
        //    //{
        //    //    TConfiguration configuration = (TConfiguration)parent.GetObject(configurationName, parent.Id, typeof(TConfiguration));
        //    //    if (configuration != null && initializeRuntimeSettings == true)
        //    //    {
        //    //        configuration.InitializeRuntimeSettings();
        //    //    }
        //    //    return configuration;
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    Logger.LogException(String.Empty, $"while retrieving configuration '{configurationName}'", TraceCategory.Configuration, ex);
        //    //}
        //    //return null;
        //}

        ///// <summary>
        ///// Deletes the global configuration (persisted object) from the SharePoint configuration database
        ///// </summary>
        ///// <param name="configurationName">Name of persisted object to delete</param>
        //public static void DeleteGlobalConfiguration(string configurationName)
        //{
        //    EntityProviderConfiguration.DeleteGlobalConfiguration(configurationName, typeof(TConfiguration));
        //    //TConfiguration configuration = GetGlobalConfiguration(configurationName);
        //    //if (configuration == null)
        //    //{
        //    //    Logger.Log($"Configuration '{configurationName}' was not found in configuration database", TraceSeverity.Medium, EventSeverity.Error, TraceCategory.Core);
        //    //    return;
        //    //}
        //    //configuration.Delete();
        //    //Logger.Log($"Configuration '{configurationName}' was successfully deleted from configuration database", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        //}

        ///// <summary>
        ///// Create the persisted object that stores the global configuration, in the SharePoint configuration database.
        ///// It will delete the configuration if it already exists.
        ///// </summary>
        ///// <param name="configurationID">GUID of the persisted object</param>
        ///// <param name="configurationName">Name of the persisted object</param>
        ///// <param name="claimsProviderName">Name of the claims provider associated with this configuration</param>
        ///// <returns></returns>
        //public static TConfiguration CreateGlobalConfiguration(string configurationID, string configurationName, string claimsProviderName)
        //{
        //    IEntityProviderSettings config = EntityProviderConfiguration.CreateGlobalConfiguration(configurationID, configurationName, claimsProviderName, typeof(TConfiguration));
        //    return (TConfiguration)config;
        //    //if (String.IsNullOrWhiteSpace(claimsProviderName))
        //    //{
        //    //    throw new ArgumentNullException(nameof(claimsProviderName));
        //    //}

        //    //// Ensure it doesn't already exists and delete it if so
        //    //TConfiguration existingConfig = GetGlobalConfiguration(configurationName);
        //    //if (existingConfig != null)
        //    //{
        //    //    DeleteGlobalConfiguration(configurationName);
        //    //}

        //    //Logger.Log($"Creating configuration '{configurationName}' with Id {configurationID}...", TraceSeverity.VerboseEx, EventSeverity.Error, TraceCategory.Core);

        //    //// Calling constructor as below is not possible and generate Compiler Error CS0304, so use reflection to call the desired constructor instead
        //    ////TConfiguration config = new TConfiguration(persistedObjectName, SPFarm.Local, claimsProviderName);
        //    //ConstructorInfo ctorWithParameters = typeof(TConfiguration).GetConstructor(new[] { typeof(string), typeof(SPFarm), typeof(string) });
        //    //TConfiguration config = (TConfiguration)ctorWithParameters.Invoke(new object[] { configurationName, SPFarm.Local, claimsProviderName });

        //    //config.Id = new Guid(configurationID);
        //    //// If parameter ensure is true, the call will not throw if the object already exists.
        //    //config.Update(true);
        //    //Logger.Log($"Created configuration '{configurationName}' with Id {config.Id}", TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
        //    //return config;
        //}
    }
}
