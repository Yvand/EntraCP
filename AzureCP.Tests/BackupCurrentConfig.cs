using azurecp;
using NUnit.Framework;
using System;
using System.Diagnostics;

namespace AzureCP.Tests
{
    /// <summary>
    /// This class creates a backup of current configuration and provides one that can be modified as needed. At the end of the test, initial configuration will be restored.
    /// </summary>
    public class BackupCurrentConfig
    {
        protected AzureCPConfig Config;
        private AzureCPConfig BackupConfig;

        [OneTimeSetUp]
        public void Init()
        {
            Config = AzureCPConfig.GetConfiguration(UnitTestsHelper.ClaimsProviderConfigName, UnitTestsHelper.SPTrust.Name);
            if (Config == null)
            {
                Trace.TraceWarning($"{DateTime.Now.ToString("s")} Configuration {UnitTestsHelper.ClaimsProviderConfigName} does not exist, create it with default settings...");
                Config = AzureCPConfig.CreateConfiguration(ClaimsProviderConstants.CONFIG_ID, ClaimsProviderConstants.CONFIG_NAME, UnitTestsHelper.SPTrust.Name);
            }
            BackupConfig = Config.CopyConfiguration();
            InitializeConfiguration();
        }

        /// <summary>
        /// Initialize configuration
        /// </summary>
        public virtual void InitializeConfiguration()
        {
            UnitTestsHelper.InitializeConfiguration(Config);
        }

        [OneTimeTearDown]
        public void Cleanup()
        {
            Config.ApplyConfiguration(BackupConfig);
            Config.Update();
        }
    }
}
