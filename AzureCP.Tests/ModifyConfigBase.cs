using azurecp;
using NUnit.Framework;
using System;
using System.Diagnostics;

namespace AzureCP.Tests
{
    /// <summary>
    /// This class creates a backup of current configuration and provides one that can be modified as needed. At the end of the test, initial configuration will be restored.
    /// </summary>
    public class ModifyConfigBase
    {
        protected AzureCPConfig Config;
        private AzureCPConfig BackupConfig;

        [OneTimeSetUp]
        public void Init()
        {
            Config = AzureCPConfig.GetConfiguration(UnitTestsHelper.ClaimsProviderConfigName, UnitTestsHelper.SPTrust.Name);
            BackupConfig = Config.CopyPersistedProperties();
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
