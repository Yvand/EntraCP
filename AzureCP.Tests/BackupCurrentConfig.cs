using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Yvand.ClaimsProviders;
using Yvand.ClaimsProviders.Configuration.AzureAD;

namespace AzureCP.Tests
{
    /// <summary>
    /// This class creates a backup of current configuration and provides one that can be modified as needed. At the end of the test, initial configuration will be restored.
    /// </summary>
    public class BackupCurrentConfig
    {
        protected AzureADEntityProviderConfiguration Config;
        private AzureADEntityProviderConfiguration BackupConfig;

        [OneTimeSetUp]
        public void Init()
        {
            Trace.WriteLine($"{DateTime.Now.ToString("s")} Start backup of current AzureCP configuration");
            Config = AzureCPSE.GetConfiguration();
            if (Config == null)
            {
                Trace.TraceWarning($"{DateTime.Now.ToString("s")} Configuration for AzureCPSE does not exist, create it with default settings...");
                Config = AzureCPSE.CreateConfiguration();
            }
            BackupConfig = Config.CopyConfiguration() as AzureADEntityProviderConfiguration;
            InitializeConfiguration();
        }

        /// <summary>
        /// Initialize configuration
        /// </summary>
        public virtual void InitializeConfiguration()
        {
            Config = AzureCPSE.CreateConfiguration();
            Config.ProxyAddress = TestContext.Parameters["ProxyAddress"];

#if DEBUG
            Config.Timeout = 99999;
#endif

            string json = File.ReadAllText(UnitTestsHelper.AzureTenantsJsonFile);
            List<AzureTenant> azureTenants = JsonConvert.DeserializeObject<List<AzureTenant>>(json);
            Config.AzureTenants = azureTenants;
            Config.Update();
            Trace.WriteLine($"{DateTime.Now.ToString("s")} Set {Config.AzureTenants.Count} Azure AD tenants to AzureCP configuration");
        }

        [OneTimeTearDown]
        public void Cleanup()
        {
            //Config.ApplyConfiguration(BackupConfig);
            //Config.Update();
            AzureCPSE.SaveConfiguration(BackupConfig);
            Trace.WriteLine($"{DateTime.Now.ToString("s")} Restored original settings of AzureCP configuration");
        }
    }
}
