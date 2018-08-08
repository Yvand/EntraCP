using azurecp;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using NUnit.Framework;
using System;
using System.Linq;
using System.Security.Claims;

namespace AzureCP.Tests
{
    [TestFixture]
    public class CustomConfigTests
    {
        public static string GroupsClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType;
        private AzureCPConfig Config;
        private AzureCPConfig BackupConfig;

        [OneTimeSetUp]
        public void Init()
        {
            Console.WriteLine($"Starting custom config test {TestContext.CurrentContext.Test.Name}...");
            Config = AzureCPConfig.GetConfiguration(UnitTestsHelper.ClaimsProviderConfigName);
            BackupConfig = Config.CopyPersistedProperties();
            Config.ResetClaimTypesList();
        }

        [OneTimeTearDown]
        public void Cleanup()
        {
            Config.ApplyConfiguration(BackupConfig);
            Config.Update();
            Console.WriteLine($"Restored actual configuration.");
        }

        [TestCase("ext:externalUser@contoso.com", 1, "externalUser@contoso.com")]
        [TestCase("ext:", 0, "")]
        public void TestPrefixToBypassLookup(string inputValue, int expectedCount, string expectedClaimValue)
        {
            ClaimTypeConfig ctConfig = Config.ClaimTypes.FirstOrDefault(x => String.Equals(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            ctConfig.PrefixToBypassLookup = "ext:";
            Config.Update();

            UnitTestsHelper.TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                UnitTestsHelper.TestValidationOperation(inputClaim, true, expectedClaimValue);
            }
        }

        [Test]
        public void BypassServer()
        {
            Config.AlwaysResolveUserInput = true;
            Config.Update();

            try
            {
                UnitTestsHelper.TestSearchOperation(UnitTestsHelper.RandomClaimValue, 2, UnitTestsHelper.RandomClaimValue);

                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, UnitTestsHelper.RandomClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                UnitTestsHelper.TestValidationOperation(inputClaim, true, UnitTestsHelper.RandomClaimValue);
            }
            finally
            {
                Config.AlwaysResolveUserInput = false;
                Config.Update();
            }
        }
    }
}
