using azurecp;
using NUnit.Framework;
using System;

namespace AzureCP.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class AugmentationTests
    {
        private AzureCPConfig Config;
        private AzureCPConfig BackupConfig;

        [OneTimeSetUp]
        public void Init()
        {
            Console.WriteLine($"Starting augmentation test {TestContext.CurrentContext.Test.Name}...");
            Config = AzureCPConfig.GetConfiguration(UnitTestsHelper.ClaimsProviderConfigName, UnitTestsHelper.SPTrust.Name);
            BackupConfig = Config.CopyPersistedProperties();
            Config.EnableAugmentation = true;
            Config.Update();
        }

        [OneTimeTearDown]
        public void Cleanup()
        {
            Config.ApplyConfiguration(BackupConfig);
            Config.Update();
            Console.WriteLine($"Restored actual configuration.");
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData")]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void AugmentEntity(ValidateEntityData registrationData)
        {
            UnitTestsHelper.TestAugmentationOperation(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, registrationData.IsMemberOfTrustedGroup);
        }

        [TestCase("i:05.t|contoso.local|xyd@FAKE.onmicrosoft.com", false)]
        [TestCase("i:05.t|contoso.local|aadUser1@yvandev.onmicrosoft.com", true)]
        [TestCase("i:05.t|contoso.local|yvand@outlook.com", true)]
        public void DEBUG_AugmentEntity(string claimValue, bool shouldHavePermissions)
        {
            UnitTestsHelper.TestAugmentationOperation(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimValue, shouldHavePermissions);
        }
    }
}
