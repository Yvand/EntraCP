using NUnit.Framework;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    public class WrongConfigBadClaimTypeTests : ClaimsProviderTestsBase
    {
        public override bool DoAugmentationTest => false;

        public override void InitializeSettings()
        {
            base.InitializeSettings();
            ClaimTypeConfig randomClaimTypeConfig = new ClaimTypeConfig
            {
                ClaimType = UnitTestsHelper.RandomClaimType,
                EntityProperty = UnitTestsHelper.RandomObjectProperty,
            };
            Settings.ClaimTypes = new ClaimTypeConfigCollection(UnitTestsHelper.SPTrust) { randomClaimTypeConfig };
            ConfigurationShouldBeValid = false;
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }
    }
}
