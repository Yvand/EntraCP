using NUnit.Framework;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    public class WrongConfigBadClaimTypeTests : ClaimsProviderTestsBase
    {
        public override bool DoAugmentationTest => false;

        public override void InitializeSettings(bool applyChanges)
        {
            base.InitializeSettings(false);
            ClaimTypeConfig randomClaimTypeConfig = new ClaimTypeConfig
            {
                ClaimType = UnitTestsHelper.RandomClaimType,
                EntityProperty = UnitTestsHelper.RandomObjectProperty,
            };
            Settings.ClaimTypes = new ClaimTypeConfigCollection(UnitTestsHelper.SPTrust) { randomClaimTypeConfig };
            ConfigurationShouldBeValid = false;
            base.TestSettingsAndApplyThemIfValid();
        }

        ///// <summary>
        /////  Disable test augmentation with real data in this test class
        ///// </summary>
        ///// <param name="registrationData"></param>
        //public override void TestAugmentationOperation(ValidateEntityData registrationData)
        //{
        //}
    }
}
