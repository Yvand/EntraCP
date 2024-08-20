using NUnit.Framework;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class ExtensionAttributeTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            ClaimTypeConfig ctConfigExtensionAttribute = new ClaimTypeConfig
            {
                ClaimType = TestContext.Parameters["MultiPurposeCustomClaimType"],
                ClaimTypeDisplayName = "extrattr1",
                EntityProperty = DirectoryObjectProperty.extensionAttribute1,
                EntityType = DirectoryObjectType.User,
                SharePointEntityType = ClaimsProviderConstants.GroupClaimEntityType,
            };
            Settings.ClaimTypes.Add(ctConfigExtensionAttribute);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

#if DEBUG
        [TestCase("val", 1, "value1")]  // Extension attribute configuration
        public void TestSearchExtensionAttributeManual(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }
#endif
    }
}
