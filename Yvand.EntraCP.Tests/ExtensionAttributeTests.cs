using NUnit.Framework;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    public class ExtensionAttributeTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings(bool applyChanges)
        {
            base.InitializeSettings(false);
            ClaimTypeConfig ctConfigExtensionAttribute = new ClaimTypeConfig
            {
                ClaimType = TestContext.Parameters["MultiPurposeCustomClaimType"],
                ClaimTypeDisplayName = "extattr1",
                EntityProperty = DirectoryObjectProperty.extensionAttribute1,
                EntityType = DirectoryObjectType.User,
                SharePointEntityType = ClaimsProviderConstants.GroupClaimEntityType,
            };
            Settings.ClaimTypes.Add(ctConfigExtensionAttribute);
            if (applyChanges)
            {
                TestSettingsAndApplyThemIfValid();
            }
        }

        [TestCase("val", 1, "value1")]  // Extension attribute configuration
        public void TestSearchExtensionAttributeManual(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }
    }
}
