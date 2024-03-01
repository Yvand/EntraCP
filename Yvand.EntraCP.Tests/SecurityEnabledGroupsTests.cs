using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System.Security.Claims;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class WithSecurityEnabledGroupsOnlyTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.EnableAugmentation = true;
            Settings.FilterSecurityEnabledGroupsOnly = true;
            Settings.ClaimTypes.UpdateGroupIdentifier(DirectoryObjectProperty.Id);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

#if DEBUG
        [TestCase("EntracpTestM365Group2", 0, "6d1efd6c-bf07-4a09-9cb0-9b3d367af415")]
        [TestCase("EntracpTestM365Group1", 1, "1c4a3a59-2c52-44e8-b210-f020fa4526a8")]
        public void TestSearchAndValidation(string inputValue, int expectedCount, string expectedClaimValue)
        {
            TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(base.GroupIdentifierClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                TestValidationOperation(inputClaim, true, expectedClaimValue);
            }
        }
#endif
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class WithoutSecurityEnabledGroupsOnlyTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.EnableAugmentation = true;
            Settings.FilterSecurityEnabledGroupsOnly = false;
            Settings.ClaimTypes.UpdateGroupIdentifier(DirectoryObjectProperty.Id);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

#if DEBUG
        [TestCase("EntracpTestM365Group2", 1, "6d1efd6c-bf07-4a09-9cb0-9b3d367af415")]
        [TestCase("EntracpTestM365Group1", 1, "1c4a3a59-2c52-44e8-b210-f020fa4526a8")]
        public void TestSearchAndValidation(string inputValue, int expectedCount, string expectedClaimValue)
        {
            TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(base.GroupIdentifierClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                TestValidationOperation(inputClaim, true, expectedClaimValue);
            }
        }
#endif
    }
}
