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
            Settings.ClaimTypes.UpdateGroupIdentifier(DirectoryObjectProperty.DisplayName);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

#if DEBUG
        [TestCase("all company", 0, "All Company")]
        [TestCase("aadgroup1", 1, "AADGroup1")]
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
            Settings.ClaimTypes.UpdateGroupIdentifier(DirectoryObjectProperty.DisplayName);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

#if DEBUG
        [TestCase("all company", 1, "All Company")]
        [TestCase("aadgroup1", 1, "AADGroup1")]
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
