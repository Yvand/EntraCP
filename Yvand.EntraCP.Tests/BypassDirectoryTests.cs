using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System.Security.Claims;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class BypassDirectoryOnClaimTypesTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings(bool applyChanges)
        {
            base.InitializeSettings(false);
            Settings.EnableAugmentation = true;
            Settings.ClaimTypes.GetMainConfigurationForDirectoryObjectType(DirectoryObjectType.User).PrefixToBypassLookup = "bypass-user:";
            Settings.ClaimTypes.GetMainConfigurationForDirectoryObjectType(DirectoryObjectType.Group).PrefixToBypassLookup = "bypass-group:";
            if (applyChanges)
            {
                TestSettingsAndApplyThemIfValid();
            }
        }

        [TestCase("bypass-user:externalUser@contoso.com", 1, "externalUser@contoso.com")]
        [TestCase("externalUser@contoso.com", 0, "")]
        [TestCase("bypass-user:", 0, "")]
        [TestCase(@"bypass-group:domain\groupValue", 1, @"domain\groupValue")]
        [TestCase(@"domain\groupValue", 0, "")]
        [TestCase("bypass-group:", 0, "")]
        public void TestBypassDirectoryByClaimType(string inputValue, int expectedCount, string expectedClaimValue)
        {
            TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                TestValidationOperation(inputClaim, true, expectedClaimValue);
            }
        }
    }

    [TestFixture]
    public class BypassDirectoryGloballyTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings(bool applyChanges)
        {
            base.InitializeSettings(false);
            Settings.AlwaysResolveUserInput = true;
            if (applyChanges)
            {
                TestSettingsAndApplyThemIfValid();
            }
        }

        [Test]
        public void TestBypassDirectoryGlobally()
        {
            TestSearchOperation(UnitTestsHelper.RandomClaimValue, 2, UnitTestsHelper.RandomClaimValue);
            TestValidationOperation(base.UserIdentifierClaimType, UnitTestsHelper.RandomClaimValue, true);
            TestValidationOperation(base.GroupIdentifierClaimType, UnitTestsHelper.RandomClaimValue, true);
        }
    }
}
