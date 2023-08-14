using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System.Security.Claims;
using Yvand.ClaimsProviders.Config;

namespace Yvand.ClaimsProviders.Tests
{
    public class CustomConfigTestsBase : EntityTestsBase
    {
        public static string GroupsClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType;

        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();

            // Extra initialization for current test class
            Config.EnableAugmentation = true;
            Config.ClaimTypes.GetByClaimType(Config.SPTrust.IdentityClaimTypeInformation.MappedClaimType).PrefixToBypassLookup = "bypass-user:";
            Config.ClaimTypes.GetByClaimType(UnitTestsHelper.TrustedGroupToAdd_ClaimType).PrefixToBypassLookup = "bypass-group:";
            ClaimTypeConfig ctConfigExtensionAttribute = new ClaimTypeConfig
            {
                ClaimType = TestContext.Parameters["MultiPurposeCustomClaimType"],
                ClaimTypeDisplayName = "extattr1",
                EntityProperty = DirectoryObjectProperty.extensionAttribute1,
                SharePointEntityType = "FormsRole",
            };
            Config.ClaimTypes.Add(ctConfigExtensionAttribute);
            Config.Update();
        }
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class CustomConfigTests : CustomConfigTestsBase
    {
        [TestCase("bypass-user:externalUser@contoso.com", 1, "externalUser@contoso.com")]
        [TestCase("externalUser@contoso.com", 0, "")]
        [TestCase("bypass-user:", 0, "")]
        public void BypassLookupOnIdentityClaimTest(string inputValue, int expectedCount, string expectedClaimValue)
        {
            TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(Config.SPTrust.IdentityClaimTypeInformation.MappedClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, Config.SPTrust.Name));
                TestValidationOperation(inputClaim, true, expectedClaimValue);
            }
        }

        [TestCase(@"bypass-group:domain\groupValue", 1, @"domain\groupValue")]
        [TestCase(@"domain\groupValue", 0, "")]
        [TestCase("bypass-group:", 0, "")]
        [TestCase("val", 1, "value1")]  // Extension attribute configuration
        public override void SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.SearchEntities(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        [Test]
        [NonParallelizable]
        public void BypassServer()
        {
            Config.AlwaysResolveUserInput = true;
            Config.Update();

            try
            {
                TestSearchOperation(UnitTestsHelper.RandomClaimValue, 3, UnitTestsHelper.RandomClaimValue);

                SPClaim inputClaim = new SPClaim(Config.SPTrust.IdentityClaimTypeInformation.MappedClaimType, UnitTestsHelper.RandomClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, Config.SPTrust.Name));
                TestValidationOperation(inputClaim, true, UnitTestsHelper.RandomClaimValue);
            }
            finally
            {
                Config.AlwaysResolveUserInput = false;
                Config.Update();
            }
        }
    }
}
