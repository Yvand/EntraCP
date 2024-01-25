using Microsoft.SharePoint.Administration.Claims;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Diagnostics;
using System.Security.Claims;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    public class CustomConfigTestsBase : EntityTestsBase
    {
        public static string GroupsClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType;

        public override void InitializeConfiguration(bool applyChanges)
        {
            base.InitializeConfiguration(false);
            Settings.EnableAugmentation = true;
            Settings.ClaimTypes.GetByClaimType(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType).PrefixToBypassLookup = "bypass-user:";
            Settings.ClaimTypes.GetByClaimType(UnitTestsHelper.TrustedGroupToAdd_ClaimType).PrefixToBypassLookup = "bypass-group:";
            ClaimTypeConfig ctConfigExtensionAttribute = new ClaimTypeConfig
            {
                ClaimType = TestContext.Parameters["MultiPurposeCustomClaimType"],
                ClaimTypeDisplayName = "extattr1",
                EntityProperty = DirectoryObjectProperty.extensionAttribute1,
                EntityType = DirectoryObjectType.User,
                SharePointEntityType = "FormsRole",
            };
            Settings.ClaimTypes.Add(ctConfigExtensionAttribute);
            if (applyChanges)
            {
                GlobalConfiguration.ApplySettings(Settings, true);
                Trace.TraceInformation($"{DateTime.Now:s} [CustomConfigTestsBase] Updated configuration: {JsonConvert.SerializeObject(Settings, Formatting.None)}");
            }
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
                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
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
            Settings.AlwaysResolveUserInput = true;
            GlobalConfiguration.ApplySettings(Settings, true);
            try
            {
                Trace.TraceInformation($"{DateTime.Now:s} [CustomConfigTests.BypassServer] Updated configuration: {JsonConvert.SerializeObject(GlobalConfiguration.Settings.ClaimTypes, Formatting.None)}");
                TestSearchOperation(UnitTestsHelper.RandomClaimValue, 3, UnitTestsHelper.RandomClaimValue);

                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, UnitTestsHelper.RandomClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                TestValidationOperation(inputClaim, true, UnitTestsHelper.RandomClaimValue);
            }
            finally
            {
                Settings.AlwaysResolveUserInput = false;
                GlobalConfiguration.ApplySettings(Settings, true);
            }
        }
    }
}
