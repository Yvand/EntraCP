using Microsoft.SharePoint.Administration.Claims;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Diagnostics;
using System.Security.Claims;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class BypassDirectoryTests : CustomConfigTestsBase
    {
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

        [Test]
        [NonParallelizable]
        public void TestBypassDirectoryGlobally()
        {
            Settings.AlwaysResolveUserInput = true;
            GlobalConfiguration.ApplySettings(Settings, true);
            try
            {
                Trace.TraceInformation($"{DateTime.Now:s} [BypassDirectoryTests.TestBypassDirectoryGlobally] Updated configuration: {JsonConvert.SerializeObject(GlobalConfiguration.Settings.ClaimTypes, Formatting.None)}");
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

    [TestFixture]
    public class ExtensionAttributeTests : CustomConfigTestsBase
    {
        [TestCase("val", 1, "value1")]  // Extension attribute configuration
        public void TestSearchManual(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }
    }
}
