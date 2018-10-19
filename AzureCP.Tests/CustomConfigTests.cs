using azurecp;
using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System;
using System.Linq;
using System.Security.Claims;

namespace AzureCP.Tests
{
    [TestFixture]
    public class CustomConfigTests : ModifyConfigBase
    {
        public static string GroupsClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType;

        [TestCase("bypass-user:externalUser@contoso.com", 1, "externalUser@contoso.com")]
        [TestCase("externalUser@contoso.com", 0, "")]
        [TestCase("bypass-user:", 0, "")]
        public void BypassLookupOnIdentityClaimTest(string inputValue, int expectedCount, string expectedClaimValue)
        {
            ClaimTypeConfig ctConfig = Config.ClaimTypes.FirstOrDefault(x => String.Equals(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            ctConfig.PrefixToBypassLookup = "bypass-user:";
            Config.Update();

            try
            {
                UnitTestsHelper.TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

                if (expectedCount > 0)
                {
                    SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                    UnitTestsHelper.TestValidationOperation(inputClaim, true, expectedClaimValue);
                }
            }
            finally
            {
                ctConfig.PrefixToBypassLookup = String.Empty;
                Config.Update();
            }
        }

        [TestCase(@"bypass-group:domain\groupValue", 1, @"domain\groupValue")]
        [TestCase(@"domain\groupValue", 0, "")]
        [TestCase("bypass-group:", 0, "")]
        public void BypassLookupOnGroupClaimTest(string inputValue, int expectedCount, string expectedClaimValue)
        {
            ClaimTypeConfig ctConfig = Config.ClaimTypes.FirstOrDefault(x => String.Equals(UnitTestsHelper.TrustedGroupToAdd_ClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            ctConfig.PrefixToBypassLookup = "bypass-group:";
            Config.Update();

            try
            {
                UnitTestsHelper.TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

                if (expectedCount > 0)
                {
                    SPClaim inputClaim = new SPClaim(UnitTestsHelper.TrustedGroupToAdd_ClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                    UnitTestsHelper.TestValidationOperation(inputClaim, true, expectedClaimValue);
                }
            }
            finally
            {
                ctConfig.PrefixToBypassLookup = String.Empty;
                Config.Update();
            }
        }

        [Test]
        public void BypassServer()
        {
            Config.AlwaysResolveUserInput = true;
            Config.Update();

            try
            {
                UnitTestsHelper.TestSearchOperation(UnitTestsHelper.RandomClaimValue, 2, UnitTestsHelper.RandomClaimValue);

                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, UnitTestsHelper.RandomClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                UnitTestsHelper.TestValidationOperation(inputClaim, true, UnitTestsHelper.RandomClaimValue);
            }
            finally
            {
                Config.AlwaysResolveUserInput = false;
                Config.Update();
            }
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { UnitTestsHelper.DataFile_MemberAccounts_Validate })]
        //[Repeat(UnitTestsHelper.TestRepeatCount)]
        public void RequireExactMatchDuringSearch(ValidateEntityData registrationData)
        {
            Config.FilterExactMatchOnly = true;
            Config.Update();

            try
            {
                int expectedCount = registrationData.ShouldValidate ? 1 : 0;
                UnitTestsHelper.TestSearchOperation(registrationData.ClaimValue, expectedCount, registrationData.ClaimValue);
            }
            finally
            {
                Config.FilterExactMatchOnly = false;
                Config.Update();
            }
        }
    }
}
