using azurecp;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
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

        [TestCase("ext:externalUser@contoso.com", 1, "externalUser@contoso.com")]
        [TestCase("ext:", 0, "")]
        public void TestPrefixToBypassLookup(string inputValue, int expectedCount, string expectedClaimValue)
        {
            ClaimTypeConfig ctConfig = Config.ClaimTypes.FirstOrDefault(x => String.Equals(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            ctConfig.PrefixToBypassLookup = "ext:";
            Config.Update();

            UnitTestsHelper.TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                UnitTestsHelper.TestValidationOperation(inputClaim, true, expectedClaimValue);
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
