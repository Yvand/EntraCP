using azurecp;
using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System;
using System.Linq;
using System.Security.Claims;

namespace AzureCP.Tests
{
    [TestFixture]
    public class CustomConfigTests : BackupCurrentConfig
    {
        public static string GroupsClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType;

        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();

            // Extra initialization for current test class
            Config.EnableAugmentation = true;
            Config.ClaimTypes.GetByClaimType(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType).PrefixToBypassLookup = "bypass-user:";
            Config.ClaimTypes.GetByClaimType(UnitTestsHelper.TrustedGroupToAdd_ClaimType).PrefixToBypassLookup = "bypass-group:";
            Config.Update();
        }

        [TestCase("bypass-user:externalUser@contoso.com", 1, "externalUser@contoso.com")]
        [TestCase("externalUser@contoso.com", 0, "")]
        [TestCase("bypass-user:", 0, "")]
        public void BypassLookupOnIdentityClaimTest(string inputValue, int expectedCount, string expectedClaimValue)
        {
            UnitTestsHelper.TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                UnitTestsHelper.TestValidationOperation(inputClaim, true, expectedClaimValue);
            }
        }

        [TestCase(@"bypass-group:domain\groupValue", 1, @"domain\groupValue")]
        [TestCase(@"domain\groupValue", 0, "")]
        [TestCase("bypass-group:", 0, "")]
        public void BypassLookupOnGroupClaimTest(string inputValue, int expectedCount, string expectedClaimValue)
        {
            UnitTestsHelper.TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(UnitTestsHelper.TrustedGroupToAdd_ClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
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

        //[Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        ////[Repeat(UnitTestsHelper.TestRepeatCount)]
        //public void RequireExactMatchDuringSearch(ValidateEntityData registrationData)
        //{
        //    Config.FilterExactMatchOnly = true;
        //    Config.Update();

        //    try
        //    {
        //        int expectedCount = registrationData.ShouldValidate ? 1 : 0;
        //        UnitTestsHelper.TestSearchOperation(registrationData.ClaimValue, expectedCount, registrationData.ClaimValue);
        //    }
        //    finally
        //    {
        //        Config.FilterExactMatchOnly = false;
        //        Config.Update();
        //    }
        //}
    }
}
