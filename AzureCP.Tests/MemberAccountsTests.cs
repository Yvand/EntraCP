using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System;
using System.Security.Claims;

namespace AzureCP.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class MemberAccountsTests : ModifyConfigBase
    {
        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();
            Config.EnableAugmentation = true;
            Config.Update();
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), "GetTestData", new object[] { UnitTestsHelper.DataFile_MemberAccounts_Search })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void SearchEntities(SearchEntityData registrationData)
        {
            UnitTestsHelper.TestSearchOperation(registrationData.Input, registrationData.ExpectedResultCount, registrationData.ExpectedEntityClaimValue);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { UnitTestsHelper.DataFile_MemberAccounts_Validate })]
        [MaxTime(UnitTestsHelper.MaxTime)]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void ValidateClaim(ValidateEntityData registrationData)
        {
            SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            UnitTestsHelper.TestValidationOperation(inputClaim, registrationData.ShouldValidate, registrationData.ClaimValue);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { UnitTestsHelper.DataFile_MemberAccounts_Validate })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void AugmentEntity(ValidateEntityData registrationData)
        {
            UnitTestsHelper.TestAugmentationOperation(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, registrationData.IsMemberOfTrustedGroup);
        }

        //[TestCaseSource(typeof(SearchEntityDataSourceCollection))]
        public void DEBUG_SearchEntitiesFromCollection(string inputValue, string expectedCount, string expectedClaimValue)
        {
            UnitTestsHelper.TestSearchOperation(inputValue, Convert.ToInt32(expectedCount), expectedClaimValue);
        }

        [TestCase(@"AADGroup1", 1, "5b0f6c56-c87f-44c3-9354-56cba03da433")]
        public void DEBUG_SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            UnitTestsHelper.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        //[TestCase("http://schemas.microsoft.com/ws/2008/06/identity/claims/role", "5b0f6c56-c87f-44c3-9354-56cba03da433", true)]
        public void DEBUG_ValidateClaim(string claimType, string claimValue, bool shouldValidate)
        {
            SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            UnitTestsHelper.TestValidationOperation(inputClaim, shouldValidate, claimValue);
        }
    }
}
