using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System.Security.Claims;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    internal class BasicConfigurationTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void TestSearchEntities(SearchEntityData registrationData)
        {
            base.ProcessAndTestSearchEntityData(registrationData);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        [MaxTime(UnitTestsHelper.MaxTime)]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void TestValidateClaim(ValidateEntityData registrationData)
        {
            base.ProcessAndTestValidateEntityData(registrationData);
        }

        /// <summary>
        /// Tests if the augmentation works as expected. By default this test is executed in every scenario.
        /// </summary>
        /// <param name="registrationData"></param>
        [Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void TestAugmentationOperation(ValidateEntityData registrationData)
        {
            base.TestAugmentationOperation(registrationData.ClaimValue, registrationData.IsMemberOfTrustedGroup);
        }

#if DEBUG
        ////[TestCaseSource(typeof(SearchEntityDataSourceCollection))]
        //public void DEBUG_SearchEntitiesFromCollection(string inputValue, string expectedCount, string expectedClaimValue)
        //{
        //    if (!TestSearchEntities) { return; }

        //    TestSearchOperation(inputValue, Convert.ToInt32(expectedCount), expectedClaimValue);
        //}

        [TestCase(@"AADGroup1130", 1, "e86ace87-37ba-4ee1-8087-ecd783728233")]
        [TestCase(@"xyzguest", 0, "xyzGUEST@contoso.com")]
        [TestCase(@"AzureGr}", 1, "ef7d18e6-5c4d-451a-9663-a976be81c91e")]
        [TestCase(@"aad", 30, "")]
        [TestCase(@"AADGroup", 30, "")]
        public void TestSearchManual(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        [TestCase("http://schemas.microsoft.com/ws/2008/06/identity/claims/role", "ef7d18e6-5c4d-451a-9663-a976be81c91e", true)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn", "FakeGuest@contoso.com", false)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress", "FakeGuest.com#EXT#@XXX.onmicrosoft.com", false)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn", "FakeGuest.com#EXT#@XXX.onmicrosoft.com", false)]
        public void TestValidateClaim(string claimType, string claimValue, bool shouldValidate)
        {
            base.TestValidationOperation(claimType, claimValue, shouldValidate);
        }

        //[TestCase("AdeleV@XXXX.OnMicrosoft.com", 1, "AdeleV@XXXX.OnMicrosoft.com")]
        //[TestCase("Adele", 1, "AdeleV@XXXX.OnMicrosoft.com")]
        //public void TestSearchAndValidation(string inputValue, int expectedCount, string expectedClaimValue)
        //{
        //    TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

        //    if (expectedCount > 0)
        //    {
        //        SPClaim inputClaim = new SPClaim(base.UserIdentifierClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
        //        TestValidationOperation(inputClaim, true, expectedClaimValue);
        //    }
        //}

        //[TestCase("xydGUEST@FAKE.onmicrosoft.com", false)]
        //public override void AugmentEntity(string claimValue, bool shouldHavePermissions)
        //{
        //    base.AugmentEntity(claimValue, shouldHavePermissions);
        //}
#endif
    }
}
