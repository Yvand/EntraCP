using NUnit.Framework;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    internal class BasicTests : EntityTestsBase
    {
        [Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void SearchEntities(SearchEntityData registrationData)
        {
            base.SearchEntities(registrationData);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        [MaxTime(UnitTestsHelper.MaxTime)]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void ValidateClaim(ValidateEntityData registrationData)
        {
            base.ValidateClaim(registrationData);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void AugmentEntity(ValidateEntityData registrationData)
        {
            base.AugmentEntity(registrationData);
        }

#if DEBUG
        ////[TestCaseSource(typeof(SearchEntityDataSourceCollection))]
        //public void DEBUG_SearchEntitiesFromCollection(string inputValue, string expectedCount, string expectedClaimValue)
        //{
        //    if (!TestSearch) { return; }

        //    TestSearchOperation(inputValue, Convert.ToInt32(expectedCount), expectedClaimValue);
        //}

        [TestCase(@"AADGroup1130", 1, "e86ace87-37ba-4ee1-8087-ecd783728233")]
        [TestCase(@"xyzguest", 0, "xyzGUEST@contoso.com")]
        [TestCase(@"AzureGr}", 1, "ef7d18e6-5c4d-451a-9663-a976be81c91e")]
        [TestCase(@"aad", 30, "")]
        [TestCase(@"AADGroup", 30, "")]
        public override void SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.SearchEntities(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        [TestCase("http://schemas.microsoft.com/ws/2008/06/identity/claims/role", "ef7d18e6-5c4d-451a-9663-a976be81c91e", true)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn", "FakeGuest@contoso.com", false)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress", "FakeGuest.com#EXT#@XXX.onmicrosoft.com", false)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn", "FakeGuest.com#EXT#@XXX.onmicrosoft.com", false)]
        public override void ValidateClaim(string claimType, string claimValue, bool shouldValidate)
        {
            base.ValidateClaim(claimType, claimValue, shouldValidate);
        }

        [TestCase("xydGUEST@FAKE.onmicrosoft.com", false)]
        public override void AugmentEntity(string claimValue, bool shouldHavePermissions)
        {
            base.AugmentEntity(claimValue, shouldHavePermissions);
        }
#endif
    }
}
