using NUnit.Framework;

namespace Yvand.ClaimsProviders.Tests
{
    [TestFixture]
    public class ExcludeAllUserAccountsTests : NewEntityTestsBase
    {
        public override bool ExcludeGuestUsers => true;
        public override bool ExcludeMemberUsers => true;

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void AugmentEntity(ValidateEntityData registrationData)
        {
            base.AugmentEntity(registrationData);
        }

        [TestCase(@"AADGroup1130", 1, "e86ace87-37ba-4ee1-8087-ecd783728233")]
        [TestCase(@"aaduser1", 0, "")]
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
    }
}
