using NUnit.Framework;

namespace Yvand.ClaimsProviders.Tests
{
    [TestFixture]
    public class RequireExactMatchOnBaseConfigTests : NewEntityTestsBase
    {
        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();

            // Extra initialization for current test class
            Config.FilterExactMatchOnly = true;
            Config.Update();
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void SearchEntities(SearchEntityData registrationData)
        {
            base.SearchEntities(registrationData);
        }

        [TestCase(@"aadgroup1143", 1, "3f4b724c-125d-47b4-b989-195b29417d6e")]
        public override void SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.SearchEntities(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void RequireExactMatchDuringSearch(ValidateEntityData registrationData)
        {
            int expectedCount = registrationData.ShouldValidate ? 1 : 0;
            TestSearchOperation(registrationData.ClaimValue, expectedCount, registrationData.ClaimValue);
        }
    }

    [TestFixture]
    public class RequireExactMatchOnCustomConfigTests : CustomConfigTestsBase
    {
        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();

            // Extra initialization for current test class
            Config.FilterExactMatchOnly = true;
            Config.Update();
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void RequireExactMatchDuringSearch(ValidateEntityData registrationData)
        {
            int expectedCount = registrationData.ShouldValidate ? 1 : 0;
            TestSearchOperation(registrationData.ClaimValue, expectedCount, registrationData.ClaimValue);
        }
    }
}
