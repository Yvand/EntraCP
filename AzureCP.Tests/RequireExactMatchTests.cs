using NUnit.Framework;

namespace Yvand.ClaimsProviders.Tests
{
    [TestFixture]
    public class RequireExactMatchOnBaseConfigTests : EntityTestsBase
    {
        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();

            // Extra initialization for current test class
            GlobalConfiguration.FilterExactMatchOnly = true;
            GlobalConfiguration.Update();
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
    }

    [TestFixture]
    public class RequireExactMatchOnCustomConfigTests : CustomConfigTestsBase
    {
        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();

            // Extra initialization for current test class
            GlobalConfiguration.FilterExactMatchOnly = true;
            GlobalConfiguration.Update();
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void SearchEntities(SearchEntityData registrationData)
        {
            base.SearchEntities(registrationData);
        }
    }
}
