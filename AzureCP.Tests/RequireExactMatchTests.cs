using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System.Security.Claims;

namespace AzureCP.Tests
{
    [TestFixture]
    public class RequireExactMatchOnBaseConfigTests : BackupCurrentConfig
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
            UnitTestsHelper.TestSearchOperation(registrationData.ClaimValue, expectedCount, registrationData.ClaimValue);
        }
    }

    [TestFixture]
    public class RequireExactMatchOnCustomConfigTests : CustomConfigTests
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
            UnitTestsHelper.TestSearchOperation(registrationData.ClaimValue, expectedCount, registrationData.ClaimValue);
        }

        [TestCase("value1", 1, "value1")]
        public override void SearchAndValidateExtensionAttributeTest(string inputValue, int expectedCount, string expectedClaimValue)
        {
            base.SearchAndValidateExtensionAttributeTest(inputValue, expectedCount, expectedClaimValue);
        }
    }
}
