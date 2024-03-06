using NUnit.Framework;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class RequireExactMatchOnBaseConfigTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.FilterExactMatchOnly = true;
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        //[Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        //[Repeat(UnitTestsHelper.TestRepeatCount)]
        //public void TestSearch(SearchEntityData registrationData)
        //{
        //    base.ProcessAndTestSearchEntityData(registrationData);
        //}

        //[TestCase(@"aadgroup1143", 1, "3f4b724c-125d-47b4-b989-195b29417d6e")]
        //public void TestSearchManual(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        //{
        //    base.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        //}

        //[Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        //[Repeat(UnitTestsHelper.TestRepeatCount)]
        //public void TestAugmentationOperation(ValidateEntityData registrationData)
        //{
        //    base.TestAugmentationOperation(registrationData.ClaimValue, registrationData.IsMemberOfTrustedGroup);
        //}
    }
}
