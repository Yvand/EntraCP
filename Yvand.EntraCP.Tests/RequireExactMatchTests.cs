using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Diagnostics;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    public class RequireExactMatchOnBaseConfigTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings(bool applyChanges)
        {
            base.InitializeSettings(false);
            Settings.FilterExactMatchOnly = true;
            if (applyChanges)
            {
                TestSettingsAndApplyThemIfValid();
            }
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void TestSearch(SearchEntityData registrationData)
        {
            base.ProcessAndTestSearchEntityData(registrationData);
        }

        [TestCase(@"aadgroup1143", 1, "3f4b724c-125d-47b4-b989-195b29417d6e")]
        public void TestSearchManual(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }
    }

    [TestFixture]
    public class RequireExactMatchOnCustomConfigTests : CustomConfigTestsBase
    {
        public override void InitializeSettings(bool applyChanges)
        {
            base.InitializeSettings(false);
            Settings.FilterExactMatchOnly = true;
            if (applyChanges)
            {
                TestSettingsAndApplyThemIfValid();
            }
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void TestSearch(SearchEntityData registrationData)
        {
            base.ProcessAndTestSearchEntityData(registrationData);
        }
    }
}
