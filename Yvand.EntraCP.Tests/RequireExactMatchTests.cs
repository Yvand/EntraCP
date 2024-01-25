using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Diagnostics;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    public class RequireExactMatchOnBaseConfigTests : EntityTestsBase
    {
        public override void InitializeConfiguration(bool applyChanges)
        {
            base.InitializeConfiguration(false);
            Settings.FilterExactMatchOnly = true;
            if (applyChanges)
            {
                GlobalConfiguration.ApplySettings(Settings, true);
                Trace.TraceInformation($"{DateTime.Now.ToString("s")} [RequireExactMatchOnBaseConfigTests] Updated configuration: {JsonConvert.SerializeObject(Settings, Formatting.None)}");
            }
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
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
        public override void InitializeConfiguration(bool applyChanges)
        {
            base.InitializeConfiguration(false);
            Settings.FilterExactMatchOnly = true;
            if(applyChanges)
            {
                GlobalConfiguration.ApplySettings(Settings, true);
                Trace.TraceInformation($"{DateTime.Now.ToString("s")} [RequireExactMatchOnCustomConfigTests] Updated configuration: {JsonConvert.SerializeObject(Settings, Formatting.None)}");
            }
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void SearchEntities(SearchEntityData registrationData)
        {
            base.SearchEntities(registrationData);
        }
    }
}
