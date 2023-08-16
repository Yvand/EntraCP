using Microsoft.SharePoint;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yvand.ClaimsProviders.Config;

namespace Yvand.ClaimsProviders.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class WrongConfigTests : EntityTestsBase
    {
        public override bool ConfigurationIsValid => false;
        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();
            ClaimTypeConfig randomClaimTypeConfig = new ClaimTypeConfig
            {
                ClaimType = UnitTestsHelper.RandomClaimType,
                EntityProperty = UnitTestsHelper.RandomObjectProperty,
            };
            Settings.ClaimTypes = new ClaimTypeConfigCollection(UnitTestsHelper.SPTrust) { randomClaimTypeConfig };
            GlobalConfiguration.ApplySettings(Settings, true);
        }

        [TestCase(@"random", 0, "")]
        public override void SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.SearchEntities(inputValue, expectedResultCount, expectedEntityClaimValue);
        }
    }
}
