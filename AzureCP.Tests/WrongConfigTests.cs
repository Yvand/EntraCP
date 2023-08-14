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
        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();

            // Extra initialization for current test class
            ClaimTypeConfig randomClaimTypeConfig = new ClaimTypeConfig
            {
                ClaimType = UnitTestsHelper.RandomClaimType,
                EntityProperty = UnitTestsHelper.RandomObjectProperty,
            };
            Config.ClaimTypes = new ClaimTypeConfigCollection(SPTrust) { randomClaimTypeConfig };
            Config.Update();
        }

        [Test]
        public void ValiateInitialization()
        {
            Assert.IsNull(Config.RefreshLocalConfigurationIfNeeded(), "RefreshLocalConfigurationIfNeeded should return null because the configuration is not valid");
            Assert.IsFalse(UnitTestsHelper.ClaimsProvider.ValidateLocalConfiguration(null), "ValidateLocalConfiguration should return false because the configuration is not valid");
        }

        [TestCase(@"random", 0, "")]
        public override void SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.SearchEntities(inputValue, expectedResultCount, expectedEntityClaimValue);
        }
    }
}
