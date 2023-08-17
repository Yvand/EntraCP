using Microsoft.SharePoint;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Yvand.ClaimsProviders.Config;

namespace Yvand.ClaimsProviders.Tests
{
    public class WrongConfigBadClaimTypeTests : EntityTestsBase
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

    public class WrongConfigNoTenantTests : EntityTestsBase
    {
        public override bool ConfigurationIsValid => false;
        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();
            Settings.AzureTenants = new List<AzureTenant>();
            GlobalConfiguration.ApplySettings(Settings, true);
        }
    }
}
