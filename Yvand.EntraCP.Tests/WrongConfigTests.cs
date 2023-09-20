using Microsoft.SharePoint;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Yvand.EntraClaimsProvider.Config;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    public class WrongConfigBadClaimTypeTests : EntityTestsBase
    {
        public override bool ConfigurationIsValid => false;
        public override void InitializeConfiguration(bool applyChanges)
        {
            base.InitializeConfiguration(false);
            ClaimTypeConfig randomClaimTypeConfig = new ClaimTypeConfig
            {
                ClaimType = UnitTestsHelper.RandomClaimType,
                EntityProperty = UnitTestsHelper.RandomObjectProperty,
            };
            Settings.ClaimTypes = new ClaimTypeConfigCollection(UnitTestsHelper.SPTrust) { randomClaimTypeConfig };
            if (applyChanges)
            {
                GlobalConfiguration.ApplySettings(Settings, true);
            }
        }

        [TestCase(@"random", 0, "")]
        [TestCase(@"aad", 0, "")]
        public override void SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.SearchEntities(inputValue, expectedResultCount, expectedEntityClaimValue);
        }
    }

    [TestFixture]
    public class WrongConfigNoTenantTests : EntityTestsBase
    {
        public override bool ConfigurationIsValid => false;
        public override void InitializeConfiguration(bool applyChanges)
        {
            base.InitializeConfiguration(false);
            Settings.EntraIDTenantList = new List<EntraIDTenant>();
            if (applyChanges)
            {
                GlobalConfiguration.ApplySettings(Settings, true);
            }
        }

        [TestCase(@"random", 0, "")]
        [TestCase(@"aad", 0, "")]
        public override void SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.SearchEntities(inputValue, expectedResultCount, expectedEntityClaimValue);
        }
    }
}
