using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Yvand.EntraClaimsProvider.Configuration;

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
            base.ValidateConfiguration();
        }
    }
}
