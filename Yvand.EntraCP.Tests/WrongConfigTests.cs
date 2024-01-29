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
        public override void InitializeSettings(bool applyChanges)
        {
            DoAugmentationTest = false;
            base.InitializeSettings(false);
            ClaimTypeConfig randomClaimTypeConfig = new ClaimTypeConfig
            {
                ClaimType = UnitTestsHelper.RandomClaimType,
                EntityProperty = UnitTestsHelper.RandomObjectProperty,
            };
            Settings.ClaimTypes = new ClaimTypeConfigCollection(UnitTestsHelper.SPTrust) { randomClaimTypeConfig };
            ConfigurationShouldBeValid = false;
            base.TestSettingsAndApplyThemIfValid();
        }

        ///// <summary>
        /////  Disable test augmentation with real data in this test class
        ///// </summary>
        ///// <param name="registrationData"></param>
        //public override void TestAugmentationOperation(ValidateEntityData registrationData)
        //{
        //}
    }
}
