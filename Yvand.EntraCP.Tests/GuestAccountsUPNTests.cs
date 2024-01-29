using Newtonsoft.Json;
using NUnit.Framework;
using System.Diagnostics;
using System;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    /// <summary>
    /// Test guest accounts when their identity claim is the UserPrincipalName
    /// </summary>
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class GuestAccountsUPNTests : EntityTestsBase
    {
        public override void InitializeSettings(bool applyChanges)
        {
            base.InitializeSettings(false);

            // Extra initialization for current test class
            Settings.ClaimTypes.UpdateIdentifierForGuestUsers(DirectoryObjectProperty.UserPrincipalName);
            Settings.EnableAugmentation = true;
            if (applyChanges)
            {
                TestSettingsAndApplyThemIfValid();
            }
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void SearchEntities(SearchEntityData registrationData)
        {
            base.SearchEntities(registrationData);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void ValidateClaim(ValidateEntityData registrationData)
        {
            base.ValidateClaim(registrationData);
        }

        //[Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        //[Repeat(UnitTestsHelper.TestRepeatCount)]
        //public override void AugmentEntity(ValidateEntityData registrationData)
        //{
        //    base.AugmentEntity(registrationData);
        //}
    }
}
