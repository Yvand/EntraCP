using NUnit.Framework;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    /// <summary>
    /// Test guest accounts when their identity claim is the UserPrincipalName
    /// </summary>
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class GuestAccountsUPNTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.ClaimTypes.UpdateIdentifierForGuestUsers(DirectoryObjectProperty.UserPrincipalName);
            Settings.EnableAugmentation = true;
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), nameof(SearchEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void TestSearch(SearchEntityData registrationData)
        {
            base.ProcessAndTestSearchEntityData(registrationData);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void TestValidateClaim(ValidateEntityData registrationData)
        {
            base.ProcessAndTestValidateEntityData(registrationData);
        }

        //[Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        //[Repeat(UnitTestsHelper.TestRepeatCount)]
        //public override void AugmentEntity(ValidateEntityData registrationData)
        //{
        //    base.AugmentEntity(registrationData);
        //}
    }
}
