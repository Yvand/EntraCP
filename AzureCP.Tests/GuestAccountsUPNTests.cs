using NUnit.Framework;
using Yvand.Config;

namespace Yvand.Tests
{
    /// <summary>
    /// Test guest accounts when their identity claim is the UserPrincipalName
    /// </summary>
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class GuestAccountsUPNTests : EntityTestsBase
    {
        public override void InitializeConfiguration(bool applyChanges)
        {
            base.InitializeConfiguration(false);

            // Extra initialization for current test class
            Settings.ClaimTypes.UpdateIdentifierForGuestUsers(DirectoryObjectProperty.UserPrincipalName);
            Settings.EnableAugmentation = true;
            if (applyChanges)
            {
                GlobalConfiguration.ApplySettings(Settings, true);
            }
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void SearchEntities(SearchEntityData registrationData)
        {
            base.SearchEntities(registrationData);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void ValidateClaim(ValidateEntityData registrationData)
        {
            base.ValidateClaim(registrationData);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.UPNB2BGuestAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void AugmentEntity(ValidateEntityData registrationData)
        {
            base.AugmentEntity(registrationData);
        }        
    }
}
