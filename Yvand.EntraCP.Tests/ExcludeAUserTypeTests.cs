using NUnit.Framework;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class ExcludeAllUserAccountsTests : ClaimsProviderTestsBase
    {
        public override bool ExcludeGuestUsers => true;
        public override bool ExcludeMemberUsers => true;

        public override void InitializeSettings()
        {
            base.InitializeSettings();
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetTestData), new object[] { true })]
        public void TestAllEntraIDGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), null)]
        public void TestAllEntraIDUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationForUsersMembersOfAllGroups()
        {
            base.TestAugmentationForUsersMembersOfAllGroups();
        }
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class ExcludeGuestUserAccountsTests : ClaimsProviderTestsBase
    {
        public override bool ExcludeGuestUsers => true;
        public override bool ExcludeMemberUsers => false;

        public override void InitializeSettings()
        {
            base.InitializeSettings();
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetTestData), new object[] { true })]
        public void TestAllEntraIDGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), null)]
        public void TestAllEntraIDUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationForUsersMembersOfAllGroups()
        {
            base.TestAugmentationForUsersMembersOfAllGroups();
        }
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class ExcludeMemberUserAccountsTests : ClaimsProviderTestsBase
    {
        public override bool ExcludeGuestUsers => false;
        public override bool ExcludeMemberUsers => true;

        public override void InitializeSettings()
        {
            base.InitializeSettings();
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetTestData), new object[] { true })]
        public void TestAllEntraIDGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), null)]
        public void TestAllEntraIDUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationForUsersMembersOfAllGroups()
        {
            base.TestAugmentationForUsersMembersOfAllGroups();
        }
    }
}
