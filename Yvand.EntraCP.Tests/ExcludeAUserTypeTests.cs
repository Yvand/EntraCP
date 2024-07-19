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

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeGroups), new object[] { TestEntitySourceManager.MaxNumberOfGroupsToTest, true })]
        public void TestGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeUsers), new object[] { TestEntitySourceManager.MaxNumberOfUsersToTest })]
        public void TestUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
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

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeGroups), new object[] { TestEntitySourceManager.MaxNumberOfGroupsToTest, true })]
        public void TestGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeUsers), new object[] { TestEntitySourceManager.MaxNumberOfUsersToTest })]
        public void TestUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
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

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeGroups), new object[] { TestEntitySourceManager.MaxNumberOfGroupsToTest, true })]
        public void TestGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeUsers), new object[] { TestEntitySourceManager.MaxNumberOfUsersToTest })]
        public void TestUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
        }
    }
}
