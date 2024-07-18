//using NUnit.Framework;

//namespace Yvand.EntraClaimsProvider.Tests
//{
//    [TestFixture]
//    [Parallelizable(ParallelScope.Children)]
//    public class ExcludeAllUserAccountsTests : ClaimsProviderTestsBase
//    {
//        public override bool ExcludeGuestUsers => true;
//        public override bool ExcludeMemberUsers => true;

//        public override void InitializeSettings()
//        {
//            base.InitializeSettings();
//            base.ApplySettings();
//        }

//        [Test]
//        public override void CheckSettingsTest()
//        {
//            base.CheckSettingsTest();
//        }

//        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetSomeEntities), new object[] { true, UnitTestsHelper.MaxNumberOfGroupsToTest })]
//        public void TestGroups(EntraIdTestGroup group)
//        {
//            TestSearchAndValidateForEntraIDGroup(group);
//        }

//        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetSomeEntities), new object[] { UnitTestsHelper.MaxNumberOfUsersToTest })]
//        public void TestUsers(EntraIdTestUser user)
//        {
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }

//        [Test]
//        [Repeat(5)]
//        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
//        {
//            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
//        }
//    }

//    [TestFixture]
//    [Parallelizable(ParallelScope.Children)]
//    public class ExcludeGuestUserAccountsTests : ClaimsProviderTestsBase
//    {
//        public override bool ExcludeGuestUsers => true;
//        public override bool ExcludeMemberUsers => false;

//        public override void InitializeSettings()
//        {
//            base.InitializeSettings();
//            base.ApplySettings();
//        }

//        [Test]
//        public override void CheckSettingsTest()
//        {
//            base.CheckSettingsTest();
//        }

//        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetSomeEntities), new object[] { true, UnitTestsHelper.MaxNumberOfGroupsToTest })]
//        public void TestGroups(EntraIdTestGroup group)
//        {
//            TestSearchAndValidateForEntraIDGroup(group);
//        }

//        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetSomeEntities), new object[] { UnitTestsHelper.MaxNumberOfUsersToTest })]
//        public void TestUsers(EntraIdTestUser user)
//        {
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }

//        [Test]
//        [Repeat(5)]
//        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
//        {
//            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
//        }
//    }

//    [TestFixture]
//    [Parallelizable(ParallelScope.Children)]
//    public class ExcludeMemberUserAccountsTests : ClaimsProviderTestsBase
//    {
//        public override bool ExcludeGuestUsers => false;
//        public override bool ExcludeMemberUsers => true;

//        public override void InitializeSettings()
//        {
//            base.InitializeSettings();
//            base.ApplySettings();
//        }

//        [Test]
//        public override void CheckSettingsTest()
//        {
//            base.CheckSettingsTest();
//        }

//        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetSomeEntities), new object[] { true, UnitTestsHelper.MaxNumberOfGroupsToTest })]
//        public void TestGroups(EntraIdTestGroup group)
//        {
//            TestSearchAndValidateForEntraIDGroup(group);
//        }

//        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetSomeEntities), new object[] { UnitTestsHelper.MaxNumberOfUsersToTest })]
//        public void TestUsers(EntraIdTestUser user)
//        {
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }

//        [Test]
//        [Repeat(5)]
//        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
//        {
//            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
//        }
//    }
//}
