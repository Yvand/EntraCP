using NUnit.Framework;
using System;
using System.Linq;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    internal class BasicConfigurationTests : ClaimsProviderTestsBase
    {
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

        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetTestData), new object[] { true, UnitTestsHelper.MaxNumberOfGroupsToTest })]
        public void TestGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        //[Test]
        //public void TestRandomTestGroups([Random(0, UnitTestsHelper.TotalNumberTestGroups - 1, 5)] int idx)
        //{
        //    EntraIdTestGroup group = EntraIdTestGroupsSource.Groups[idx];
        //    TestSearchAndValidateForEntraIDGroup(group);
        //}

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), new object[] { UnitTestsHelper.MaxNumberOfUsersToTest })]
        public void TestUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        //[Test]
        //public void TestRandomTestUsers([Random(0, UnitTestsHelper.TotalNumberTestUsers - 1, 5)] int idx)
        //{
        //    var user = EntraIdTestUsersSource.Users[idx];
        //    base.TestSearchAndValidateForEntraIDUser(user);
        //}

        [Test]
        [Repeat(5)]
        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
        }

#if DEBUG
        [TestCase("testEntraCPUser_001")]
        [TestCase("testEntraCPUser_020")]
        public void DebugTestUser(string upnPrefix)
        {
            EntraIdTestUser user = EntraIdTestUsersSource.Users.Find(x => x.UserPrincipalName.StartsWith(upnPrefix));
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [TestCase(@"testentracp", 30, "")]
        public void DebugSearchManual(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            base.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn", "FakeGuest.com#EXT#@XXX.onmicrosoft.com", false)]
        public void DebugValidateClaim(string claimType, string claimValue, bool shouldValidate)
        {
            base.TestValidationOperation(claimType, claimValue, shouldValidate);
        }

        [TestCase("xydGUEST@FAKE.onmicrosoft.com", false)]
        public void DebugAugmentEntity(string claimValue, bool shouldHavePermissions)
        {
            base.TestAugmentationOperation(claimValue, shouldHavePermissions, String.Empty);
        }
#endif
    }
}
