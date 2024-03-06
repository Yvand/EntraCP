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

        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetTestData), new object[] { true })]
        public void TestAllEntraIDGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test]
        public void TestSomeEntraIDGroups([Random(0, UnitTestsHelper.TotalNumberOfGroupsInSource - 1, 5)] int idx)
        {
            EntraIdTestGroup group = EntraIdTestGroupsSource.Groups[idx];
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), null)]
        public void TestAllEntraIDUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        public void TestSomeEntraIDUsers([Random(0, UnitTestsHelper.TotalNumberOfUsersInSource - 1, 5)] int idx)
        {
            var user = EntraIdTestUsersSource.Users[idx];
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationForUsersMembersOfAllGroups()
        {
            base.TestAugmentationForUsersMembersOfAllGroups();
        }

#if DEBUG
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
