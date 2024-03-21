using Microsoft.Graph.Models;
using NUnit.Framework;
using System;
using System.Linq;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class FilterUsersBasedOnSingleGroupTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.GroupsWhichUsersMustBeMemberOfAny = EntraIdTestGroupsSource.ASecurityEnabledGroup.Id;
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), null)]
        public void TestAllTestUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }

#if DEBUG
        [TestCase("testEntraCPUser_001")]
        [TestCase("testEntraCPUser_020")]
        public void DebugTestUser(string upnPrefix)
        {
            EntraIdTestUser user = EntraIdTestUsersSource.Users.Find(x => x.UserPrincipalName.StartsWith(upnPrefix));
            base.TestSearchAndValidateForEntraIDUser(user);
        }
#endif
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class FilterUsersBasedOnMultipleGroupsTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            string[] groupIds = EntraIdTestGroupsSource.Groups.Take(18).Select(x => x.Id).ToArray();
            Settings.GroupsWhichUsersMustBeMemberOfAny = String.Join(",", groupIds);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), null)]
        public void TestAllTestUsers(EntraIdTestUser user)
        {
            base.TestSearchAndValidateForEntraIDUser(user);
        }
    }
}
