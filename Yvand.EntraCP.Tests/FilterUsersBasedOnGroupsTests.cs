using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
            Settings.RestrictSearchableUsersByGroups = TestEntitySourceManager.GetOneGroup(true).Id;
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeUsers), new object[] { TestEntitySourceManager.MaxNumberOfUsersToTest })]
        public void TestUsers(TestUser user)
        {
            base.TestSearchAndValidateForTestUser(user);
            base.TestAugmentationAgainst1RandomGroup(user);
        }

        [Test]
        public void TestAGuestUser()
        {
            TestUser user = TestEntitySourceManager.GetOneUser(UserType.Guest);
            base.TestSearchAndValidateForTestUser(user);
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeGroups), new object[] { TestEntitySourceManager.MaxNumberOfGroupsToTest, true })]
        public void TestGroups(TestGroup group)
        {
            TestSearchAndValidateForTestGroup(group);
        }

#if DEBUG
        [TestCase("testEntraCPUser_001")]
        [TestCase("testEntraCPUser_020")]
        public void DebugTestUser(string upnPrefix)
        {
            TestUser user = TestEntitySourceManager.AllTestUsers.Find(x => x.UserPrincipalName.StartsWith(upnPrefix));
            base.TestSearchAndValidateForTestUser(user);
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

            // Pick the Id of 18 (max possible) random groups, and set them in property RestrictSearchableUsersByGroups
            Settings.RestrictSearchableUsersByGroups = String.Join(",", TestEntitySourceManager.GetSomeGroups(18, true).Select(x => x.Id).ToArray());
            //Settings.RestrictSearchableUsersByGroups = "3c1c6c1a-2565-4cfd-b5f8-8ec732f93077,3c98541c-9601-47c0-aeea-fc0679b9d756,807c95cd-88de-49d9-a06e-12ce2329dfb7,807c95cd-88de-49d9-a06e-12ce2329dfb7,1beb24dd-0fae-46cb-b321-dd0baf5c9ecc,01572e9f-4a9a-4dd1-9314-05972d87d1c2,89d4f192-8eb0-4011-ada7-4a1d4f678b1c,bdd53ff1-866c-442b-b6d5-ac43b4306aa7,2d407401-192c-4a25-9f0e-3693cfad6f27,1c607c55-f1a0-408c-ae52-306cd89de742,1090383f-7ea5-4a16-9ba8-0551a061d7f9,874b1dcf-aa82-428a-b107-b71a09c3d452,1090383f-7ea5-4a16-9ba8-0551a061d7f9,1831bd90-e413-4b86-a8ab-5d26d8a75498,04ec1e1c-196d-4b85-85b2-c3b982644114,043e997e-0b2c-412c-b8f0-13253251569c,40d53a73-130b-48e1-946b-5fec5ec35d4f,3c98541c-9601-47c0-aeea-fc0679b9d756";
            //Settings.RestrictSearchableUsersByGroups = "3c98541c-9601-47c0-aeea-fc0679b9d756";
            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Set property RestrictSearchableUsersByGroups: \"{Settings.RestrictSearchableUsersByGroups}\".");
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeUsers), new object[] { TestEntitySourceManager.MaxNumberOfUsersToTest })]
        public void TestUsers(TestUser user)
        {
            base.TestSearchAndValidateForTestUser(user);
            base.TestAugmentationAgainst1RandomGroup(user);
        }

        [Test]
        public void TestAGuestUser()
        {
            TestUser user = TestEntitySourceManager.GetOneUser(UserType.Guest);
            base.TestSearchAndValidateForTestUser(user);
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeGroups), new object[] { TestEntitySourceManager.MaxNumberOfGroupsToTest, true })]
        public void TestGroups(TestGroup group)
        {
            TestSearchAndValidateForTestGroup(group);
        }

#if DEBUG
        [TestCase("testEntraCPUser_001")]
        [TestCase("testEntraCPUser_020")]
        public void DebugTestUser(string upnPrefix)
        {
            TestUser user = TestEntitySourceManager.FindUser(upnPrefix);
            base.TestSearchAndValidateForTestUser(user);
        }
#endif
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class DebugFilterUsersBasedOnMultipleGroupsTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.RestrictSearchableUsersByGroups = String.Join(",", TestEntitySourceManager.GetSomeGroups(18, true).Select(x => x.Id).ToArray());
            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Set property RestrictSearchableUsersByGroups: \"{Settings.RestrictSearchableUsersByGroups}\".");
            base.ApplySettings();
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeUsers), new object[] { TestEntitySourceManager.MaxNumberOfUsersToTest })]
        public void TestUsers(TestUser user)
        {
            base.TestSearchAndValidateForTestUser(user);
            base.TestAugmentationAgainst1RandomGroup(user);
        }

        [Test]
        public void TestAGuestUser()
        {
            TestUser user = TestEntitySourceManager.GetOneUser(UserType.Guest);
            base.TestSearchAndValidateForTestUser(user);
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeGroups), new object[] { TestEntitySourceManager.MaxNumberOfGroupsToTest, true })]
        public void TestGroups(TestGroup group)
        {
            TestSearchAndValidateForTestGroup(group);
        }

        [TestCase("testEntraCPUser_001")]
        [TestCase("testEntraCPUser_020")]
        public void DebugTestUser(string upnPrefix)
        {
            TestUser user = TestEntitySourceManager.FindUser(upnPrefix);
            base.TestSearchAndValidateForTestUser(user);
        }
    }
}
