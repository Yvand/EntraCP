//using NUnit.Framework;
//using System;
//using System.Collections.Generic;
//using System.Diagnostics;
//using System.Linq;

//namespace Yvand.EntraClaimsProvider.Tests
//{
//    [TestFixture]
//    [Parallelizable(ParallelScope.Children)]
//    public class FilterUsersBasedOnSingleGroupTests : ClaimsProviderTestsBase
//    {
//        public override void InitializeSettings()
//        {
//            base.InitializeSettings();
//            Settings.RestrictSearchableUsersByGroups = EntraIdTestGroupsSource.GetSomeEntities(true, 1).ToArray()[0].Id;
//            base.ApplySettings();
//        }

//        [Test]
//        public override void CheckSettingsTest()
//        {
//            base.CheckSettingsTest();
//        }

//        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetSomeEntities), new object[] { UnitTestsHelper.MaxNumberOfUsersToTest })]
//        public void TestUsers(EntraIdTestUser user)
//        {
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }

//#if DEBUG
//        [TestCase("testEntraCPUser_001")]
//        [TestCase("testEntraCPUser_020")]
//        public void DebugTestUser(string upnPrefix)
//        {
//            EntraIdTestUser user = EntraIdTestUsersSource.Users.Find(x => x.UserPrincipalName.StartsWith(upnPrefix));
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }
//#endif
//    }

//    [TestFixture]
//    [Parallelizable(ParallelScope.Children)]
//    public class FilterUsersBasedOnMultipleGroupsTests : ClaimsProviderTestsBase
//    {
//        public override void InitializeSettings()
//        {
//            base.InitializeSettings();

//            // Pick the Id of 18 (max possible) random groups, and set them in property RestrictSearchableUsersByGroups
//            Settings.RestrictSearchableUsersByGroups = String.Join(",", EntraIdTestGroupsSource.GetSomeEntities(true, 18).Select(x => x.Id).ToArray());
//            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Set property RestrictSearchableUsersByGroups: \"{Settings.RestrictSearchableUsersByGroups}\".");
//            base.ApplySettings();
//        }

//        [Test]
//        public override void CheckSettingsTest()
//        {
//            base.CheckSettingsTest();
//        }

//        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetSomeEntities), new object[] { UnitTestsHelper.MaxNumberOfUsersToTest })]
//        public void TestUsers(EntraIdTestUser user)
//        {
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }

//#if DEBUG
//        [TestCase("testEntraCPUser_001")]
//        [TestCase("testEntraCPUser_020")]
//        public void DebugTestUser(string upnPrefix)
//        {
//            EntraIdTestUser user = EntraIdTestUsersSource.Users.Find(x => x.UserPrincipalName.StartsWith(upnPrefix));
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }
//#endif
//    }

//#if DEBUG
//    [TestFixture]
//    [Parallelizable(ParallelScope.Children)]
//    public class DebugFilterUsersBasedOnMultipleGroupsTests : ClaimsProviderTestsBase
//    {
//        public override void InitializeSettings()
//        {
//            base.InitializeSettings();
//            Settings.RestrictSearchableUsersByGroups = String.Join(",", EntraIdTestGroupsSource.GetSomeEntities(true, 18).Select(x => x.Id).ToArray());
//            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Set property RestrictSearchableUsersByGroups: \"{Settings.RestrictSearchableUsersByGroups}\".");
//            base.ApplySettings();
//        }

//        [TestCase("testEntraCPUser_001")]
//        [TestCase("testEntraCPUser_020")]
//        public void DebugTestUser(string upnPrefix)
//        {
//            EntraIdTestUser user = EntraIdTestUsersSource.Users.Find(x => x.UserPrincipalName.StartsWith(upnPrefix));
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }

//        [Test]
//        public void DebugGuestUser()
//        {
//            EntraIdTestUser user = EntraIdTestUsersSource.Users.Find(x => x.Mail.StartsWith("testEntraCPGuestUser_001"));
//            base.TestSearchAndValidateForEntraIDUser(user);
//        }
//    }
//#endif
//}
