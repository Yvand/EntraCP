﻿using NUnit.Framework;
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

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeGroups), new object[] { TestEntitySourceManager.MaxNumberOfGroupsToTest, true })]
        public void TestGroups(TestGroup group)
        {
            TestSearchAndValidateForTestGroup(group);
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

        //[Test]
        //public void TestRandomUsers([Random(0, UnitTestsHelper.TotalNumberTestUsers - 1, 5)] int idx)
        //{
        //    var user = EntraIdTestUsersSource.Users[idx];
        //    base.TestSearchAndValidateForTestUser(user);
        //}

        [Test]
        [Repeat(1)]
        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
        }

#if DEBUG
        [TestCase("testEntraCPUser_001")]
        [TestCase("testEntraCPUser_326")]
        public void DebugTestUser(string upnPrefix)
        {
            TestUser user = TestEntitySourceManager.FindUser(upnPrefix);
            base.TestSearchAndValidateForTestUser(user);
            base.TestAugmentationAgainst1RandomGroup(user);
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
