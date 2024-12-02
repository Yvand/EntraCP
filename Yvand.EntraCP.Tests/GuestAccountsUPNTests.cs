using NUnit.Framework;
using System;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    /// <summary>
    /// Test guest accounts when their identity claim is the UserPrincipalName
    /// </summary>
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class GuestAccountsUPNTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.ClaimTypes.UpdateIdentifierForGuestUsers(DirectoryObjectProperty.UserPrincipalName);
            Settings.EnableAugmentation = true;
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
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
        }

        [Test]
        [Repeat(5)]
        public void TestAGuestUser()
        {
            TestUser user = TestEntitySourceManager.GetOneUser(UserType.Guest);
            base.TestSearchAndValidateForTestUser(user);
        }

        [Test]
        [Repeat(5)]
        public void TestValidationOfGuestUser()
        {
            TestUser user = TestEntitySourceManager.GetOneUser(UserType.Guest);
            // Test below must validate, because DirectoryObjectPropertyForGuestUsers UserPrincipalName
            base.TestValidationOperation(UserIdentifierClaimType, user.UserPrincipalName, true);
            // Test below must NOT validate, because DirectoryObjectPropertyForGuestUsers UserPrincipalName, NOT mail
            base.TestValidationOperation(UserIdentifierClaimType, user.Mail, false);
        }

#if DEBUG
        [Test]
        public void DebugAugmentTestUser()
        {
            TestUser user = TestEntitySourceManager.GetOneUser(UserType.Guest);
            base.TestAugmentationOperation(user.UserPrincipalName, user.IsMemberOfAllGroups, String.Empty);
        }
#endif
    }
}
