using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System.Security.Claims;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class BypassDirectoryOnClaimTypesTests : ClaimsProviderTestsBase
    {
        const string PrefixBypassUserSearch = "bypass-user:";
        const string PrefixBypassGroupSearch = "bypass-group:";
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.EnableAugmentation = true;
            Settings.ClaimTypes.UserIdentifierConfig.PrefixToBypassLookup = PrefixBypassUserSearch;
            Settings.ClaimTypes.GroupIdentifierConfig.PrefixToBypassLookup = PrefixBypassGroupSearch;
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
            user.UserPrincipalName = user.DisplayName;
            user.Mail = user.DisplayName;
            user.DisplayName = $"{PrefixBypassUserSearch}{user.DisplayName}";
            base.TestSearchAndValidateForTestUser(user);
        }

        [Test, TestCaseSource(typeof(TestEntitySourceManager), nameof(TestEntitySourceManager.GetSomeGroups), new object[] { TestEntitySourceManager.MaxNumberOfGroupsToTest, true })]
        public void TestGroups(TestGroup group)
        {
            TestSearchAndValidateForTestGroup(group);
            group.Id = group.DisplayName;
            group.DisplayName = $"{PrefixBypassGroupSearch}{group.DisplayName}";
            TestSearchAndValidateForTestGroup(group);
        }

        [TestCase(PrefixBypassUserSearch + "externalUser@contoso.com", 1, "externalUser@contoso.com")]
        [TestCase(PrefixBypassUserSearch, 0, "")]
        [TestCase(PrefixBypassGroupSearch, 0, "")]
        public void TestBypassDirectoryByClaimType(string inputValue, int expectedCount, string expectedClaimValue)
        {
            TestSearchOperation(inputValue, expectedCount, expectedClaimValue);

            if (expectedCount > 0)
            {
                SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, expectedClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                TestValidationOperation(inputClaim, true, expectedClaimValue);
            }
        }
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class BypassDirectoryGloballyTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.AlwaysResolveUserInput = true;
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
    }
}
