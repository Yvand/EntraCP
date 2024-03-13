using NUnit.Framework;
using System;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class RequireExactMatchOnBaseConfigTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.FilterExactMatchOnly = true;
            Settings.ClaimTypes.UpdateIdentifierForGuestUsers(Configuration.DirectoryObjectProperty.UserPrincipalName);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), null)]
        public void TestAllEntraIDUsers(EntraIdTestUser user)
        {
            // Input is not the full UPN value: it should not return any result
            TestSearchOperation(user.UserPrincipalName.Substring(0, 5), 0, String.Empty);
            // Input is exactly the UPN value: it should return 1 result
            TestSearchOperation(user.UserPrincipalName, 1, user.UserPrincipalName);
            TestValidationOperation(UserIdentifierClaimType, user.UserPrincipalName, true);
        }
    }
}
