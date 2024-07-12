using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class FilterUsersBasedOnSingleGroupTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.RestrictSearchableUsersByGroups = EntraIdTestGroupsSource.ASecurityEnabledGroup.Id;
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), new object[] { UnitTestsHelper.MaxNumberOfUsersToTest })]
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

            // Pick the Id of 18 (max possible) random groups, and it in property RestrictSearchableUsersByGroups
            List<string> groupIdsList = new List<string>();
            Random rnd = new Random();
            for (int groupsCount = 1; groupsCount <= 18; groupsCount++)
            {
                int randomIdx = rnd.Next(0, EntraIdTestGroupsSource.Groups.Count - 1);
                groupIdsList.Add(EntraIdTestGroupsSource.Groups[randomIdx].Id);
            }
            Settings.RestrictSearchableUsersByGroups = String.Join(",", groupIdsList);
            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Set property RestrictSearchableUsersByGroups: \"{Settings.RestrictSearchableUsersByGroups}\".");

            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestUsersSource), nameof(EntraIdTestUsersSource.GetTestData), new object[] { UnitTestsHelper.MaxNumberOfUsersToTest })]
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

#if DEBUG
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class DebugFilterUsersBasedOnMultipleGroupsTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.RestrictSearchableUsersByGroups = "dbcdaf68-5949-4a15-b2f8-f385b41a9fca,72860e7b-93d5-46f7-ab56-480112f76548,461e8865-0f39-4199-86c6-0c8e25ad57ea,33780f8d-8345-4402-bc6a-d9c7e5d40d36,2c6b2fde-4f89-417a-9d8c-5e459e1520cf,7059e0ae-0cbc-4f4d-9d87-fef7289a7f50,dbcdaf68-5949-4a15-b2f8-f385b41a9fca,3cc758d6-7198-470f-878a-e5cd41a35e02,b05ddc92-b639-4ba2-95b3-9fdbc1bff2f2,40a7290c-9b67-4b7c-98ff-0c9871544423,71aa60d9-c4d1-4eaf-b398-3099b12afd88,0f51d30e-e03c-44ea-8d13-df0ca0df7a16,0f51d30e-e03c-44ea-8d13-df0ca0df7a16,982369ed-e88f-4b21-9b89-29067a0fa326,a40d2ded-2ab0-463f-b000-b3351ca6341d,ce5a4725-e719-4c2d-89bc-1c356facde99,de600b84-29aa-470c-b6ca-1459591728fb,152456d1-73a0-46d6-ac02-0403f2b5593e";
            Settings.RestrictSearchableUsersByGroups = "db41d655-d796-43fb-9e23-351ee8b5bdb0,461e8865-0f39-4199-86c6-0c8e25ad57ea,21dd6198-c447-48dd-9ea8-347f804c4dec,dcf1e533-6d55-4b00-9788-f0d81e287c8a,21dd6198-c447-48dd-9ea8-347f804c4dec,719006f9-8eb0-48fd-8f95-556f07b0123b,d6896744-f16b-4802-9f13-0e2c9fa06274,dcf1e533-6d55-4b00-9788-f0d81e287c8a,2ae8ff19-0e4f-45cc-98ea-84f1c53e60f2,bea2607f-513a-4324-a6a1-620ee1c0ced4,bcd82b83-97c5-4c1c-9cce-58643a286298,34fe3af4-7ed9-4b5e-a64e-c0b230d5dfb4,136f71a2-c57c-4a3f-8aec-4d694e442b87,dbf5a5c3-5f51-42d6-a519-258c76960f75,33780f8d-8345-4402-bc6a-d9c7e5d40d36,a40d2ded-2ab0-463f-b000-b3351ca6341d,21dd6198-c447-48dd-9ea8-347f804c4dec,34fe3af4-7ed9-4b5e-a64e-c0b230d5dfb4";
            Settings.RestrictSearchableUsersByGroups = "71aa60d9-c4d1-4eaf-b398-3099b12afd88,56e7a2e2-f565-450e-94cd-0d7d314217d1,c9a94341-89b5-4109-a501-2a14027b5bf0,8962bad6-ceca-43ff-a4be-9258ff81af2f,36d78c5a-80f2-4f3d-8f37-3a347d000d56,152456d1-73a0-46d6-ac02-0403f2b5593e,d6896744-f16b-4802-9f13-0e2c9fa06274,de600b84-29aa-470c-b6ca-1459591728fb,d51f225f-b484-4898-b425-5d48553aad16,36d78c5a-80f2-4f3d-8f37-3a347d000d56,0f51d30e-e03c-44ea-8d13-df0ca0df7a16,1eb5e51e-0bea-40c3-9ffb-0b85dbe2f9bf,dcf1e533-6d55-4b00-9788-f0d81e287c8a,c9a94341-89b5-4109-a501-2a14027b5bf0,3cc758d6-7198-470f-878a-e5cd41a35e02,dcf1e533-6d55-4b00-9788-f0d81e287c8a,21dd6198-c447-48dd-9ea8-347f804c4dec,253fe6d4-8e07-49a1-8a74-143d265eefbe";
            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Set property RestrictSearchableUsersByGroups: \"{Settings.RestrictSearchableUsersByGroups}\".");
            base.ApplySettings();
        }

        [TestCase("testEntraCPUser_001")]
        [TestCase("testEntraCPUser_020")]
        public void DebugTestUser(string upnPrefix)
        {
            EntraIdTestUser user = EntraIdTestUsersSource.Users.Find(x => x.UserPrincipalName.StartsWith(upnPrefix));
            base.TestSearchAndValidateForEntraIDUser(user);
        }

        [Test]
        public void DebugGuestUser()
        {
            EntraIdTestUser user = EntraIdTestUsersSource.Users.Find(x => x.Mail.StartsWith("testEntraCPGuestUser_001"));
            base.TestSearchAndValidateForEntraIDUser(user);
        }
    }
#endif
}
