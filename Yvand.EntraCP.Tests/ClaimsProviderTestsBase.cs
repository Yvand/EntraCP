using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Claims;
using System.Text;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    public class ClaimsProviderTestsBase
    {
        /// <summary>
        /// Configures whether to exclude AAD Guest users from search and validation. This does not impact augmentation.
        /// </summary>
        public virtual bool ExcludeGuestUsers => false;

        /// <summary>
        /// Configures whether to exclude AAD Member users from search and validation. This does not impact augmentation.
        /// </summary>
        public virtual bool ExcludeMemberUsers => false;

        /// <summary>
        /// Configures whether the configuration applied is valid, and whether the claims provider should be able to use it
        /// </summary>
        public bool ConfigurationShouldBeValid = true;

        public string UserIdentifierClaimType
        {
            get
            {
                return Settings.ClaimTypes.UserIdentifierConfig.ClaimType;
            }
        }

        public string GroupIdentifierClaimType
        {
            get
            {
                return Settings.ClaimTypes.GroupIdentifierConfig.ClaimType;
            }
        }

        protected EntraIDProviderSettings Settings = new EntraIDProviderSettings();

        private object _LockVerifyIfCurrentUserShouldBeFound = new object();
        private object _LockInitGroupsWhichUsersMustBeMemberOfAny = new object();
        private List<EntraIdTestGroupSettings> _GroupsWhichUsersMustBeMemberOfAny;
        protected List<EntraIdTestGroupSettings> GroupsWhichUsersMustBeMemberOfAny
        {
            get
            {
                if (_GroupsWhichUsersMustBeMemberOfAny != null) { return _GroupsWhichUsersMustBeMemberOfAny; }
                lock (_LockInitGroupsWhichUsersMustBeMemberOfAny)
                {
                    if (_GroupsWhichUsersMustBeMemberOfAny != null) { return _GroupsWhichUsersMustBeMemberOfAny; }
                    _GroupsWhichUsersMustBeMemberOfAny = new List<EntraIdTestGroupSettings>();
                    string groupsWhichUsersMustBeMemberOfAny = Settings.RestrictSearchableUsersByGroups;
                    if (!String.IsNullOrWhiteSpace(groupsWhichUsersMustBeMemberOfAny))
                    {
                        string[] groupIds = groupsWhichUsersMustBeMemberOfAny.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string groupId in groupIds)
                        {
                            EntraIdTestGroupSettings groupSettings = EntraIdTestGroupSettings.GroupsSettings.FirstOrDefault(x => x.Id == groupId);
                            if (groupSettings == null) { groupSettings = new EntraIdTestGroupSettings(); }
                            _GroupsWhichUsersMustBeMemberOfAny.Add(groupSettings);
                        }
                    }
                    return _GroupsWhichUsersMustBeMemberOfAny;
                }
            }
        }


        /// <summary>
        /// Initialize settings
        /// </summary>
        [OneTimeSetUp]
        public virtual void InitializeSettings()
        {
            Settings = new EntraIDProviderSettings();
            Settings.ClaimTypes = EntraIDProviderSettings.ReturnDefaultClaimTypesConfig(UnitTestsHelper.ClaimsProvider.Name);
            Settings.ProxyAddress = TestContext.Parameters["ProxyAddress"];

#if DEBUG
            Settings.Timeout = 99999;
#endif

            Settings.EntraIDTenants = new List<EntraIDTenant> { UnitTestsHelper.TenantConnection };
            foreach (EntraIDTenant tenant in Settings.EntraIDTenants)
            {
                tenant.ExcludeMemberUsers = ExcludeMemberUsers;
                tenant.ExcludeGuestUsers = ExcludeGuestUsers;
            }
            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Initialized default settings.");
        }

        /// <summary>
        /// Override this method and decorate it with [Test] if the settings applied in the inherited class should be tested
        /// </summary>
        public virtual void CheckSettingsTest()
        {
            UnitTestsHelper.PersistedConfiguration.ApplySettings(Settings, false);
            if (ConfigurationShouldBeValid)
            {
                Assert.DoesNotThrow(() => UnitTestsHelper.PersistedConfiguration.ValidateConfiguration(), "ValidateLocalConfiguration should NOT throw a InvalidOperationException because the configuration is valid");
            }
            else
            {
                Assert.Throws<InvalidOperationException>(() => UnitTestsHelper.PersistedConfiguration.ValidateConfiguration(), "ValidateLocalConfiguration should throw a InvalidOperationException because the configuration is invalid");
                Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Invalid configuration: {JsonConvert.SerializeObject(Settings, Formatting.None)}");
            }
        }

        public void TestSearchAndValidateForEntraIDGroup(EntraIdTestGroup entity)
        {
            string inputValue = entity.DisplayName;
            int expectedCount = 1;
            bool shouldValidate = true;

            if (Settings.AlwaysResolveUserInput)
            {
                inputValue = entity.Id;
                expectedCount = Settings.ClaimTypes.GetConfigsMappedToClaimType().Count();
            }
            if (Settings.FilterSecurityEnabledGroupsOnly && entity.SecurityEnabled == false)
            {
                expectedCount = 0;
                shouldValidate = false;
            }

            TestSearchOperation(inputValue, expectedCount, entity.Id);
            TestValidationOperation(GroupIdentifierClaimType, entity.Id, shouldValidate);
        }

        public void TestSearchAndValidateForEntraIDUser(EntraIdTestUser entity)
        {
            int expectedCount = 1;
            string inputValue = entity.DisplayName;
            string claimValue = entity.UserPrincipalName;
            bool shouldValidate = true;

            if (Settings.AlwaysResolveUserInput)
            {
                inputValue = entity.UserPrincipalName;
                claimValue = entity.UserPrincipalName;
                expectedCount = Settings.ClaimTypes.GetConfigsMappedToClaimType().Count();
            }
            else
            {
                if (!String.IsNullOrWhiteSpace(Settings.RestrictSearchableUsersByGroups))
                {
                    lock (_LockVerifyIfCurrentUserShouldBeFound) // TODO: understand why this lock is necessary
                    {
                        // Test 1: Does Settings.RestrictSearchableUsersByGroups contain any group where all test users are members?
                        bool groupWithAllTestUsersAreMembersFound = false;
                        foreach (var groupSettings in GroupsWhichUsersMustBeMemberOfAny)
                        {
                            if (groupSettings.AllTestUsersAreMembers)
                            {
                                groupWithAllTestUsersAreMembersFound = true;
                                Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] User \"{entity.UserPrincipalName}\" may be found because Settings.RestrictSearchableUsersByGroups contains group: \"{groupSettings.DisplayName}\" with AllTestUsersAreMembers {groupSettings.AllTestUsersAreMembers}.");
                                break;  // No need to change shouldValidate, which is true by default, or process other groups
                            }
                        }

                        // Test 2: If test 1 is false, is current entity member of all the test groups?
                        if (!groupWithAllTestUsersAreMembersFound)
                        {

                            EntraIdTestUserSettings userSettings = EntraIdTestUserSettings.UsersWithSpecificSettings.FirstOrDefault(x => String.Equals(x.UserPrincipalName, entity.UserPrincipalName, StringComparison.InvariantCultureIgnoreCase));
                            if (userSettings == null) { userSettings = new EntraIdTestUserSettings(); }
                            if (!userSettings.IsMemberOfAllGroups)
                            {
                                shouldValidate = false;
                                expectedCount = 0;
                                Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] User \"{entity.UserPrincipalName}\" should not be found because it has IsMemberOfAllGroups {userSettings.IsMemberOfAllGroups} and no group set in Settings.RestrictSearchableUsersByGroups has AllTestUsersAreMembers set to true.");
                            }
                        }
                    }
                }
                else
                {
                    Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Property Settings.RestrictSearchableUsersByGroups IsNullOrWhiteSpace.");
                }

                // If shouldValidate is false, user should not be found anyway so no need to do additional checks
                if (shouldValidate)
                {
                    if (entity.UserType == UserType.Member)
                    {
                        claimValue = Settings.ClaimTypes.UserIdentifierConfig.EntityProperty == Configuration.DirectoryObjectProperty.UserPrincipalName ?
                            entity.UserPrincipalName :
                            entity.Mail;
                        expectedCount = ExcludeMemberUsers ? 0 : 1;
                        shouldValidate = !ExcludeMemberUsers;
                    }
                    else
                    {
                        claimValue = Settings.ClaimTypes.UserIdentifierConfig.DirectoryObjectPropertyForGuestUsers == Configuration.DirectoryObjectProperty.UserPrincipalName ?
                            entity.UserPrincipalName :
                            entity.Mail;
                        expectedCount = ExcludeGuestUsers ? 0 : 1;
                        shouldValidate = !ExcludeGuestUsers;
                    }
                }
            }
            TestSearchOperation(inputValue, expectedCount, claimValue);
            TestValidationOperation(UserIdentifierClaimType, claimValue, shouldValidate);
        }

        /// <summary>
        /// Gold users are the test users who are members of all the test groups
        /// </summary>
        public virtual void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            Random rnd = new Random();
            int randomIdx = rnd.Next(0, TestEntitySourceManager.AllTestGroups.Count - 1);
            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] TestAugmentationOfGoldUsersAgainstRandomGroups: Get group in EntraIdTestGroupsSource.Groups at index {randomIdx}.");
            EntraIdTestGroup randomGroup = null;
            try
            {
                randomGroup = TestEntitySourceManager.AllTestGroups[randomIdx];
            }
            catch (ArgumentOutOfRangeException)
            {
                string errorMessage = $"{DateTime.Now:s} [{this.GetType().Name}] TestAugmentationOfGoldUsersAgainstRandomGroups: Could not get group in EntraIdTestGroupsSource.Groups at index {randomIdx}. EntraIdTestGroupsSource.Groups has {TestEntitySourceManager.AllTestGroups.Count} items.";
                Trace.TraceError(errorMessage);
                throw new ArgumentOutOfRangeException(errorMessage);
            }
            bool shouldBeMember = Settings.FilterSecurityEnabledGroupsOnly && !randomGroup.SecurityEnabled ? false : true;

            foreach (string userPrincipalName in EntraIdTestUserSettings.UsersWithSpecificSettings.Where(x => x.IsMemberOfAllGroups).Select(x => x.UserPrincipalName))
            {
                TestAugmentationOperation(userPrincipalName, shouldBeMember, randomGroup.Id);
            }
        }

        /// <summary>
        /// Applies the <see cref="Settings"/> to the configuration object and save it in the configuration database
        /// </summary>
        public void ApplySettings()
        {
            UnitTestsHelper.PersistedConfiguration.ApplySettings(Settings, true);
            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Updated configuration: {JsonConvert.SerializeObject(Settings, Formatting.None)}");
        }

        [OneTimeTearDown]
        public void Cleanup()
        {
            Trace.TraceInformation($"{DateTime.Now:s} [{this.GetType().Name}] Cleanup.");
        }

        /// <summary>
        /// Start search operation on a specific claims provider
        /// </summary>
        /// <param name="inputValue"></param>
        /// <param name="expectedCount">How many entities are expected to be returned. Set to Int32.MaxValue if exact number is unknown but greater than 0</param>
        /// <param name="expectedClaimValue"></param>
        protected void TestSearchOperation(string inputValue, int expectedCount, string expectedClaimValue)
        {
            try
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                var entityTypes = new[] { "User", "SecGroup", "SharePointGroup", "System", "FormsRole" };

                SPProviderHierarchyTree providerResults = UnitTestsHelper.ClaimsProvider.Search(UnitTestsHelper.TestSiteCollUri, entityTypes, inputValue, null, 30);
                List<PickerEntity> entities = new List<PickerEntity>();
                foreach (var children in providerResults.Children)
                {
                    entities.AddRange(children.EntityData);
                }
                VerifySearchTest(entities, inputValue, expectedCount, expectedClaimValue);

                entities = UnitTestsHelper.ClaimsProvider.Resolve(UnitTestsHelper.TestSiteCollUri, entityTypes, inputValue).ToList();
                VerifySearchTest(entities, inputValue, expectedCount, expectedClaimValue);
                timer.Stop();
                Trace.TraceInformation($"{DateTime.Now:s} TestSearchOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: inputValue: '{inputValue}', expectedCount: '{expectedCount}', expectedClaimValue: '{expectedClaimValue}'.");
            }
            catch (Exception ex)
            {
                Trace.TraceError($"{DateTime.Now:s} TestSearchOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: inputValue: '{inputValue}', expectedCount: '{expectedCount}', expectedClaimValue: '{expectedClaimValue}'.");
            }
        }

        private void VerifySearchTest(List<PickerEntity> entities, string input, int expectedCount, string expectedClaimValue)
        {
            bool entityValueFound = false;
            StringBuilder detailedLog = new StringBuilder($"It returned {entities.Count} entities: ");
            string entityLogPattern = "entity \"{0}\", claim type: \"{1}\"; ";
            foreach (PickerEntity entity in entities)
            {
                detailedLog.AppendLine(String.Format(entityLogPattern, entity.Claim.Value, entity.Claim.ClaimType));
                if (String.Equals(expectedClaimValue, entity.Claim.Value, StringComparison.InvariantCultureIgnoreCase))
                {
                    entityValueFound = true;
                }
            }

            if (!String.IsNullOrWhiteSpace(expectedClaimValue) && !entityValueFound && expectedCount > 0)
            {
                Assert.Fail($"Input \"{input}\" returned no entity with claim value \"{expectedClaimValue}\". {detailedLog}");
            }

            if (expectedCount == Int32.MaxValue)
            {
                expectedCount = entities.Count;
            }

            Assert.That(entities.Count, Is.EqualTo(expectedCount), $"Input \"{input}\" should have returned {expectedCount} entities, but it returned {entities.Count} instead. {detailedLog}");
        }

        protected void TestValidationOperation(string claimType, string claimValue, bool shouldValidate)
        {
            SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            TestValidationOperation(inputClaim, shouldValidate, claimValue);
        }

        protected void TestValidationOperation(SPClaim inputClaim, bool shouldValidate, string expectedClaimValue)
        {
            try
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                var entityTypes = new[] { "User" };

                PickerEntity[] entities = UnitTestsHelper.ClaimsProvider.Resolve(UnitTestsHelper.TestSiteCollUri, entityTypes, inputClaim);

                int expectedCount = shouldValidate ? 1 : 0;
                Assert.That(entities.Length, Is.EqualTo(expectedCount), $"Validation of entity \"{inputClaim.Value}\" should have returned {expectedCount} entity, but it returned {entities.Length} instead.");
                if (shouldValidate)
                {
                    Assert.That(entities[0].Claim.Value, Is.EqualTo(expectedClaimValue).IgnoreCase, $"Validation of entity \"{inputClaim.Value}\" should have returned value \"{expectedClaimValue}\", but it returned \"{entities[0].Claim.Value}\" instead.");
                }
                timer.Stop();
                Trace.TraceInformation($"{DateTime.Now:s} TestValidationOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: inputClaim.Value: '{inputClaim.Value}', shouldValidate: '{shouldValidate}', expectedClaimValue: '{expectedClaimValue}'.");
            }
            catch (Exception ex)
            {
                Trace.TraceError($"{DateTime.Now:s} TestValidationOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: inputClaim.Value: '{inputClaim.Value}', shouldValidate: '{shouldValidate}', expectedClaimValue: '{expectedClaimValue}'.");
            }
        }

        /// <summary>
        /// Tests if the augmentation works as expected. By default this test is executed in every scenario.
        /// </summary>
        /// <param name="claimValue"></param>
        /// <param name="isMemberOfTrustedGroup"></param>
        protected void TestAugmentationOperation(string claimValue, bool isMemberOfTrustedGroup, string groupNameToTestInGroupMembership)
        {
            string claimType = UserIdentifierClaimType;
            SPClaim groupClaimToTestInGroupMembership = new SPClaim(GroupIdentifierClaimType, groupNameToTestInGroupMembership, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            try
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                Uri context = new Uri(UnitTestsHelper.TestSiteCollUri.AbsoluteUri);

                SPClaim[] groups = UnitTestsHelper.ClaimsProvider.GetClaimsForEntity(context, inputClaim);

                bool groupFound = false;
                if (groups != null && groups.Contains(groupClaimToTestInGroupMembership))
                {
                    groupFound = true;
                }

                if (isMemberOfTrustedGroup)
                {
                    Assert.That(groupFound, Is.True, $"Entity \"{claimValue}\" should be member of group \"{groupNameToTestInGroupMembership}\", but this group was not found in the claims returned by the claims provider.");
                }
                else
                {
                    Assert.That(groupFound, Is.False, $"Entity \"{claimValue}\" should NOT be member of group \"{groupNameToTestInGroupMembership}\", but this group was found in the claims returned by the claims provider.");
                }
                timer.Stop();
                Trace.TraceInformation($"{DateTime.Now:s} TestAugmentationOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: claimType: '{claimType}', claimValue: '{claimValue}', isMemberOfTrustedGroup: '{isMemberOfTrustedGroup}'.");
            }
            catch (Exception ex)
            {
                Trace.TraceError($"{DateTime.Now:s} TestAugmentationOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: claimType: '{claimType}', claimValue: '{claimValue}', isMemberOfTrustedGroup: '{isMemberOfTrustedGroup}'.");
            }
        }
    }
}
