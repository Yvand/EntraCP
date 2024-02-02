using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    public class ClaimsProviderTestsBase
    {
        /// <summary>
        /// Configures whether to run the entity augmentation tests.
        /// </summary>
        public virtual bool DoAugmentationTest => true;

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
                return Settings.ClaimTypes.GetMainConfigurationForDirectoryObjectType(DirectoryObjectType.User).ClaimType;
            }
        }

        public string GroupIdentifierClaimType
        {
            get
            {
                return Settings.ClaimTypes.GetMainConfigurationForDirectoryObjectType(DirectoryObjectType.Group).ClaimType;
            }
        }

        protected EntraIDProviderSettings Settings = new EntraIDProviderSettings();

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

            string json = File.ReadAllText(UnitTestsHelper.AzureTenantsJsonFile);
            List<EntraIDTenant> azureTenants = JsonConvert.DeserializeObject<List<EntraIDTenant>>(json);
            Settings.EntraIDTenants = azureTenants;
            foreach (EntraIDTenant tenant in azureTenants)
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

        public void ProcessAndTestValidateEntityData(ValidateEntityData registrationData)
        {
            bool shouldValidate = registrationData.ShouldValidate;
            if (ExcludeGuestUsers && registrationData.UserType == ResultUserType.Guest)
            {
                shouldValidate = false;
            }
            if (ExcludeMemberUsers && registrationData.UserType == ResultUserType.Member)
            {
                shouldValidate = false;
            }

            string claimType = registrationData.EntityType == ResultEntityType.User ?
                UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType :
                UnitTestsHelper.TrustedGroupToAdd_ClaimType;

            SPClaim inputClaim = new SPClaim(claimType, registrationData.ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            TestValidationOperation(inputClaim, shouldValidate, registrationData.ClaimValue);
        }

        public void ProcessAndTestSearchEntityData(SearchEntityData registrationData)
        {
            // If current entry does not return only users AND either guests or members are excluded, ExpectedResultCount cannot be determined so test cannot run
            if (registrationData.SearchResultEntityTypes != ResultEntityType.User &&
                (ExcludeGuestUsers || ExcludeMemberUsers))
            {
                return;
            }

            if (ExcludeGuestUsers && registrationData.SearchResultUserTypes == ResultUserType.Guest)
            {
                registrationData.SearchResultCount = 0;
            }
            if (ExcludeMemberUsers && registrationData.SearchResultUserTypes == ResultUserType.Member)
            {
                registrationData.SearchResultCount = 0;
            }

            if (Settings.FilterExactMatchOnly == true)
            {
                registrationData.SearchResultCount = registrationData.ExactMatch ? 1 : 0;
            }

            TestSearchOperation(registrationData.Input, registrationData.SearchResultCount, registrationData.SearchResultSingleEntityClaimValue);
        }

        /// <summary>
        /// Start search operation on a specific claims provider
        /// </summary>
        /// <param name="inputValue"></param>
        /// <param name="expectedCount">How many entities are expected to be returned. Set to Int32.MaxValue if exact number is unknown but greater than 0</param>
        /// <param name="expectedClaimValue"></param>
        public void TestSearchOperation(string inputValue, int expectedCount, string expectedClaimValue)
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

        public void TestValidationOperation(string claimType, string claimValue, bool shouldValidate)
        {
            SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            TestValidationOperation(inputClaim, shouldValidate, claimValue);
        }

        public void TestValidationOperation(SPClaim inputClaim, bool shouldValidate, string expectedClaimValue)
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
        /// <param name="registrationData"></param>
        [Test, TestCaseSource(typeof(ValidateEntityDataSource), nameof(ValidateEntityDataSource.GetTestData), new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public virtual void TestAugmentationOperation(ValidateEntityData registrationData)
        {
            TestAugmentationOperation(registrationData.ClaimValue, registrationData.IsMemberOfTrustedGroup);
        }

        /// <summary>
        /// Tests if the augmentation works as expected. By default this test is executed in every scenario.
        /// </summary>
        /// <param name="claimValue"></param>
        /// <param name="isMemberOfTrustedGroup"></param>
        [TestCase("fake@FAKE.onmicrosoft.com", false)]
        public virtual void TestAugmentationOperation(string claimValue, bool isMemberOfTrustedGroup)
        {
            if (!DoAugmentationTest) { return; }
            string claimType = UserIdentifierClaimType;
            try
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
                Uri context = new Uri(UnitTestsHelper.TestSiteCollUri.AbsoluteUri);

                SPClaim[] groups = UnitTestsHelper.ClaimsProvider.GetClaimsForEntity(context, inputClaim);

                bool groupFound = false;
                if (groups != null && groups.Contains(UnitTestsHelper.TrustedGroup))
                {
                    groupFound = true;
                }

                if (isMemberOfTrustedGroup)
                {
                    Assert.That(groupFound, Is.True, $"Entity \"{claimValue}\" should be member of group \"{UnitTestsHelper.TrustedGroupToAdd_ClaimValue}\", but this group was not found in the claims returned by the claims provider.");
                }
                else
                {
                    Assert.That(groupFound, Is.False, $"Entity \"{claimValue}\" should NOT be member of group \"{UnitTestsHelper.TrustedGroupToAdd_ClaimValue}\", but this group was found in the claims returned by the claims provider.");
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
