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
using System.Threading.Tasks;
using Yvand.ClaimsProviders.Configuration;
using Yvand.ClaimsProviders.Configuration.AzureAD;

namespace Yvand.ClaimsProviders.Tests
{
    public class NewEntityTestsBase
    {
        /// <summary>
        /// Configure whether to run entity search tests.
        /// </summary>
        public virtual bool TestSearch => true;

        /// <summary>
        /// Configure whether to run entity validation tests.
        /// </summary>
        public virtual bool TestValidation => true;

        /// <summary>
        /// Configure whether to run entity augmentation tests.
        /// </summary>
        public virtual bool TestAugmentation => true;

        /// <summary>
        /// Configure whether to exclude AAD Guest users from search and validation. This does not impact augmentation.
        /// </summary>
        public virtual bool ExcludeGuestUsers => false;

        /// <summary>
        /// Configure whether to exclude AAD Member users from search and validation. This does not impact augmentation.
        /// </summary>
        public virtual bool ExcludeMemberUsers => false;

        private static readonly AzureCPSE ClaimsProvider = new AzureCPSE("AzureCPSE");
        public static SPTrustedLoginProvider SPTrust => SPSecurityTokenServiceManager.Local.TrustedLoginProviders.FirstOrDefault(x => String.Equals(x.ClaimProviderName, "AzureCPSE", StringComparison.InvariantCultureIgnoreCase));
        public static Uri TestSiteCollUri;
        public static string TrustedGroupToAdd_ClaimType => TestContext.Parameters["TrustedGroupToAdd_ClaimType"];
        public static string TrustedGroupToAdd_ClaimValue => TestContext.Parameters["TrustedGroupToAdd_ClaimValue"];
        public static SPClaim TrustedGroup => new SPClaim(TrustedGroupToAdd_ClaimType, TrustedGroupToAdd_ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, SPTrust.Name));
        protected AzureADEntityProviderConfiguration Config;
        private static AzureADEntityProviderConfiguration BackupConfig;

        [OneTimeSetUp]
        public void Init()
        {
            Trace.TraceInformation($"{DateTime.Now.ToString("s")} Start backup of current AzureCP configuration");
            Config = AzureCPSE.GetConfiguration();
            if (Config != null && BackupConfig != null)
            {
                BackupConfig = Config.CopyConfiguration() as AzureADEntityProviderConfiguration;
            }
            InitializeConfiguration();
        }

        /// <summary>
        /// Initialize configuration
        /// </summary>
        public virtual void InitializeConfiguration()
        {
            Config = AzureCPSE.CreateConfiguration();
            Config.ProxyAddress = TestContext.Parameters["ProxyAddress"];

#if DEBUG
            Config.Timeout = 99999;
#endif

            string json = File.ReadAllText(UnitTestsHelper.AzureTenantsJsonFile);
            List<AzureTenant> azureTenants = JsonConvert.DeserializeObject<List<AzureTenant>>(json);
            Config.AzureTenants = azureTenants;
            foreach (AzureTenant tenant in azureTenants)
            {
                tenant.ExcludeMemberUsers = ExcludeMemberUsers;
                tenant.ExcludeGuestUsers = ExcludeGuestUsers;
            }
            Config.Update();
            Trace.TraceInformation($"{DateTime.Now.ToString("s")} Set {Config.AzureTenants.Count} Azure AD tenants to AzureCP configuration");
        }

        [OneTimeTearDown]
        public void Cleanup()
        {
            try
            {
                //Config.ApplyConfiguration(BackupConfig);
                //Config.Update();
                Config = BackupConfig.CopyConfiguration() as AzureADEntityProviderConfiguration;
                Config.Update();
                //AzureCPSE.SaveConfiguration(Config);
                Trace.TraceInformation($"{DateTime.Now.ToString("s")} Restored original settings of AzureCP configuration");
            }
            catch (Exception ex)
            {
                Trace.TraceError($"{DateTime.Now.ToString("s")} Unexpected error while restoring the original settings of AzureCP configuration: {ex.Message}");
            }
        }

        public virtual void SearchEntities(SearchEntityData registrationData)
        {
            if (!TestSearch)
            {
                return;
            }

            // If current entry does not return only users, cannot reliably test number of results returned if guest and/or members should be excluded
            if (!String.Equals(registrationData.ResultType, "User", StringComparison.InvariantCultureIgnoreCase) &&
                (ExcludeGuestUsers || ExcludeMemberUsers))
            {
                return;
            }

            int expectedResultCount = registrationData.ExpectedResultCount;
            if (ExcludeGuestUsers && String.Equals(registrationData.UserType, ClaimsProviderConstants.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
            {
                expectedResultCount = 0;
            }
            if (ExcludeMemberUsers && String.Equals(registrationData.UserType, ClaimsProviderConstants.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
            {
                expectedResultCount = 0;
            }

            TestSearchOperation(registrationData.Input, expectedResultCount, registrationData.ExpectedEntityClaimValue);
        }

        public virtual void SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            if (!TestSearch) { return; }

            TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        public virtual void ValidateClaim(ValidateEntityData registrationData)
        {
            if (!TestValidation) { return; }

            bool shouldValidate = registrationData.ShouldValidate;
            if (ExcludeGuestUsers && String.Equals(registrationData.UserType, ClaimsProviderConstants.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
            {
                shouldValidate = false;
            }
            if (ExcludeMemberUsers && String.Equals(registrationData.UserType, ClaimsProviderConstants.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
            {
                shouldValidate = false;
            }

            SPClaim inputClaim = new SPClaim(Config.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, Config.SPTrust.Name));
            TestValidationOperation(inputClaim, shouldValidate, registrationData.ClaimValue);
        }

        public virtual void ValidateClaim(string claimType, string claimValue, bool shouldValidate)
        {
            if (!TestValidation) { return; }

            SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, Config.SPTrust.Name));
            TestValidationOperation(inputClaim, shouldValidate, claimValue);
        }

        public virtual void AugmentEntity(ValidateEntityData registrationData)
        {
            if (!TestAugmentation) { return; }

            TestAugmentationOperation(Config.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, registrationData.IsMemberOfTrustedGroup);
        }

        public virtual void AugmentEntity(string claimValue, bool shouldHavePermissions)
        {
            if (!TestAugmentation) { return; }

            TestAugmentationOperation(Config.SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimValue, shouldHavePermissions);
        }

        /// <summary>
        /// Start search operation on a specific claims provider
        /// </summary>
        /// <param name="inputValue"></param>
        /// <param name="expectedCount">How many entities are expected to be returned. Set to Int32.MaxValue if exact number is unknown but greater than 0</param>
        /// <param name="expectedClaimValue"></param>
        public static void TestSearchOperation(string inputValue, int expectedCount, string expectedClaimValue)
        {
            try
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                var entityTypes = new[] { "User", "SecGroup", "SharePointGroup", "System", "FormsRole" };

                SPProviderHierarchyTree providerResults = ClaimsProvider.Search(TestSiteCollUri, entityTypes, inputValue, null, 30);
                List<PickerEntity> entities = new List<PickerEntity>();
                foreach (var children in providerResults.Children)
                {
                    entities.AddRange(children.EntityData);
                }
                VerifySearchTest(entities, inputValue, expectedCount, expectedClaimValue);

                entities = ClaimsProvider.Resolve(TestSiteCollUri, entityTypes, inputValue).ToList();
                VerifySearchTest(entities, inputValue, expectedCount, expectedClaimValue);
                timer.Stop();
                Trace.TraceInformation($"{DateTime.Now.ToString("s")} TestSearchOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: inputValue: '{inputValue}', expectedCount: '{expectedCount}', expectedClaimValue: '{expectedClaimValue}'.");
            }
            catch (Exception ex)
            {
                Trace.TraceError($"{DateTime.Now.ToString("s")} TestSearchOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: inputValue: '{inputValue}', expectedCount: '{expectedCount}', expectedClaimValue: '{expectedClaimValue}'.");
            }
        }

        public static void VerifySearchTest(List<PickerEntity> entities, string input, int expectedCount, string expectedClaimValue)
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
                Assert.Fail($"Input \"{input}\" returned no entity with claim value \"{expectedClaimValue}\". {detailedLog.ToString()}");
            }

            if (expectedCount == Int32.MaxValue)
            {
                expectedCount = entities.Count;
            }

            Assert.AreEqual(expectedCount, entities.Count, $"Input \"{input}\" should have returned {expectedCount} entities, but it returned {entities.Count} instead. {detailedLog.ToString()}");
        }

        public static void TestValidationOperation(SPClaim inputClaim, bool shouldValidate, string expectedClaimValue)
        {
            try
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                var entityTypes = new[] { "User" };

                PickerEntity[] entities = ClaimsProvider.Resolve(TestSiteCollUri, entityTypes, inputClaim);

                int expectedCount = shouldValidate ? 1 : 0;
                Assert.AreEqual(expectedCount, entities.Length, $"Validation of entity \"{inputClaim.Value}\" should have returned {expectedCount} entity, but it returned {entities.Length} instead.");
                if (shouldValidate)
                {
                    StringAssert.AreEqualIgnoringCase(expectedClaimValue, entities[0].Claim.Value, $"Validation of entity \"{inputClaim.Value}\" should have returned value \"{expectedClaimValue}\", but it returned \"{entities[0].Claim.Value}\" instead.");
                }
                timer.Stop();
                Trace.TraceInformation($"{DateTime.Now.ToString("s")} TestValidationOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: inputClaim.Value: '{inputClaim.Value}', shouldValidate: '{shouldValidate}', expectedClaimValue: '{expectedClaimValue}'.");
            }
            catch (Exception ex)
            {
                Trace.TraceError($"{DateTime.Now.ToString("s")} TestValidationOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: inputClaim.Value: '{inputClaim.Value}', shouldValidate: '{shouldValidate}', expectedClaimValue: '{expectedClaimValue}'.");
            }
        }

        public static void TestAugmentationOperation(string claimType, string claimValue, bool isMemberOfTrustedGroup)
        {
            try
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, SPTrust.Name));
                Uri context = new Uri(UnitTestsHelper.TestSiteCollUri.AbsoluteUri);

                SPClaim[] groups = ClaimsProvider.GetClaimsForEntity(context, inputClaim);

                bool groupFound = false;
                if (groups != null && groups.Contains(TrustedGroup))
                {
                    groupFound = true;
                }

                if (isMemberOfTrustedGroup)
                {
                    Assert.IsTrue(groupFound, $"Entity \"{claimValue}\" should be member of group \"{TrustedGroupToAdd_ClaimValue}\", but this group was not found in the claims returned by the claims provider.");
                }
                else
                {
                    Assert.IsFalse(groupFound, $"Entity \"{claimValue}\" should NOT be member of group \"{TrustedGroupToAdd_ClaimValue}\", but this group was found in the claims returned by the claims provider.");
                }
                timer.Stop();
                Trace.TraceInformation($"{DateTime.Now.ToString("s")} TestAugmentationOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: claimType: '{claimType}', claimValue: '{claimValue}', isMemberOfTrustedGroup: '{isMemberOfTrustedGroup}'.");
            }
            catch (Exception ex)
            {
                Trace.TraceError($"{DateTime.Now.ToString("s")} TestAugmentationOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: claimType: '{claimType}', claimValue: '{claimValue}', isMemberOfTrustedGroup: '{isMemberOfTrustedGroup}'.");
            }
        }
    }
}
