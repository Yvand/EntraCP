using azurecp;
using NUnit.Framework;
using System;
using System.Linq;

namespace AzureCP.Tests
{
    [TestFixture]
    public class ModifyConfigTests
    {
        public const string ClaimsProviderConfigName = "AzureCPConfig";
        public const string NonExistingClaimType = "http://schemas.yvand.com/ws/claims/random";

        private AzureCPConfig Config;

        [OneTimeSetUp]
        public void Init()
        {
            AzureCPConfig configFromConfigDB = AzureCPConfig.GetConfiguration(ClaimsProviderConfigName);
            // Create a local copy, otherwise changes will impact the whole process (even without calling Update method)
            Config = configFromConfigDB.CopyPersistedProperties();
            // Reset configuration to test its default for the tests
            Config.ResetCurrentConfiguration();
        }

        [Test]
        public void AddClaimTypeConfig()
        {
            ClaimTypeConfig ctConfig = new ClaimTypeConfig();

            // Identity claim type already exists and a claim type cannot be added twice
            ctConfig.ClaimType = UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Property DirectoryObjectProperty should be set
            ctConfig.ClaimType = NonExistingClaimType;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Property ClaimType should be empty if UseMainClaimTypeOfDirectoryObject is true
            ctConfig.UseMainClaimTypeOfDirectoryObject = true;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // AzureCP allows only 1 claim type for EntityType 'Group'
            ctConfig.EntityType = DirectoryObjectType.Group;
            ctConfig.UseMainClaimTypeOfDirectoryObject = false;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Valid ClaimTypeConfig
            ctConfig.UseMainClaimTypeOfDirectoryObject = false;
            ctConfig.EntityType = DirectoryObjectType.User;
            ctConfig.DirectoryObjectProperty = AzureADObjectProperty.UserPrincipalName;
            Config.ClaimTypes.Add(ctConfig);
        }

        [Test]
        public void DeleteIdentityClaimTypeConfig()
        {
            string identityClaimType = UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType;
            Assert.IsNotEmpty(identityClaimType);
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Remove(identityClaimType));

            ClaimTypeConfig identityCTConfig = Config.ClaimTypes.FirstOrDefault(x => String.Equals(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            Assert.IsNotNull(identityCTConfig);
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Remove(identityCTConfig));
        }

        [Test]
        public void DuplicateClaimType()
        {
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));

            // Set a duplicate claim type on a new item
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = firstCTConfig.ClaimType, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation };
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Set a duplicate claim type on items already existing in the list
            var anotherCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType) && !String.Equals(firstCTConfig.ClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            anotherCTConfig.ClaimType = firstCTConfig.ClaimType;
            Assert.Throws<InvalidOperationException>(() => Config.Update());
        }

        [Test]
        public void DuplicatePrefixToBypassLookup()
        {
            string prefixToBypassLookup = "test:";

            // Set a duplicate PrefixToBypassLookup on 2 items already existing in the list
            Config.ClaimTypes.Where(x => !String.IsNullOrEmpty(x.ClaimType)).Take(2).Select(x => x.PrefixToBypassLookup = prefixToBypassLookup).ToList();
            Assert.Throws<InvalidOperationException>(() => Config.Update());

            // Set a PrefixToBypassLookup on an existing item and add a new item with the same PrefixToBypassLookup
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));
            firstCTConfig.PrefixToBypassLookup = prefixToBypassLookup;
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = NonExistingClaimType, PrefixToBypassLookup = prefixToBypassLookup, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation };
            Assert.Throws<InvalidOperationException>(() => Config.Update());
        }

        [Test]
        public void DuplicateEntityDataKey()
        {
            string entityDataKey = "test";

            // Set a duplicate EntityDataKey on 2 items already existing in the list
            Config.ClaimTypes.Where(x => !String.IsNullOrEmpty(x.ClaimType)).Take(2).Select(x => x.EntityDataKey = entityDataKey).ToList();
            Assert.Throws<InvalidOperationException>(() => Config.Update());

            // Set a EntityDataKey on an existing item and add a new item with the same EntityDataKey
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));
            firstCTConfig.EntityDataKey = entityDataKey;
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = NonExistingClaimType, EntityDataKey = entityDataKey, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation };
            Assert.Throws<InvalidOperationException>(() => Config.Update());
        }
    }
}
