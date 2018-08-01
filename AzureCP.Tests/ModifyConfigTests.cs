using azurecp;
using NUnit.Framework;
using System;
using System.Linq;

namespace AzureCP.Tests
{
    [TestFixture]
    public class ModifyConfigTests
    {
        private AzureCPConfig Config;

        [OneTimeSetUp]
        public void Init()
        {
            AzureCPConfig configFromConfigDB = AzureCPConfig.GetConfiguration(UnitTestsHelper.ClaimsProviderConfigName);
            // Create a local copy, otherwise changes will impact the whole process (even without calling Update method)
            Config = configFromConfigDB.CopyPersistedProperties();
            // Reset configuration to test its default for the tests
            Config.ResetCurrentConfiguration();
        }

        [Test]
        public void AddClaimTypeConfig()
        {
            ClaimTypeConfig ctConfig = new ClaimTypeConfig();

            // Adding a ClaimTypeConfig with a claim type already set should fail
            ctConfig.ClaimType = UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType;
            ctConfig.DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Adding a ClaimTypeConfig with UseMainClaimTypeOfDirectoryObject = false (default value) and DirectoryObjectProperty not set should fail
            ctConfig.ClaimType = UnitTestsHelper.RandomClaimType;
            ctConfig.DirectoryObjectProperty = AzureADObjectProperty.NotSet;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Adding a ClaimTypeConfig with UseMainClaimTypeOfDirectoryObject = true and ClaimType set should fail
            ctConfig.ClaimType = UnitTestsHelper.RandomClaimType;
            ctConfig.DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty;
            ctConfig.UseMainClaimTypeOfDirectoryObject = true;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Adding a ClaimTypeConfig with EntityType 'Group' should fail since 1 already exists by default and AzureCP allows only 1 claim type for EntityType 'Group'
            ctConfig.ClaimType = UnitTestsHelper.RandomClaimType;
            ctConfig.DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty;
            ctConfig.EntityType = DirectoryObjectType.Group;
            ctConfig.UseMainClaimTypeOfDirectoryObject = false;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Adding a valid ClaimTypeConfig should succeed
            ctConfig.ClaimType = UnitTestsHelper.RandomClaimType;
            ctConfig.DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty;
            ctConfig.EntityType = DirectoryObjectType.User;
            ctConfig.UseMainClaimTypeOfDirectoryObject = false;
            Assert.DoesNotThrow(() => Config.ClaimTypes.Add(ctConfig));

            // Adding a ClaimTypeConfig twice should fail
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Deleting the ClaimTypeConfig should succeed
            Assert.IsTrue(Config.ClaimTypes.Remove(ctConfig));
        }

        [Test]
        public void ModifyOrDeleteIdentityClaimTypeConfig()
        {
            // Deleting identity claim type from its claim type should fail
            string identityClaimType = UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Remove(identityClaimType));

            // Deleting identity claim type from its ClaimTypeConfig should fail
            ClaimTypeConfig identityCTConfig = Config.ClaimTypes.FirstOrDefault(x => String.Equals(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            Assert.IsNotNull(identityCTConfig);
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Remove(identityCTConfig));

            // Modify identity ClaimTypeConfig to set its EntityType to Group should fail
            identityCTConfig.EntityType = DirectoryObjectType.Group;
            Assert.Throws<InvalidOperationException>(() => Config.Update());
        }

        [Test]
        public void DuplicateClaimType()
        {
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));

            // Setting a duplicate claim type on a new item should fail
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = firstCTConfig.ClaimType, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation };
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Setting a duplicate claim type on items already existing in the list should fail
            var anotherCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType) && !String.Equals(firstCTConfig.ClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            anotherCTConfig.ClaimType = firstCTConfig.ClaimType;
            Assert.Throws<InvalidOperationException>(() => Config.Update());
        }

        [Test]
        public void DuplicatePrefixToBypassLookup()
        {
            string prefixToBypassLookup = "test:";

            // Setting a duplicate PrefixToBypassLookup on 2 items already existing in the list should fail
            Config.ClaimTypes.Where(x => !String.IsNullOrEmpty(x.ClaimType)).Take(2).Select(x => x.PrefixToBypassLookup = prefixToBypassLookup).ToList();
            Assert.Throws<InvalidOperationException>(() => Config.Update());

            // Setting a PrefixToBypassLookup on an existing item and add a new item with the same PrefixToBypassLookup should fail
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));
            firstCTConfig.PrefixToBypassLookup = prefixToBypassLookup;
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = UnitTestsHelper.RandomClaimType, PrefixToBypassLookup = prefixToBypassLookup, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation };
            Assert.Throws<InvalidOperationException>(() => Config.Update());
        }

        [Test]
        public void DuplicateEntityDataKey()
        {
            string entityDataKey = "test";

            // Setting a duplicate EntityDataKey on 2 items already existing in the list should fail
            Config.ClaimTypes.Where(x => !String.IsNullOrEmpty(x.ClaimType)).Take(2).Select(x => x.EntityDataKey = entityDataKey).ToList();
            Assert.Throws<InvalidOperationException>(() => Config.Update());

            // Setting a EntityDataKey on an existing item and add a new item with the same EntityDataKey should fail
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));
            firstCTConfig.EntityDataKey = entityDataKey;
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = UnitTestsHelper.RandomClaimType, EntityDataKey = entityDataKey, DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty };
            Assert.Throws<InvalidOperationException>(() => Config.Update());
        }

        [Test]
        public void DuplicateDirectoryObjectProperty()
        {
            ClaimTypeConfig existingCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType) && x.EntityType == DirectoryObjectType.User);

            // Create a new ClaimTypeConfig with a DirectoryObjectProperty already set should fail
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = UnitTestsHelper.RandomClaimType, EntityType = DirectoryObjectType.User, DirectoryObjectProperty = existingCTConfig.DirectoryObjectProperty };
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig));

            // Should be added successfully (for next test)
            ctConfig.DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty;
            Assert.DoesNotThrow(() => Config.ClaimTypes.Add(ctConfig));

            // Update an existing ClaimTypeConfig with a DirectoryObjectProperty already set should fail
            ctConfig.DirectoryObjectProperty = existingCTConfig.DirectoryObjectProperty;
            Assert.Throws<InvalidOperationException>(() => Config.Update());
        }
    }
}
