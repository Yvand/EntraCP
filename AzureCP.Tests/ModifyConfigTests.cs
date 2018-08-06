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
        const string ConfigUpdateErrorMessage = "Some changes made to list ClaimTypes are invalid and cannot be committed to configuration database. Inspect inner exception for more details about the error.";

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
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig), $"Adding a ClaimTypeConfig with a claim type already set should throw exception InvalidOperationException with this message: \"Claim type '{UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType}' already exists in the collection\"");

            // Adding a ClaimTypeConfig with UseMainClaimTypeOfDirectoryObject = false (default value) and DirectoryObjectProperty not set should fail
            ctConfig.ClaimType = UnitTestsHelper.RandomClaimType;
            ctConfig.DirectoryObjectProperty = AzureADObjectProperty.NotSet;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig), $"Adding a ClaimTypeConfig with a claim type already set should throw exception InvalidOperationException with this message: \"Property DirectoryObjectProperty is required\"");

            // Adding a ClaimTypeConfig with UseMainClaimTypeOfDirectoryObject = true and ClaimType set should fail
            ctConfig.ClaimType = UnitTestsHelper.RandomClaimType;
            ctConfig.DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty;
            ctConfig.UseMainClaimTypeOfDirectoryObject = true;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig), $"Adding a ClaimTypeConfig with a claim type already set should throw exception InvalidOperationException with this message: \"No claim type should be set if UseMainClaimTypeOfDirectoryObject is set to true\"");

            // Adding a ClaimTypeConfig with EntityType 'Group' should fail since 1 already exists by default and AzureCP allows only 1 claim type for EntityType 'Group'
            ctConfig.ClaimType = UnitTestsHelper.RandomClaimType;
            ctConfig.DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty;
            ctConfig.EntityType = DirectoryObjectType.Group;
            ctConfig.UseMainClaimTypeOfDirectoryObject = false;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig), $"Adding a ClaimTypeConfig with a claim type already set should throw exception InvalidOperationException with this message: \"A claim type for EntityType 'Group' already exists in the collection\"");

            // Adding a valid ClaimTypeConfig should succeed
            ctConfig.ClaimType = UnitTestsHelper.RandomClaimType;
            ctConfig.DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty;
            ctConfig.EntityType = DirectoryObjectType.User;
            ctConfig.UseMainClaimTypeOfDirectoryObject = false;
            Assert.DoesNotThrow(() => Config.ClaimTypes.Add(ctConfig), $"ClaimTypeConfig {UnitTestsHelper.RandomClaimType} should have been added successfully but was not");

            // Adding a ClaimTypeConfig twice should fail
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig), $"Adding a ClaimTypeConfig with a claim type already set should throw exception InvalidOperationException with this message: \"Claim type '{UnitTestsHelper.RandomClaimType}' already exists in the collection\"");

            // Deleting the ClaimTypeConfig should succeed
            Assert.IsTrue(Config.ClaimTypes.Remove(ctConfig), $"ClaimTypeConfig {UnitTestsHelper.RandomClaimType} should have been removed successfully but was not");
        }

        [Test]
        public void ModifyOrDeleteIdentityClaimTypeConfig()
        {
            // Deleting identity claim type from ClaimTypes list based on its claim type should fail
            string identityClaimType = UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType;
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Remove(identityClaimType), $"Deleting identity claim type from ClaimTypes list should throw exception InvalidOperationException with this message: \"Cannot delete claim type \"{UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType}\" because it is the identity claim type of \"{UnitTestsHelper.SPTrust.Name}\"\"");

            // Deleting identity claim type from ClaimTypes list based on its ClaimTypeConfig should fail
            ClaimTypeConfig identityCTConfig = Config.ClaimTypes.FirstOrDefault(x => String.Equals(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            Assert.IsNotNull(identityCTConfig);
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Remove(identityClaimType), $"Deleting identity claim type from ClaimTypes list should throw exception InvalidOperationException with this message: \"Cannot delete claim type \"{UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType}\" because it is the identity claim type of \"{UnitTestsHelper.SPTrust.Name}\"\"");

            // Modify identity ClaimTypeConfig to set its EntityType to Group should fail
            identityCTConfig.EntityType = DirectoryObjectType.Group;
            Assert.Throws<InvalidOperationException>(() => Config.Update(), $"Modifying identity claim type to set its EntityType to Group should throw exception InvalidOperationException with this message: \"{ConfigUpdateErrorMessage}\"");
        }

        [Test]
        public void DuplicateClaimType()
        {
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));

            // Setting a duplicate claim type on a new item should fail
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = firstCTConfig.ClaimType, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation };
            Assert.Throws<InvalidOperationException>(() => Config.ClaimTypes.Add(ctConfig), $"Adding a ClaimTypeConfig with property ClaimType already defined in another ClaimTypeConfig should throw exception InvalidOperationException with this message: \"Claim type '{firstCTConfig.ClaimType}' already exists in the collection\"");

            // Setting a duplicate claim type on items already existing in the list should fail
            ClaimTypeConfig anotherCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType) && !String.Equals(firstCTConfig.ClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            anotherCTConfig.ClaimType = firstCTConfig.ClaimType;
            Assert.Throws<InvalidOperationException>(() => Config.Update(), $"Modifying an existing claim type to set a claim type already defined should throw exception InvalidOperationException with this message: \"{ConfigUpdateErrorMessage}\"");
        }

        [Test]
        public void DuplicatePrefixToBypassLookup()
        {
            string prefixToBypassLookup = "test:";

            // Setting a duplicate PrefixToBypassLookup on 2 items already existing in the list should fail
            Config.ClaimTypes.Where(x => !String.IsNullOrEmpty(x.ClaimType)).Take(2).Select(x => x.PrefixToBypassLookup = prefixToBypassLookup).ToList();
            Assert.Throws<InvalidOperationException>(() => Config.Update(), $"Setting a duplicate PrefixToBypassLookup on 2 items already existing in the list should throw exception InvalidOperationException with this message: \"{ConfigUpdateErrorMessage}\"");

            // Setting a PrefixToBypassLookup on an existing item and add a new item with the same PrefixToBypassLookup should fail
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));
            firstCTConfig.PrefixToBypassLookup = prefixToBypassLookup;
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = UnitTestsHelper.RandomClaimType, PrefixToBypassLookup = prefixToBypassLookup, DirectoryObjectProperty = AzureADObjectProperty.OfficeLocation };
            Assert.Throws<InvalidOperationException>(() => Config.Update(), $"Setting a duplicate PrefixToBypassLookup on an existing item and add a new item with the same PrefixToBypassLookup should throw exception InvalidOperationException with this message: \"{ConfigUpdateErrorMessage}\"");
        }

        [Test]
        public void DuplicateEntityDataKey()
        {
            string entityDataKey = "test";

            // Setting a duplicate EntityDataKey on 2 items already existing in the list should fail
            Config.ClaimTypes.Where(x => x.EntityType == DirectoryObjectType.User).Take(2).Select(x => x.EntityDataKey = entityDataKey).ToList();
            Assert.Throws<InvalidOperationException>(() => Config.Update(), $"Setting a duplicate EntityDataKey on 2 items already existing in the list should throw exception InvalidOperationException with this message: \"{ConfigUpdateErrorMessage}\"");

            // Setting a EntityDataKey on an existing item and add a new item with the same EntityDataKey should fail
            var firstCTConfig = Config.ClaimTypes.FirstOrDefault(x => !String.IsNullOrEmpty(x.ClaimType));
            firstCTConfig.EntityDataKey = entityDataKey;
            ClaimTypeConfig ctConfig = new ClaimTypeConfig() { ClaimType = UnitTestsHelper.RandomClaimType, EntityDataKey = entityDataKey, DirectoryObjectProperty = UnitTestsHelper.RandomObjectProperty };
            Assert.Throws<InvalidOperationException>(() => Config.Update(), $"Setting a EntityDataKey on an existing item and add a new item with the same EntityDataKey should throw exception InvalidOperationException with this message: \"{ConfigUpdateErrorMessage}\"");
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
