using NUnit.Framework;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    public class WrongConfigBadClaimTypeTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            ClaimTypeConfig randomClaimTypeConfig = new ClaimTypeConfig
            {
                ClaimType = UnitTestsHelper.RandomClaimType,
                EntityProperty = UnitTestsHelper.RandomObjectProperty,
            };
            Settings.ClaimTypes = new ClaimTypeConfigCollection(UnitTestsHelper.SPTrust) { randomClaimTypeConfig };
            ConfigurationShouldBeValid = false;
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }
    }

    [TestFixture]
    public class WrongConfigInvalidGroupIdentifierTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            // Set GroupIdentifierConfig to use UserPrincipalName, which exists for User but not Group
            ClaimTypeConfig groupClaimTypeConfig = new ClaimTypeConfig
            {
                ClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType,
                EntityType = DirectoryObjectType.Group,
                EntityProperty = DirectoryObjectProperty.UserPrincipalName, // Invalid for Group - this property exists for User but not for Group
                EntityPropertyToUseAsDisplayText = DirectoryObjectProperty.DisplayName,
            };
            
            // Get the default user identity config and add the invalid group config
            ClaimTypeConfigCollection newClaimTypes = new ClaimTypeConfigCollection(UnitTestsHelper.SPTrust);
            
            // Add a minimal user identity config (required for the configuration to be somewhat valid)
            IdentityClaimTypeConfig userIdentityConfig = new IdentityClaimTypeConfig
            {
                EntityType = DirectoryObjectType.User,
                EntityProperty = DirectoryObjectProperty.UserPrincipalName,
                ClaimType = UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType
            };
            newClaimTypes.Add(userIdentityConfig);
            
            // Add the invalid group config
            newClaimTypes.Add(groupClaimTypeConfig);
            
            Settings.ClaimTypes = newClaimTypes;
            ConfigurationShouldBeValid = false;
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }
    }
}
