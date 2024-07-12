using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System.Security.Claims;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class WithSecurityEnabledGroupsOnlyTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.EnableAugmentation = true;
            Settings.FilterSecurityEnabledGroupsOnly = true;
            Settings.ClaimTypes.UpdateGroupIdentifier(DirectoryObjectProperty.Id);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetTestData), new object[] { true, UnitTestsHelper.MaxNumberOfGroupsToTest })]
        public void TestGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
        }
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class WithoutSecurityEnabledGroupsOnlyTests : ClaimsProviderTestsBase
    {
        public override void InitializeSettings()
        {
            base.InitializeSettings();
            Settings.EnableAugmentation = true;
            Settings.FilterSecurityEnabledGroupsOnly = false;
            Settings.ClaimTypes.UpdateGroupIdentifier(DirectoryObjectProperty.Id);
            base.ApplySettings();
        }

        [Test]
        public override void CheckSettingsTest()
        {
            base.CheckSettingsTest();
        }

        [Test, TestCaseSource(typeof(EntraIdTestGroupsSource), nameof(EntraIdTestGroupsSource.GetTestData), new object[] { true, UnitTestsHelper.MaxNumberOfGroupsToTest })]
        public void TestGroups(EntraIdTestGroup group)
        {
            TestSearchAndValidateForEntraIDGroup(group);
        }

        [Test]
        [Repeat(5)]
        public override void TestAugmentationOfGoldUsersAgainstRandomGroups()
        {
            base.TestAugmentationOfGoldUsersAgainstRandomGroups();
        }
    }
}
