using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System;
using System.Security.Claims;

namespace AzureCP.Tests
{
    [TestFixture]
    //[Parallelizable(ParallelScope.Children)]
    public class EntityTestsBase : BackupCurrentConfig
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

        public override void InitializeConfiguration()
        {
            base.InitializeConfiguration();
            Config.EnableAugmentation = true;
            foreach (var tenant in Config.AzureTenants)
            {
                tenant.ExcludeGuests = ExcludeGuestUsers;
                tenant.ExcludeMembers = ExcludeMemberUsers;
            }
            Config.Update();
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
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
            if (ExcludeGuestUsers && String.Equals(registrationData.UserType, UnitTestsHelper.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
            {
                expectedResultCount = 0;
            }
            if (ExcludeMemberUsers && String.Equals(registrationData.UserType, UnitTestsHelper.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
            {
                expectedResultCount = 0;
            }

            UnitTestsHelper.TestSearchOperation(registrationData.Input, expectedResultCount, registrationData.ExpectedEntityClaimValue);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [MaxTime(UnitTestsHelper.MaxTime)]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public virtual void ValidateClaim(ValidateEntityData registrationData)
        {
            if (!TestValidation) { return; }

            bool shouldValidate = registrationData.ShouldValidate;
            if (ExcludeGuestUsers && String.Equals(registrationData.UserType, UnitTestsHelper.GUEST_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
            {
                shouldValidate = false;
            }
            if (ExcludeMemberUsers && String.Equals(registrationData.UserType, UnitTestsHelper.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase))
            {
                shouldValidate = false;
            }

            SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            UnitTestsHelper.TestValidationOperation(inputClaim, shouldValidate, registrationData.ClaimValue);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public virtual void AugmentEntity(ValidateEntityData registrationData)
        {
            if (!TestAugmentation) { return; }

            UnitTestsHelper.TestAugmentationOperation(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, registrationData.IsMemberOfTrustedGroup);
        }

#if DEBUG
        //[TestCaseSource(typeof(SearchEntityDataSourceCollection))]
        public void DEBUG_SearchEntitiesFromCollection(string inputValue, string expectedCount, string expectedClaimValue)
        {
            if (!TestSearch) { return; }

            UnitTestsHelper.TestSearchOperation(inputValue, Convert.ToInt32(expectedCount), expectedClaimValue);
        }

        [TestCase(@"AADGroup1", 1, "5b0f6c56-c87f-44c3-9354-56cba03da433")]
        [TestCase(@"xyzguest", 0, "xyzGUEST@contoso.com")]
        [TestCase(@"AzureGr}", 1, "141cfd15-3941-4cbc-859f-d7125938fb72")]
        public void DEBUG_SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            if (!TestSearch) { return; }

            UnitTestsHelper.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        //[TestCase("http://schemas.microsoft.com/ws/2008/06/identity/claims/role", "5b0f6c56-c87f-44c3-9354-56cba03da433", true)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn", "FakeGuest@contoso.com", false)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress", "FakeGuest.com#EXT#@XXX.onmicrosoft.com", false)]
        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn", "FakeGuest.com#EXT#@XXX.onmicrosoft.com", false)]
        public void DEBUG_ValidateClaim(string claimType, string claimValue, bool shouldValidate)
        {
            if (!TestValidation) { return; }

            SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            UnitTestsHelper.TestValidationOperation(inputClaim, shouldValidate, claimValue);
        }

        [TestCase("xydGUEST@FAKE.onmicrosoft.com", false)]
        public void DEBUG_AugmentEntity(string claimValue, bool shouldHavePermissions)
        {
            if (!TestAugmentation) { return; }

            UnitTestsHelper.TestAugmentationOperation(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimValue, shouldHavePermissions);
        }
#endif
    }
}
