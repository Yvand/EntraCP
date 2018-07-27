using NUnit.Framework;

namespace AzureCP.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class AugmentationTests
    {
        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData")]
        [MaxTime(UnitTestsHelper.MaxTime)]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void AugmentEntity(ValidateEntityData registrationData)
        {
            UnitTestsHelper.DoAugmentationOperationAndVerifyResult(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, registrationData.IsMemberOfTrustedGroup);
        }

        [TestCase("i:05.t|contoso.local|xyd@FAKE.onmicrosoft.com", false)]
        public void DEBUG_AugmentEntity(string claimValue, bool shouldHavePermissions)
        {
            UnitTestsHelper.DoAugmentationOperationAndVerifyResult(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, claimValue, shouldHavePermissions);
        }
    }
}
