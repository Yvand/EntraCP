using NUnit.Framework;
using System.Runtime.CompilerServices;

namespace Yvand.ClaimsProviders.Tests
{
    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class ExcludeAllUserAccountsTests : NewEntityTestsBase
    {
        public override bool ExcludeGuestUsers => true;
        public override bool ExcludeMemberUsers => true;

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void AugmentEntity(ValidateEntityData registrationData)
        {
            base.AugmentEntity(registrationData);
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void SearchEntities(SearchEntityData registrationData)
        {
            if (registrationData.ExpectedEntityType != ResultEntityType.User)
            {
                return; // Cannot run this test if Mixed results since ExpectedResultCount cannot be determined
            }

            switch (registrationData.ExpectedUserType)
            {
                case ResultUserType.Mixed:
                    registrationData.ExpectedResultCount = 0;
                    break;
                case ResultUserType.Member:
                    registrationData.ExpectedResultCount = 0;
                    break;
                case ResultUserType.Guest:
                    registrationData.ExpectedResultCount = 0;
                    break;
            }
            base.SearchEntities(registrationData);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [MaxTime(UnitTestsHelper.MaxTime)]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void ValidateClaim(ValidateEntityData registrationData)
        {
            base.ValidateClaim(registrationData);
        }
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class ExcludeGuestUserAccountsTests : NewEntityTestsBase
    {
        public override bool ExcludeGuestUsers => true;
        public override bool ExcludeMemberUsers => false;
    }

    [TestFixture]
    [Parallelizable(ParallelScope.Children)]
    public class ExcludeMemberUserAccountsTests : NewEntityTestsBase
    {
        public override bool ExcludeGuestUsers => false;
        public override bool ExcludeMemberUsers => true;

        [Test, TestCaseSource(typeof(SearchEntityDataSource), "GetTestData", new object[] { EntityDataSourceType.AllAccounts })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public override void SearchEntities(SearchEntityData registrationData)
        {
            if (registrationData.ExpectedEntityType != ResultEntityType.User)
            {
                return; // Cannot run this test if Mixed results since ExpectedResultCount cannot be determined
            }

            switch (registrationData.ExpectedUserType)
            {
                case ResultUserType.Mixed:
                    return; // Cannot run this test if Mixed results since ExpectedResultCount cannot be determined
                case ResultUserType.Member:
                    registrationData.ExpectedResultCount = 0;
                    break;
            }
            base.SearchEntities(registrationData);
        }
    }
}
