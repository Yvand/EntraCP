using NUnit.Framework;

namespace AzureCP.Tests
{
    [TestFixture]
    public class ExcludeAllUserAccountsTests : UserAccountsTestsBase
    {
        public override bool ExcludeGuestUsers => false;
        public override bool ExcludeMemberUsers => false;
    }
}
