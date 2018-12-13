using NUnit.Framework;

namespace AzureCP.Tests
{
    [TestFixture]
    public class ExcludeGuestUserAccountsTests : EntityTestsBase
    {
        public override bool ExcludeGuestUsers => true;
        public override bool ExcludeMemberUsers => false;
    }
}
