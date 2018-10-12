using azurecp;
using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;

namespace AzureCP.Tests
{
    /// <summary>
    /// Test guest accounts when their identity claim is the Email
    /// </summary>
    [TestFixture]
    public class GuestAccountsEmailTests : ModifyConfigBase
    {
        public override void Init()
        {
            base.Init();
            
            // Extra initialization for current test class
            Config.ClaimTypes.UpdateIdentifierForGuestUsers(AzureADObjectProperty.Mail);
        }

        [Test, TestCaseSource(typeof(SearchEntityDataSource), "GetTestData", new object[] { UnitTestsHelper.DataFile_GuestAccountsEmail_Search })]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void SearchEntities(SearchEntityData registrationData)
        {
            UnitTestsHelper.TestSearchOperation(registrationData.Input, registrationData.ExpectedResultCount, registrationData.ExpectedEntityClaimValue);
        }

        [Test, TestCaseSource(typeof(ValidateEntityDataSource), "GetTestData", new object[] { UnitTestsHelper.DataFile_GuestAccountsEmail_Validate })]
        [MaxTime(UnitTestsHelper.MaxTime)]
        [Repeat(UnitTestsHelper.TestRepeatCount)]
        public void ValidateClaim(ValidateEntityData registrationData)
        {
            SPClaim inputClaim = new SPClaim(UnitTestsHelper.SPTrust.IdentityClaimTypeInformation.MappedClaimType, registrationData.ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            UnitTestsHelper.TestValidationOperation(inputClaim, registrationData.ShouldValidate, registrationData.ClaimValue);
        }

        [TestCase(@"guest", 0, "GUEST@contoso.com")]
        [TestCase(@"yvand@mic", 2, "yvand@microsoft.com")]
        public void DEBUG_SearchEntities(string inputValue, int expectedResultCount, string expectedEntityClaimValue)
        {
            UnitTestsHelper.TestSearchOperation(inputValue, expectedResultCount, expectedEntityClaimValue);
        }

        [TestCase("http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress", "GUEST@contoso.com", false)]
        public void DEBUG_ValidateClaim(string claimType, string claimValue, bool shouldValidate)
        {
            SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
            UnitTestsHelper.TestValidationOperation(inputClaim, shouldValidate, claimValue);
        }
    }
}
