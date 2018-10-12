using azurecp;
using DataAccess;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using NUnit.Framework;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;

[SetUpFixture]
public class UnitTestsHelper
{
    public static azurecp.AzureCP ClaimsProvider = new azurecp.AzureCP(UnitTestsHelper.ClaimsProviderName);
    public const string ClaimsProviderName = "AzureCP";
    public const string ClaimsProviderConfigName = "AzureCPConfig";
    public static Uri Context = new Uri("http://spsites/sites/AzureCP.UnitTests");
    public const int MaxTime = 50000;
    public const int TestRepeatCount = 100;
    public const string FarmAdmin = @"i:0#.w|contoso\yvand";

    public const string RandomClaimType = "http://schemas.yvand.com/ws/claims/random";
    public const string RandomClaimValue = "IDoNotExist";
    public const AzureADObjectProperty RandomObjectProperty = AzureADObjectProperty.AccountEnabled;

    public const string TrustedGroupToAdd_ClaimValue = "99abdc91-e6e0-475c-a0ba-5014f91de853";
    public const string TrustedGroupToAdd_ClaimType = "http://schemas.microsoft.com/ws/2008/06/identity/claims/role";
    public static SPClaim TrustedGroup = new SPClaim(TrustedGroupToAdd_ClaimType, TrustedGroupToAdd_ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, SPTrust.Name));

    public const string DataFile_MemberAccounts_Search = @"F:\Data\Dev\AzureCP_Tests_MemberAccounts_Search.csv";
    public const string DataFile_MemberAccounts_Validate = @"F:\Data\Dev\AzureCP_Tests_MemberAccounts_Validate.csv";
    public const string DataFile_GuestAccountsUPN_Search = @"F:\Data\Dev\AzureCP_Tests_GuestAccountsUPN_Search.csv";
    public const string DataFile_GuestAccountsUPN_Validate = @"F:\Data\Dev\AzureCP_Tests_GuestAccountsUPN_Validate.csv";
    public const string DataFile_GuestAccountsEmail_Search = @"F:\Data\Dev\AzureCP_Tests_GuestAccountsEmail_Search.csv";
    public const string DataFile_GuestAccountsEmail_Validate = @"F:\Data\Dev\AzureCP_Tests_GuestAccountsEmail_Validate.csv";

    public static SPTrustedLoginProvider SPTrust
    {
        get => SPSecurityTokenServiceManager.Local.TrustedLoginProviders.FirstOrDefault(x => String.Equals(x.ClaimProviderName, UnitTestsHelper.ClaimsProviderName, StringComparison.InvariantCultureIgnoreCase));
    }

    [OneTimeSetUp]
    public static void CheckSiteCollection()
    {
        //return; // Uncommented when debugging AzureCP code from unit tests

        AzureCPConfig config = AzureCPConfig.GetConfiguration(UnitTestsHelper.ClaimsProviderConfigName, UnitTestsHelper.SPTrust.Name);
        if (config == null)
        {
            AzureCPConfig.CreateConfiguration(ClaimsProviderConstants.CONFIG_ID, ClaimsProviderConstants.CONFIG_NAME, SPTrust.Name);
        }

        SPWebApplication wa = SPWebApplication.Lookup(Context);
        if (wa != null)
        {
            Console.WriteLine($"Web app {wa.Name} found.");
            SPClaimProviderManager claimMgr = SPClaimProviderManager.Local;
            string encodedClaim = claimMgr.EncodeClaim(TrustedGroup);
            SPUserInfo userInfo = new SPUserInfo { LoginName = encodedClaim, Name = TrustedGroupToAdd_ClaimValue };

            if (!SPSite.Exists(Context))
            {
                Console.WriteLine($"Creating site collection {Context.AbsoluteUri}...");
                SPSite spSite = wa.Sites.Add(Context.AbsoluteUri, ClaimsProviderName, ClaimsProviderName, 1033, "STS#0", FarmAdmin, String.Empty, String.Empty);
                spSite.RootWeb.CreateDefaultAssociatedGroups(FarmAdmin, FarmAdmin, spSite.RootWeb.Title);

                SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
                spSite.Dispose();
            }
            else
            {
                using (SPSite spSite = new SPSite(Context.AbsoluteUri))
                {
                    SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                    membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
                }
            }
        }
        else
        {
            Console.WriteLine($"Web app {Context} was NOT found.");
        }
    }

    public static void TestSearchOperation(string inputValue, int expectedCount, string expectedClaimValue)
    {
        string[] entityTypes = new string[] { "User", "SecGroup", "SharePointGroup", "System", "FormsRole" };

        SPProviderHierarchyTree providerResults = ClaimsProvider.Search(Context, entityTypes, inputValue, null, 30);
        List<PickerEntity> entities = new List<PickerEntity>();
        foreach (var children in providerResults.Children)
        {
            entities.AddRange(children.EntityData);
        }
        VerifySearchTest(entities, inputValue, expectedCount, expectedClaimValue);

        entities = ClaimsProvider.Resolve(Context, entityTypes, inputValue).ToList();
        VerifySearchTest(entities, inputValue, expectedCount, expectedClaimValue);
    }

    public static void VerifySearchTest(List<PickerEntity> entities, string input, int expectedCount, string expectedClaimValue)
    {
        bool entityValueFound = false;
        foreach (PickerEntity entity in entities)
        {
            if (String.Equals(expectedClaimValue, entity.Claim.Value, StringComparison.InvariantCultureIgnoreCase))
            {
                entityValueFound = true;
            }
        }

        if (!entityValueFound && expectedCount > 0)
        {
            Assert.Fail($"Input \"{input}\" returned no entity with claim value \"{expectedClaimValue}\".");
        }
        Assert.AreEqual(expectedCount, entities.Count, $"Input \"{input}\" should have returned {expectedCount} entities, but it returned {entities.Count} instead.");
    }

    public static void TestValidationOperation(SPClaim inputClaim, bool shouldValidate, string expectedClaimValue)
    {
        string[] entityTypes = new string[] { "User" };

        PickerEntity[] entities = ClaimsProvider.Resolve(Context, entityTypes, inputClaim);

        int expectedCount = shouldValidate ? 1 : 0;
        Assert.AreEqual(expectedCount, entities.Length, $"Validation of entity \"{inputClaim.Value}\" should have returned {expectedCount} entity, but it returned {entities.Length} instead.");
        if (shouldValidate)
        {
            StringAssert.AreEqualIgnoringCase(expectedClaimValue, entities[0].Claim.Value, $"Validation of entity \"{inputClaim.Value}\" should have returned value \"{expectedClaimValue}\", but it returned \"{entities[0].Claim.Value}\" instead.");
        }
    }

    public static void TestAugmentationOperation(string claimType, string claimValue, bool isMemberOfTrustedGroup)
    {
        SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name));
        Uri context = new Uri(UnitTestsHelper.Context.AbsoluteUri);

        SPClaim[] groups = ClaimsProvider.GetClaimsForEntity(context, inputClaim);

        bool groupFound = false;
        if (groups != null && groups.Contains(TrustedGroup))
            groupFound = true;

        if (isMemberOfTrustedGroup)
            Assert.IsTrue(groupFound, $"Entity \"{claimValue}\" should be member of group \"{TrustedGroupToAdd_ClaimValue}\", but this group was not found in the claims returned by the claims provider.");
        else
            Assert.IsFalse(groupFound, $"Entity \"{claimValue}\" should NOT be member of group \"{TrustedGroupToAdd_ClaimValue}\", but this group was found in the claims returned by the claims provider.");
    }
}

public class SearchEntityDataSourceCollection : IEnumerable
{
    public IEnumerator GetEnumerator()
    {
        yield return new[] { "AADGroup1", "1", "5b0f6c56-c87f-44c3-9354-56cba03da433" };
        yield return new[] { "AADGroupTes", "1", "99abdc91-e6e0-475c-a0ba-5014f91de853" };
    }
}

public class SearchEntityDataSource
{
    public static IEnumerable<TestCaseData> GetTestData(string dataFile)
    {
        DataTable dt = DataTable.New.ReadCsv(dataFile);

        foreach (Row row in dt.Rows)
        {
            var registrationData = new SearchEntityData();
            registrationData.Input = row["Input"];
            registrationData.ExpectedResultCount = Convert.ToInt32(row["ExpectedResultCount"]);
            registrationData.ExpectedEntityClaimValue = row["ExpectedEntityClaimValue"];
            yield return new TestCaseData(new object[] { registrationData });
        }
    }

    //public class ReadCSV
    //{
    //    public void GetValue()
    //    {
    //        TextReader tr1 = new StreamReader(@"c:\pathtofile\filename", true);

    //        var Data = tr1.ReadToEnd().Split('\n')
    //        .Where(l => l.Length > 0)  //nonempty strings
    //        .Skip(1)               // skip header 
    //        .Select(s => s.Trim())   // delete whitespace
    //        .Select(l => l.Split(',')) // get arrays of values
    //        .Select(l => new { Field1 = l[0], Field2 = l[1], Field3 = l[2] });
    //    }
    //}
}

public class SearchEntityData
{
    public string Input;
    public int ExpectedResultCount;
    public string ExpectedEntityClaimValue;
}

public class ValidateEntityDataSource
{
    public static IEnumerable<TestCaseData> GetTestData(string dataFile)
    {
        DataTable dt = DataTable.New.ReadCsv(dataFile);
        foreach (Row row in dt.Rows)
        {
            var registrationData = new ValidateEntityData();
            registrationData.ClaimValue = row["ClaimValue"];
            registrationData.ShouldValidate = Convert.ToBoolean(row["ShouldValidate"]);
            registrationData.IsMemberOfTrustedGroup = Convert.ToBoolean(row["IsMemberOfTrustedGroup"]);
            yield return new TestCaseData(new object[] { registrationData });
        }
    }
}

public class ValidateEntityData
{
    public string ClaimValue;
    public bool ShouldValidate;
    public bool IsMemberOfTrustedGroup;
}
