using DataAccess;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Security.Claims;
using System.Text;
using Yvand.ClaimsProviders;
using Yvand.ClaimsProviders.Configuration;
using Yvand.ClaimsProviders.Configuration.AzureAD;

[SetUpFixture]
public class UnitTestsHelper
{
    public static readonly AzureCPSE ClaimsProvider = new AzureCPSE("AzureCPSE");
    public static string TestSiteRelativePath => $"/sites/{TestContext.Parameters["TestSiteCollectionName"]}";
    private static Uri TestSiteCollUri;
    public const int MaxTime = 50000;
    public static string FarmAdmin => TestContext.Parameters["FarmAdmin"];
#if DEBUG
    public const int TestRepeatCount = 1;
#else
    public const int TestRepeatCount = 20;
#endif

    public static string RandomClaimType => "http://schemas.yvand.net/ws/claims/random";
    public static string RandomClaimValue => "IDoNotExist";
    public static DirectoryObjectProperty RandomObjectProperty => DirectoryObjectProperty.AccountEnabled;

    public static string TrustedGroupToAdd_ClaimType => TestContext.Parameters["TrustedGroupToAdd_ClaimType"];
    public static string TrustedGroupToAdd_ClaimValue => TestContext.Parameters["TrustedGroupToAdd_ClaimValue"];
    public static SPClaim TrustedGroup => new SPClaim(TrustedGroupToAdd_ClaimType, TrustedGroupToAdd_ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, ClaimsProvider.EntityProvider.Configuration.SPTrust.Name));

    public static string AzureTenantsJsonFile => TestContext.Parameters["AzureTenantsJsonFile"];
    public static string DataFile_GuestAccountsUPN_Search => TestContext.Parameters["DataFile_GuestAccountsUPN_Search"];
    public static string DataFile_GuestAccountsUPN_Validate => TestContext.Parameters["DataFile_GuestAccountsUPN_Validate"];
    public static string DataFile_AllAccounts_Search => TestContext.Parameters["DataFile_AllAccounts_Search"];
    public static string DataFile_AllAccounts_Validate => TestContext.Parameters["DataFile_AllAccounts_Validate"];
    static TextWriterTraceListener logFileListener { get; set; }

    [OneTimeSetUp]
    public static void InitializeSiteCollection()
    {
        logFileListener = new TextWriterTraceListener(TestContext.Parameters["TestLogFileName"]);
        Trace.Listeners.Add(logFileListener);
        Trace.AutoFlush = true;
        Trace.TraceInformation($"{DateTime.Now.ToString("s")} Start integration tests of {AzureCPSE.ClaimsProviderName} {FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(AzureCPSE)).Location).FileVersion}.");
        Trace.WriteLine($"{DateTime.Now.ToString("s")} DataFile_AllAccounts_Search: {DataFile_AllAccounts_Search}");
        Trace.WriteLine($"{DateTime.Now.ToString("s")} DataFile_AllAccounts_Validate: {DataFile_AllAccounts_Validate}");
        Trace.WriteLine($"{DateTime.Now.ToString("s")} DataFile_GuestAccountsUPN_Search: {DataFile_GuestAccountsUPN_Search}");
        Trace.WriteLine($"{DateTime.Now.ToString("s")} DataFile_GuestAccountsUPN_Validate: {DataFile_GuestAccountsUPN_Validate}");
        Trace.WriteLine($"{DateTime.Now.ToString("s")} TestSiteCollectionName: {TestContext.Parameters["TestSiteCollectionName"]}");

        AzureADEntityProviderConfiguration config = AzureCPSE.GetConfiguration();
        if (config == null)
        {
            AzureCPSE.CreateConfiguration();
        }

        ClaimsProvider.ValidateLocalConfiguration(null);
        if (ClaimsProvider.EntityProvider.Configuration?.SPTrust == null)
        {
            Trace.TraceError($"{DateTime.Now.ToString("s")} SPTrust: is null");
        }
        else
        {
            Trace.WriteLine($"{DateTime.Now.ToString("s")} SPTrust: {ClaimsProvider.EntityProvider.Configuration.SPTrust.Name}");
        }

#if DEBUG
        //return; // Uncommented when debugging AzureCP code from unit tests
#endif

        var service = SPFarm.Local.Services.GetValue<SPWebService>(String.Empty);
        SPWebApplication wa = service.WebApplications.FirstOrDefault(x =>
        {
            foreach (var iisSetting in x.IisSettings.Values)
            {
                foreach (SPAuthenticationProvider authenticationProviders in iisSetting.ClaimsAuthenticationProviders)
                {
                    if (String.Equals(authenticationProviders.ClaimProviderName, AzureCPSE.ClaimsProviderName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            return false;
        });
        if (wa == null)
        {
            Trace.TraceError($"{DateTime.Now.ToString("s")} Web application was NOT found.");
            return;
        }

        Trace.WriteLine($"{DateTime.Now.ToString("s")} Web application {wa.Name} found.");
        Uri waRootAuthority = wa.AlternateUrls[0].Uri;
        TestSiteCollUri = new Uri($"{waRootAuthority.GetLeftPart(UriPartial.Authority)}{TestSiteRelativePath}");
        SPClaimProviderManager claimMgr = SPClaimProviderManager.Local;
        string encodedClaim = claimMgr.EncodeClaim(TrustedGroup);
        SPUserInfo userInfo = new SPUserInfo { LoginName = encodedClaim, Name = TrustedGroupToAdd_ClaimValue };

        // The root site may not exist, but it must be present for tests to run
        if (!SPSite.Exists(waRootAuthority))
        {
            Trace.WriteLine($"{DateTime.Now.ToString("s")} Creating root site collection {waRootAuthority.AbsoluteUri}...");
            SPSite spSite = wa.Sites.Add(waRootAuthority.AbsoluteUri, "root", "root", 1033, "STS#3", FarmAdmin, String.Empty, String.Empty);
            spSite.RootWeb.CreateDefaultAssociatedGroups(FarmAdmin, FarmAdmin, spSite.RootWeb.Title);

            SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
            membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
            spSite.Dispose();
        }

        if (!SPSite.Exists(TestSiteCollUri))
        {
            Trace.WriteLine($"{DateTime.Now.ToString("s")} Creating site collection {TestSiteCollUri.AbsoluteUri}...");
            SPSite spSite = wa.Sites.Add(TestSiteCollUri.AbsoluteUri, AzureCPSE.ClaimsProviderName, AzureCPSE.ClaimsProviderName, 1033, "STS#3", FarmAdmin, String.Empty, String.Empty);
            spSite.RootWeb.CreateDefaultAssociatedGroups(FarmAdmin, FarmAdmin, spSite.RootWeb.Title);

            SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
            membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
            spSite.Dispose();
        }
        else
        {
            using (SPSite spSite = new SPSite(TestSiteCollUri.AbsoluteUri))
            {
                SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
            }
        }
    }

    [OneTimeTearDown]
    public static void Cleanup()
    {
        Trace.WriteLine($"{DateTime.Now.ToString("s")} Integration tests of {AzureCPSE.ClaimsProviderName} {FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(AzureCPSE)).Location).FileVersion} finished.");
        Trace.Flush();
        if (logFileListener != null)
        {
            logFileListener.Dispose();
        }
    }

    /// <summary>
    /// Start search operation on a specific claims provider
    /// </summary>
    /// <param name="inputValue"></param>
    /// <param name="expectedCount">How many entities are expected to be returned. Set to Int32.MaxValue if exact number is unknown but greater than 0</param>
    /// <param name="expectedClaimValue"></param>
    public static void TestSearchOperation(string inputValue, int expectedCount, string expectedClaimValue)
    {
        try
        {
            Stopwatch timer = new Stopwatch();
            timer.Start();
            var entityTypes = new[] { "User", "SecGroup", "SharePointGroup", "System", "FormsRole" };

            SPProviderHierarchyTree providerResults = ClaimsProvider.Search(TestSiteCollUri, entityTypes, inputValue, null, 30);
            List<PickerEntity> entities = new List<PickerEntity>();
            foreach (var children in providerResults.Children)
            {
                entities.AddRange(children.EntityData);
            }
            VerifySearchTest(entities, inputValue, expectedCount, expectedClaimValue);

            entities = ClaimsProvider.Resolve(TestSiteCollUri, entityTypes, inputValue).ToList();
            VerifySearchTest(entities, inputValue, expectedCount, expectedClaimValue);
            timer.Stop();
            Trace.WriteLine($"{DateTime.Now.ToString("s")} TestSearchOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: inputValue: '{inputValue}', expectedCount: '{expectedCount}', expectedClaimValue: '{expectedClaimValue}'.");
        }
        catch (Exception ex)
        {
            Trace.TraceError($"{DateTime.Now.ToString("s")} TestSearchOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: inputValue: '{inputValue}', expectedCount: '{expectedCount}', expectedClaimValue: '{expectedClaimValue}'.");
        }
    }

    public static void VerifySearchTest(List<PickerEntity> entities, string input, int expectedCount, string expectedClaimValue)
    {
        bool entityValueFound = false;
        StringBuilder detailedLog = new StringBuilder($"It returned {entities.Count} entities: ");
        string entityLogPattern = "entity \"{0}\", claim type: \"{1}\"; ";
        foreach (PickerEntity entity in entities)
        {
            detailedLog.AppendLine(String.Format(entityLogPattern, entity.Claim.Value, entity.Claim.ClaimType));
            if (String.Equals(expectedClaimValue, entity.Claim.Value, StringComparison.InvariantCultureIgnoreCase))
            {
                entityValueFound = true;
            }
        }

        if (!entityValueFound && expectedCount > 0)
        {
            Assert.Fail($"Input \"{input}\" returned no entity with claim value \"{expectedClaimValue}\". {detailedLog.ToString()}");
        }

        if (expectedCount == Int32.MaxValue)
        {
            expectedCount = entities.Count;
        }

        Assert.AreEqual(expectedCount, entities.Count, $"Input \"{input}\" should have returned {expectedCount} entities, but it returned {entities.Count} instead. {detailedLog.ToString()}");
    }

    public static void TestValidationOperation(SPClaim inputClaim, bool shouldValidate, string expectedClaimValue)
    {
        try
        {
            Stopwatch timer = new Stopwatch();
            timer.Start();
            var entityTypes = new[] { "User" };

            PickerEntity[] entities = ClaimsProvider.Resolve(TestSiteCollUri, entityTypes, inputClaim);

            int expectedCount = shouldValidate ? 1 : 0;
            Assert.AreEqual(expectedCount, entities.Length, $"Validation of entity \"{inputClaim.Value}\" should have returned {expectedCount} entity, but it returned {entities.Length} instead.");
            if (shouldValidate)
            {
                StringAssert.AreEqualIgnoringCase(expectedClaimValue, entities[0].Claim.Value, $"Validation of entity \"{inputClaim.Value}\" should have returned value \"{expectedClaimValue}\", but it returned \"{entities[0].Claim.Value}\" instead.");
            }
            timer.Stop();
            Trace.WriteLine($"{DateTime.Now.ToString("s")} TestValidationOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: inputClaim.Value: '{inputClaim.Value}', shouldValidate: '{shouldValidate}', expectedClaimValue: '{expectedClaimValue}'.");
        }
        catch (Exception ex)
        {
            Trace.TraceError($"{DateTime.Now.ToString("s")} TestValidationOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: inputClaim.Value: '{inputClaim.Value}', shouldValidate: '{shouldValidate}', expectedClaimValue: '{expectedClaimValue}'.");
        }
    }

    public static void TestAugmentationOperation(string claimType, string claimValue, bool isMemberOfTrustedGroup)
    {
        try
        {
            Stopwatch timer = new Stopwatch();
            timer.Start();
            SPClaim inputClaim = new SPClaim(claimType, claimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, ClaimsProvider.EntityProvider.Configuration.SPTrust.Name));
            Uri context = new Uri(UnitTestsHelper.TestSiteCollUri.AbsoluteUri);

            SPClaim[] groups = ClaimsProvider.GetClaimsForEntity(context, inputClaim);

            bool groupFound = false;
            if (groups != null && groups.Contains(TrustedGroup))
            {
                groupFound = true;
            }

            if (isMemberOfTrustedGroup)
            {
                Assert.IsTrue(groupFound, $"Entity \"{claimValue}\" should be member of group \"{TrustedGroupToAdd_ClaimValue}\", but this group was not found in the claims returned by the claims provider.");
            }
            else
            {
                Assert.IsFalse(groupFound, $"Entity \"{claimValue}\" should NOT be member of group \"{TrustedGroupToAdd_ClaimValue}\", but this group was found in the claims returned by the claims provider.");
            }
            timer.Stop();
            Trace.WriteLine($"{DateTime.Now.ToString("s")} TestAugmentationOperation finished in {timer.ElapsedMilliseconds} ms. Parameters: claimType: '{claimType}', claimValue: '{claimValue}', isMemberOfTrustedGroup: '{isMemberOfTrustedGroup}'.");
        }
        catch (Exception ex)
        {
            Trace.TraceError($"{DateTime.Now.ToString("s")} TestAugmentationOperation failed with exception '{ex.GetType()}', message '{ex.Message}'. Parameters: claimType: '{claimType}', claimValue: '{claimValue}', isMemberOfTrustedGroup: '{isMemberOfTrustedGroup}'.");
        }
    }
}

//public class SearchEntityDataSourceCollection : IEnumerable
//{
//    public IEnumerator GetEnumerator()
//    {
//        yield return new[] { "AADGroup1", "1", "5b0f6c56-c87f-44c3-9354-56cba03da433" };
//        yield return new[] { "AADGroupTes", "1", "99abdc91-e6e0-475c-a0ba-5014f91de853" };
//    }
//}

public enum EntityDataSourceType
{
    AllAccounts,
    UPNB2BGuestAccounts
}

public class SearchEntityDataSource
{
    public static IEnumerable<TestCaseData> GetTestData(EntityDataSourceType entityDataSourceType)
    {
        string csvPath = UnitTestsHelper.DataFile_AllAccounts_Search;
        if (entityDataSourceType == EntityDataSourceType.UPNB2BGuestAccounts)
        {
            csvPath = UnitTestsHelper.DataFile_GuestAccountsUPN_Search;
        }

        DataTable dt = DataTable.New.ReadCsv(csvPath);

        foreach (Row row in dt.Rows)
        {
            var registrationData = new SearchEntityData();
            registrationData.Input = row["Input"];
            registrationData.ExpectedResultCount = Convert.ToInt32(row["ExpectedResultCount"]);
            registrationData.ExpectedEntityClaimValue = row["ExpectedEntityClaimValue"];
            registrationData.ResultType = row["ResultType"];
            registrationData.UserType = row["UserType"];
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
    public string ResultType;
    public string UserType;
}

public class ValidateEntityDataSource
{
    public static IEnumerable<TestCaseData> GetTestData(EntityDataSourceType entityDataSourceType)
    {
        string csvPath = UnitTestsHelper.DataFile_AllAccounts_Validate;
        if (entityDataSourceType == EntityDataSourceType.UPNB2BGuestAccounts)
        {
            csvPath = UnitTestsHelper.DataFile_GuestAccountsUPN_Validate;
        }

        DataTable dt = DataTable.New.ReadCsv(csvPath);

        foreach (Row row in dt.Rows)
        {
            var registrationData = new ValidateEntityData();
            registrationData.ClaimValue = row["ClaimValue"];
            registrationData.ShouldValidate = Convert.ToBoolean(row["ShouldValidate"]);
            registrationData.IsMemberOfTrustedGroup = Convert.ToBoolean(row["IsMemberOfTrustedGroup"]);
            registrationData.UserType = row["UserType"];
            yield return new TestCaseData(new object[] { registrationData });
        }
    }
}

public class ValidateEntityData
{
    public string ClaimValue;
    public bool ShouldValidate;
    public bool IsMemberOfTrustedGroup;
    public string UserType;
}
