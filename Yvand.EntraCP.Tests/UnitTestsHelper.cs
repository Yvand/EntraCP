﻿using DataAccess;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Security.Claims;
using Yvand.EntraClaimsProvider.Configuration;

namespace Yvand.EntraClaimsProvider.Tests
{
    [SetUpFixture]
    public class UnitTestsHelper
    {
        public static readonly EntraCP ClaimsProvider = new EntraCP(TestContext.Parameters["ClaimsProviderName"]);
        public static SPTrustedLoginProvider SPTrust => SPSecurityTokenServiceManager.Local.TrustedLoginProviders.FirstOrDefault(x => String.Equals(x.ClaimProviderName, TestContext.Parameters["ClaimsProviderName"], StringComparison.InvariantCultureIgnoreCase));
        public static Uri TestSiteCollUri;
        public static string TestSiteRelativePath => $"/sites/{TestContext.Parameters["TestSiteCollectionName"]}";
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
        public static SPClaim TrustedGroup => new SPClaim(TrustedGroupToAdd_ClaimType, TrustedGroupToAdd_ClaimValue, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, SPTrust.Name));

        public static string AzureTenantsJsonFile => TestContext.Parameters["AzureTenantsJsonFile"];
        public static string DataFile_GuestAccountsUPN_Search => TestContext.Parameters["DataFile_GuestAccountsUPN_Search"];
        public static string DataFile_GuestAccountsUPN_Validate => TestContext.Parameters["DataFile_GuestAccountsUPN_Validate"];
        public static string DataFile_AllAccounts_Search => TestContext.Parameters["DataFile_AllAccounts_Search"];
        public static string DataFile_AllAccounts_Validate => TestContext.Parameters["DataFile_AllAccounts_Validate"];
        static TextWriterTraceListener Logger { get; set; }

        [OneTimeSetUp]
        public static void InitializeSiteCollection()
        {
            Logger = new TextWriterTraceListener(TestContext.Parameters["TestLogFileName"]);
            Trace.Listeners.Add(Logger);
            Trace.AutoFlush = true;
            Trace.TraceInformation($"{DateTime.Now:s} Start integration tests of {EntraCP.ClaimsProviderName} {FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(EntraCP)).Location).FileVersion}.");
            Trace.TraceInformation($"{DateTime.Now:s} DataFile_AllAccounts_Search: {DataFile_AllAccounts_Search}");
            Trace.TraceInformation($"{DateTime.Now:s} DataFile_AllAccounts_Validate: {DataFile_AllAccounts_Validate}");
            Trace.TraceInformation($"{DateTime.Now:s} DataFile_GuestAccountsUPN_Search: {DataFile_GuestAccountsUPN_Search}");
            Trace.TraceInformation($"{DateTime.Now:s} DataFile_GuestAccountsUPN_Validate: {DataFile_GuestAccountsUPN_Validate}");
            Trace.TraceInformation($"{DateTime.Now:s} TestSiteCollectionName: {TestContext.Parameters["TestSiteCollectionName"]}");

            if (SPTrust == null)
            {
                Trace.TraceError($"{DateTime.Now:s} SPTrust: is null");
            }
            else
            {
                Trace.TraceInformation($"{DateTime.Now:s} SPTrust: {SPTrust.Name}");
            }

#if DEBUG
            TestSiteCollUri = new Uri($"http://spsites{TestSiteRelativePath}");
            //return; // Uncommented when debugging from unit tests
#endif

            var service = SPFarm.Local.Services.GetValue<SPWebService>(String.Empty);
            SPWebApplication wa = service.WebApplications.FirstOrDefault(x =>
            {
                foreach (var iisSetting in x.IisSettings.Values)
                {
                    foreach (SPAuthenticationProvider authenticationProviders in iisSetting.ClaimsAuthenticationProviders)
                    {
                        if (String.Equals(authenticationProviders.ClaimProviderName, EntraCP.ClaimsProviderName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
                return false;
            });
            if (wa == null)
            {
                Trace.TraceError($"{DateTime.Now:s} Web application was NOT found.");
                return;
            }

            Trace.TraceInformation($"{DateTime.Now:s} Web application {wa.Name} found.");
            Uri waRootAuthority = wa.AlternateUrls[0].Uri;
            TestSiteCollUri = new Uri($"{waRootAuthority.GetLeftPart(UriPartial.Authority)}{TestSiteRelativePath}");
            SPClaimProviderManager claimMgr = SPClaimProviderManager.Local;
            string encodedClaim = claimMgr.EncodeClaim(TrustedGroup);
            SPUserInfo userInfo = new SPUserInfo { LoginName = encodedClaim, Name = TrustedGroupToAdd_ClaimValue };

            FileVersionInfo spAssemblyVersion = FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(SPSite)).Location);
            string spSiteTemplate = "STS#3"; // modern team site template
            if (spAssemblyVersion.FileBuildPart < 10000)
            {
                // If SharePoint 2016, must use the classic team site template
                spSiteTemplate = "STS#0";
            }

            // The root site may not exist, but it must be present for tests to run
            if (!SPSite.Exists(waRootAuthority))
            {
                Trace.TraceInformation($"{DateTime.Now:s} Creating root site collection {waRootAuthority.AbsoluteUri}...");
                SPSite spSite = wa.Sites.Add(waRootAuthority.AbsoluteUri, "root", "root", 1033, spSiteTemplate, FarmAdmin, String.Empty, String.Empty);
                spSite.RootWeb.CreateDefaultAssociatedGroups(FarmAdmin, FarmAdmin, spSite.RootWeb.Title);

                SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
                spSite.Dispose();
            }

            if (!SPSite.Exists(TestSiteCollUri))
            {
                Trace.TraceInformation($"{DateTime.Now:s} Creating site collection {TestSiteCollUri.AbsoluteUri} with template '{spSiteTemplate}'...");
                SPSite spSite = wa.Sites.Add(TestSiteCollUri.AbsoluteUri, EntraCP.ClaimsProviderName, EntraCP.ClaimsProviderName, 1033, spSiteTemplate, FarmAdmin, String.Empty, String.Empty);
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
            Trace.TraceInformation($"{DateTime.Now:s} Integration tests of {EntraCP.ClaimsProviderName} {FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(EntraCP)).Location).FileVersion} finished.");
            Trace.Flush();
            if (Logger != null)
            {
                Logger.Dispose();
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

    public enum ResultUserType
    {
        None,
        Mixed,
        Member,
        Guest,
    }

    public enum ResultEntityType
    {
        None,
        Mixed,
        User,
        Group,
    }

    public enum EntityDataSourceType
    {
        AllAccounts,
        UPNB2BGuestAccounts
    }

    public class SearchEntityData
    {
        public string Input;
        public int SearchResultCount;
        public string SearchResultSingleEntityClaimValue;
        public ResultEntityType SearchResultEntityTypes;
        public ResultUserType SearchResultUserTypes;
        public bool ExactMatch;
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
                registrationData.SearchResultCount = Convert.ToInt32(row["SearchResultCount"]);
                registrationData.SearchResultSingleEntityClaimValue = row["SearchResultSingleEntityClaimValue"];
                registrationData.SearchResultEntityTypes = (ResultEntityType) Enum.Parse(typeof(ResultEntityType), row["SearchResultEntityTypes"]);
                registrationData.SearchResultUserTypes = (ResultUserType)Enum.Parse(typeof(ResultUserType), row["SearchResultUserTypes"]);
                registrationData.ExactMatch = Convert.ToBoolean(row["ExactMatch"]);
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
                registrationData.EntityType = (ResultEntityType)Enum.Parse(typeof(ResultEntityType), row["EntityType"]);
                registrationData.UserType = (ResultUserType)Enum.Parse(typeof(ResultUserType), row["UserType"]);
                yield return new TestCaseData(new object[] { registrationData });
            }
        }
    }

    public class ValidateEntityData
    {
        public string ClaimValue;
        public bool ShouldValidate;
        public bool IsMemberOfTrustedGroup;
        public ResultEntityType EntityType;
        public ResultUserType UserType;
    }
}