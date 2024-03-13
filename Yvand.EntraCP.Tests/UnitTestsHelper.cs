using DataAccess;
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
        public static SPTrustedLoginProvider SPTrust => Utils.GetSPTrustAssociatedWithClaimsProvider(TestContext.Parameters["ClaimsProviderName"]);
        public static Uri TestSiteCollUri;
        public static string TestSiteRelativePath => $"/sites/{TestContext.Parameters["TestSiteCollectionName"]}";
        public const int MaxTime = 50000;
        public const int TestRepeatCount = 1;
        public static string FarmAdmin => TestContext.Parameters["FarmAdmin"];

        public static string RandomClaimType => "http://schemas.yvand.net/ws/claims/random";
        public static string RandomClaimValue => "IDoNotExist";
        public static DirectoryObjectProperty RandomObjectProperty => DirectoryObjectProperty.AccountEnabled;

        public static string AzureTenantsJsonFile => TestContext.Parameters["AzureTenantsJsonFile"];
        public static string DataFile_EntraId_TestUsers => TestContext.Parameters["DataFile_EntraId_TestUsers"];
        public static string DataFile_EntraId_TestGroups => TestContext.Parameters["DataFile_EntraId_TestGroups"];
        public static string UserAccountNamePrefix => TestContext.Parameters["UserAccountNamePrefix"];
        public static string GroupAccountNamePrefix => TestContext.Parameters["GroupAccountNamePrefix"];
        public static string[] UsersMembersOfAllGroups => new string[] { $"{UserAccountNamePrefix}001", $"{UserAccountNamePrefix}010", $"{UserAccountNamePrefix}011", $"{UserAccountNamePrefix}012", $"{UserAccountNamePrefix}013", $"{UserAccountNamePrefix}014", $"{UserAccountNamePrefix}015" };
        public const int TotalNumberOfUsersInSource = 50 + 3; // 50 members + 3 guests
        public const int TotalNumberOfGroupsInSource = 50;
        static TextWriterTraceListener Logger { get; set; }
        public static EntraIDProviderConfiguration PersistedConfiguration;
        private static IEntraIDProviderSettings OriginalSettings;

        [OneTimeSetUp]
        public static void InitializeSiteCollection()
        {
            Logger = new TextWriterTraceListener(TestContext.Parameters["TestLogFileName"]);
            Trace.Listeners.Add(Logger);
            Trace.AutoFlush = true;
            Trace.TraceInformation($"{DateTime.Now:s} Start integration tests of {EntraCP.ClaimsProviderName} {FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(EntraCP)).Location).FileVersion}.");
            Trace.TraceInformation($"{DateTime.Now:s} DataFile_EntraId_TestGroups: {DataFile_EntraId_TestGroups}");
            Trace.TraceInformation($"{DateTime.Now:s} DataFile_EntraId_TestUsers: {DataFile_EntraId_TestUsers}");
            Trace.TraceInformation($"{DateTime.Now:s} TestSiteCollectionName: {TestContext.Parameters["TestSiteCollectionName"]}");

            if (SPTrust == null)
            {
                Trace.TraceError($"{DateTime.Now:s} SPTrust: is null");
            }
            else
            {
                Trace.TraceInformation($"{DateTime.Now:s} SPTrust: {SPTrust.Name}");
            }

            PersistedConfiguration = EntraCP.GetConfiguration(true);
            if (PersistedConfiguration != null)
            {
                OriginalSettings = PersistedConfiguration.Settings;
                Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Took a backup of the original settings");
            }
            else
            {
                PersistedConfiguration = EntraCP.CreateConfiguration();
                Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Persisted configuration not found, created it");
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
            //string encodedClaim = claimMgr.EncodeClaim(TrustedGroup);
            //SPUserInfo userInfo = new SPUserInfo { LoginName = encodedClaim, Name = TrustedGroupToAdd_ClaimValue };

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
                //membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
                spSite.Dispose();
            }

            if (!SPSite.Exists(TestSiteCollUri))
            {
                Trace.TraceInformation($"{DateTime.Now:s} Creating site collection {TestSiteCollUri.AbsoluteUri} with template '{spSiteTemplate}'...");
                SPSite spSite = wa.Sites.Add(TestSiteCollUri.AbsoluteUri, EntraCP.ClaimsProviderName, EntraCP.ClaimsProviderName, 1033, spSiteTemplate, FarmAdmin, String.Empty, String.Empty);
                spSite.RootWeb.CreateDefaultAssociatedGroups(FarmAdmin, FarmAdmin, spSite.RootWeb.Title);

                SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                //membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
                spSite.Dispose();
            }
            else
            {
                using (SPSite spSite = new SPSite(TestSiteCollUri.AbsoluteUri))
                {
                    SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                    //membersGroup.AddUser(userInfo.LoginName, userInfo.Email, userInfo.Name, userInfo.Notes);
                }
            }
        }

        [OneTimeTearDown]
        public static void Cleanup()
        {
            Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Cleanup.");
            try
            {
                if (PersistedConfiguration != null && OriginalSettings != null)
                {
                    PersistedConfiguration.ApplySettings(OriginalSettings, true);
                    Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Restored original settings of LDAPCPSE configuration");
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError($"{DateTime.Now:s} [SETUP] Unexpected error while restoring the original settings of LDAPCPSE configuration: {ex.Message}");
            }

            Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Integration tests of {EntraCP.ClaimsProviderName} {FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(EntraCP)).Location).FileVersion} finished.");
            Trace.Flush();
            if (Logger != null)
            {
                Logger.Dispose();
            }
        }
    }

    public class EntraIdTestGroup
    {
        public string Id;
        public string DisplayName;
        public string GroupType;
        public bool SecurityEnabled;
    }

    public class EntraIdTestGroupsSource
    {
        private static object _LockInitList = new object();
        private static List<EntraIdTestGroup> _Groups;
        public static List<EntraIdTestGroup> Groups
        {
            get
            {
                if (_Groups != null) { return _Groups; }
                lock (_LockInitList)
                {
                    if (_Groups != null) { return _Groups; }
                    _Groups = new List<EntraIdTestGroup>();
                    foreach (EntraIdTestGroup group in GetTestData(false))
                    {
                        _Groups.Add(group);
                    }
                    return _Groups;
                }
            }
        }

        public static EntraIdTestGroup ASecurityEnabledGroup => Groups.FirstOrDefault(x => x.SecurityEnabled);
        public static EntraIdTestGroup ANonSecurityEnabledGroup => Groups.FirstOrDefault(x => !x.SecurityEnabled);

        public static IEnumerable<EntraIdTestGroup> GetTestData(bool securityEnabledGroupsOnly = false)
        {
            string csvPath = UnitTestsHelper.DataFile_EntraId_TestGroups;
            DataTable dt = DataTable.New.ReadCsv(csvPath);
            foreach (Row row in dt.Rows)
            {
                var registrationData = new EntraIdTestGroup();
                registrationData.Id = row["id"];
                registrationData.DisplayName = row["displayName"];
                registrationData.GroupType = row["groupType"];
                registrationData.SecurityEnabled = Convert.ToBoolean(row["SecurityEnabled"]);
                if (securityEnabledGroupsOnly && !registrationData.SecurityEnabled)
                {
                    continue;
                }
                yield return registrationData;
            }
        }
    }

    public enum UserType
    {
        Member,
        Guest
    }

    public class EntraIdTestUser
    {
        public string Id;
        public string DisplayName;
        public string UserPrincipalName;
        public UserType UserType;
        public string Mail;
        public string GivenName;
    }

    public class EntraIdTestUsersSource
    {
        private static object _LockInitList = new object();
        private static List<EntraIdTestUser> _Users;
        public static List<EntraIdTestUser> Users
        {
            get
            {
                if (_Users != null) { return _Users; }
                lock (_LockInitList)
                {
                    if (_Users != null) { return _Users; }
                    _Users = new List<EntraIdTestUser>();
                    foreach (EntraIdTestUser user in GetTestData())
                    {
                        _Users.Add(user);
                    }
                    return _Users;
                }
            }
        }

        public static EntraIdTestUser AGuestUser => Users.FirstOrDefault(x => x.UserType == UserType.Guest);
        public static IEnumerable<EntraIdTestUser> AllGuestUsers => Users.Where(x => x.UserType == UserType.Guest);

        public static IEnumerable<EntraIdTestUser> GetTestData()
        {
            string csvPath = UnitTestsHelper.DataFile_EntraId_TestUsers;
            DataTable dt = DataTable.New.ReadCsv(csvPath);
            foreach (Row row in dt.Rows)
            {
                var registrationData = new EntraIdTestUser();
                registrationData.Id = row["id"];
                registrationData.DisplayName = row["displayName"];
                registrationData.UserPrincipalName = row["userPrincipalName"];
                registrationData.UserType = String.Equals(row["userType"], ClaimsProviderConstants.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase) ? UserType.Member : UserType.Guest;
                registrationData.Mail = row["mail"];
                registrationData.GivenName = row["givenName"];
                yield return registrationData;
            }
        }
    }
}