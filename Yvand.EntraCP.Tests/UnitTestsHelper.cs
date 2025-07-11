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
        public static readonly string ClaimsProviderName = TestContext.Parameters["ClaimsProviderName"];
        public static readonly EntraCP ClaimsProvider = new EntraCP(ClaimsProviderName);
        public static SPTrustedLoginProvider SPTrust => Utils.GetSPTrustAssociatedWithClaimsProvider(ClaimsProviderName);
        public static Uri TestSiteCollUri;
        public static string TestSiteRelativePath => $"/sites/{ClaimsProviderName}.UnitTests";
        public const int MaxTime = 50000;
        public const int TestRepeatCount = 1;
        public static string FarmAdmin => TestContext.Parameters["FarmAdmin"];

        public static string RandomClaimType => "http://schemas.yvand.net/ws/claims/random";
        public static string RandomClaimValue => "IDoNotExist";
        public static DirectoryObjectProperty RandomObjectProperty => DirectoryObjectProperty.AccountEnabled;

        public static string AzureTenantsJsonFile => TestContext.Parameters["AzureTenantsJsonFile"];
        public static string DataFile_EntraId_TestUsers => TestContext.Parameters["DataFile_EntraId_TestUsers"];
        public static string DataFile_EntraId_TestGroups => TestContext.Parameters["DataFile_EntraId_TestGroups"];
        public static string GroupsClaimType => TestContext.Parameters["GroupsClaimType"];

        static TextWriterTraceListener Logger { get; set; }
        public static EntraIDProviderConfiguration PersistedConfiguration;
        private static IEntraIDProviderSettings OriginalSettings;

        [OneTimeSetUp]
        public static void InitializeSiteCollection()
        {
            Logger = new TextWriterTraceListener($"{ClaimsProviderName}IntegrationTests.log");
            Trace.Listeners.Add(Logger);
            Trace.AutoFlush = true;
            Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Start integration tests of {EntraCP.ClaimsProviderName} {FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(EntraCP)).Location).FileVersion}.");
            Trace.TraceInformation($"{DateTime.Now:s} [SETUP] DataFile_EntraId_TestGroups: {DataFile_EntraId_TestGroups}");
            Trace.TraceInformation($"{DateTime.Now:s} [SETUP] DataFile_EntraId_TestUsers: {DataFile_EntraId_TestUsers}");
            Trace.TraceInformation($"{DateTime.Now:s} [SETUP] TestSiteCollectionName: {TestContext.Parameters["TestSiteCollectionName"]}");

            if (SPTrust == null)
            {
                Trace.TraceError($"{DateTime.Now:s} [SETUP] SPTrust: is null");
            }
            else
            {
                Trace.TraceInformation($"{DateTime.Now:s} [SETUP] SPTrust: {SPTrust.Name}");
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
                Trace.TraceError($"{DateTime.Now:s} [SETUP] Web application was NOT found.");
                return;
            }

            Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Web application {wa.Name} found.");
            Uri waRootAuthority = wa.AlternateUrls[0].Uri;
            TestSiteCollUri = new Uri($"{waRootAuthority.GetLeftPart(UriPartial.Authority)}{TestSiteRelativePath}");
            SPClaimProviderManager claimMgr = SPClaimProviderManager.Local;
            string trustedGroupName = TestEntitySourceManager.GetOneGroup().Id;
            string encodedGroupClaim = claimMgr.EncodeClaim(new SPClaim(GroupsClaimType, trustedGroupName, ClaimValueTypes.String, SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, UnitTestsHelper.SPTrust.Name)));
            SPUserInfo groupInfo = new SPUserInfo { LoginName = encodedGroupClaim, Name = trustedGroupName };

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
                Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Creating root site collection {waRootAuthority.AbsoluteUri}...");
                SPSite spSite = wa.Sites.Add(waRootAuthority.AbsoluteUri, "root", "root", 1033, spSiteTemplate, FarmAdmin, String.Empty, String.Empty);
                spSite.RootWeb.CreateDefaultAssociatedGroups(FarmAdmin, FarmAdmin, spSite.RootWeb.Title);

                SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                membersGroup.AddUser(groupInfo.LoginName, groupInfo.Email, groupInfo.Name, groupInfo.Notes);
                spSite.Dispose();
            }

            if (!SPSite.Exists(TestSiteCollUri))
            {
                Trace.TraceInformation($"{DateTime.Now:s} [SETUP] Creating site collection {TestSiteCollUri.AbsoluteUri} with template '{spSiteTemplate}'...");
                SPSite spSite = wa.Sites.Add(TestSiteCollUri.AbsoluteUri, EntraCP.ClaimsProviderName, EntraCP.ClaimsProviderName, 1033, spSiteTemplate, FarmAdmin, String.Empty, String.Empty);
                spSite.RootWeb.CreateDefaultAssociatedGroups(FarmAdmin, FarmAdmin, spSite.RootWeb.Title);

                SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                membersGroup.AddUser(groupInfo.LoginName, groupInfo.Email, groupInfo.Name, groupInfo.Notes);
                spSite.Dispose();
            }
            else
            {
                using (SPSite spSite = new SPSite(TestSiteCollUri.AbsoluteUri))
                {
                    SPGroup membersGroup = spSite.RootWeb.AssociatedMemberGroup;
                    membersGroup.AddUser(groupInfo.LoginName, groupInfo.Email, groupInfo.Name, groupInfo.Notes);
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

    public abstract class TestEntity : ICloneable
    {
        public string Id;
        public string DisplayName;

        public object Clone()
        {
            return this.MemberwiseClone();
        }

        public abstract void SetEntityFromDataSourceRow(Row row);
    }

    public class TestUser : TestEntity
    {
        public string UserPrincipalName;
        public UserType UserType;
        public string Mail;
        public string GivenName;
        public bool IsMemberOfAllGroups = false;
        public bool AccountEnabled = false;

        public override void SetEntityFromDataSourceRow(Row row)
        {
            Id = row["id"];
            DisplayName = row["displayName"];
            UserPrincipalName = row["userPrincipalName"];
            UserType = String.Equals(row["userType"], ClaimsProviderConstants.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase) ? UserType.Member : UserType.Guest;
            Mail = row["mail"];
            GivenName = row["givenName"];
            IsMemberOfAllGroups = Convert.ToBoolean(row["IsMemberOfAllGroups"]);
            AccountEnabled = Convert.ToBoolean(row["AccountEnabled"]);
        }
    }

    public class TestGroup : TestEntity
    {
        public string GroupType;
        public bool SecurityEnabled = true;
        public bool EveryoneIsMember;

        public override void SetEntityFromDataSourceRow(Row row)
        {
            Id = row["id"];
            DisplayName = row["displayName"];
            GroupType = row["groupType"];
            SecurityEnabled = Convert.ToBoolean(row["SecurityEnabled"]);
            EveryoneIsMember = Convert.ToBoolean(row["EveryoneIsMember"]);
        }
    }

    public enum UserType
    {
        Member,
        Guest
    }

    public class TestEntitySource<T> where T : TestEntity, new()
    {
        private object _LockInitEntitiesList = new object();
        private bool EntitiesReady = false;
        private List<T> _Entities = new List<T>();
        public List<T> Entities
        {
            get
            {
                if (EntitiesReady) { return _Entities; }
                lock (_LockInitEntitiesList)
                {
                    if (EntitiesReady) { return _Entities; }
                    foreach (T entity in ReadDataSource())
                    {
                        _Entities.Add(entity);
                    }
                    EntitiesReady = true;
                    Trace.TraceInformation($"{DateTime.Now:s} [{typeof(T).Name}] Initialized List of {nameof(Entities)} with {Entities.Count} items.");
                    return _Entities;
                }
            }
        }

        private Random RandomNumber = new Random();
        private string DataSourceFilePath;

        public TestEntitySource(string dataSourceFilePath)
        {
            DataSourceFilePath = dataSourceFilePath;
        }

        private IEnumerable<T> ReadDataSource()
        {
            DataTable dt = DataTable.New.ReadCsv(DataSourceFilePath);
            foreach (Row row in dt.Rows)
            {
                T entity = new T();
                entity.SetEntityFromDataSourceRow(row);
                yield return entity;
            }
        }

        public IEnumerable<T> GetSomeEntities(int count, Func<T, bool> filter = null)
        {
            IEnumerable<T> entitiesFiltered = Entities.Where(filter ?? (x => true));
            int entitiesFilteredCount = entitiesFiltered.Count();
            if (count > entitiesFilteredCount) { count = entitiesFilteredCount; }
            for (int i = 0; i < count; i++)
            {
                int randomIdx = RandomNumber.Next(0, entitiesFilteredCount);
                yield return entitiesFiltered.ElementAt(randomIdx).Clone() as T;
            }
        }
    }

    public class TestEntitySourceManager
    {
        private static TestEntitySource<TestUser> TestUsersSource = new TestEntitySource<TestUser>(UnitTestsHelper.DataFile_EntraId_TestUsers);
        public static List<TestUser> AllTestUsers
        {
            get => TestUsersSource.Entities;
        }
        private static TestEntitySource<TestGroup> TestGroupsSource = new TestEntitySource<TestGroup>(UnitTestsHelper.DataFile_EntraId_TestGroups);
        public static List<TestGroup> AllTestGroups
        {
            get => TestGroupsSource.Entities;
        }
        public const int MaxNumberOfUsersToTest = 100;
        public const int MaxNumberOfGroupsToTest = 100;

        public static IEnumerable<TestUser> GetSomeUsers(int count)
        {
            return TestUsersSource.GetSomeEntities(count, null);
        }

        public static IEnumerable<TestUser> GetSomeDisabledUsers(int count)
        {
            Func<TestUser, bool> filter = x => x.AccountEnabled == false;
            return TestUsersSource.GetSomeEntities(count, null);
        }

        public static IEnumerable<TestUser> GetUsersMembersOfAllGroups()
        {
            Func<TestUser, bool> filter = x => x.IsMemberOfAllGroups == true;
            return TestUsersSource.GetSomeEntities(Int16.MaxValue, filter);
        }

        public static TestUser FindUser(string upnPrefix)
        {
            Func<TestUser, bool> filter = x => x.UserPrincipalName.StartsWith(upnPrefix);
            return TestUsersSource.GetSomeEntities(1, filter).First();
        }

        public static TestUser GetOneUser(UserType userType)
        {
            Func<TestUser, bool> filter = x => x.UserType == userType;
            return TestUsersSource.GetSomeEntities(1, filter).First();
        }

        public static IEnumerable<TestGroup> GetSomeGroups(int count, bool securityEnabledOnly)
        {
            Func<TestGroup, bool> filter = x => x.SecurityEnabled == securityEnabledOnly;
            return TestGroupsSource.GetSomeEntities(count, filter);
        }

        public static TestGroup GetOneGroup()
        {
            return TestGroupsSource.GetSomeEntities(1, null).First();
        }

        public static TestGroup GetOneGroup(bool securityEnabledOnly)
        {
            Func<TestGroup, bool> filter = x => x.SecurityEnabled == securityEnabledOnly;
            return TestGroupsSource.GetSomeEntities(1, filter).First();
        }
    }
}