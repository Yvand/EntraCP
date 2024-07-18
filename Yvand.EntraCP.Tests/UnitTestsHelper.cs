﻿using DataAccess;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.Web.Hosting.Administration;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
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
        public static string TestUsersAccountNamePrefix => TestContext.Parameters["UserAccountNamePrefix"];
        public static string TestGroupsAccountNamePrefix => TestContext.Parameters["GroupAccountNamePrefix"];
        
        static TextWriterTraceListener Logger { get; set; }
        public static EntraIDProviderConfiguration PersistedConfiguration;
        private static IEntraIDProviderSettings OriginalSettings;

        private static EntraIDTenant _TenantConnection;
        public static EntraIDTenant TenantConnection
        {
            get
            {
                if (_TenantConnection != null) { return _TenantConnection; }
                string json = File.ReadAllText(UnitTestsHelper.AzureTenantsJsonFile);
                List<EntraIDTenant> azureTenants = JsonConvert.DeserializeObject<List<EntraIDTenant>>(json);
                _TenantConnection = azureTenants.First();
                return _TenantConnection;
            }
        }

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
            return; // Uncommented when debugging from unit tests
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

    public class TestEntity : ICloneable
    {
        public string Id;
        public string DisplayName;

        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }

    public class EntraIdTestGroup : TestEntity 
    {
        public string GroupType;
        public bool SecurityEnabled = true;
    }

    public class EntraIdTestGroupSettings : EntraIdTestGroup
    {
        public bool AllTestUsersAreMembers = false;

        private static object _LockInitGroupsSettingsList = new object();
        private static List<EntraIdTestGroupSettings> _GroupsSettings;
        public static List<EntraIdTestGroupSettings> GroupsSettings
        {
            get
            {
                if (_GroupsSettings != null) { return _GroupsSettings; }
                lock (_LockInitGroupsSettingsList)
                {
                    if (_GroupsSettings != null) { return _GroupsSettings; }
                    _GroupsSettings = new List<EntraIdTestGroupSettings>
                        {
                            new EntraIdTestGroupSettings { DisplayName = $"{UnitTestsHelper.TestGroupsAccountNamePrefix}001" , SecurityEnabled = false, AllTestUsersAreMembers = true},
                            new EntraIdTestGroupSettings { DisplayName = $"{UnitTestsHelper.TestGroupsAccountNamePrefix}005" , SecurityEnabled = true, AllTestUsersAreMembers = true },
                            new EntraIdTestGroupSettings { DisplayName = $"{UnitTestsHelper.TestGroupsAccountNamePrefix}008" , SecurityEnabled = false, AllTestUsersAreMembers = false },
                            new EntraIdTestGroupSettings { DisplayName = $"{UnitTestsHelper.TestGroupsAccountNamePrefix}018" , SecurityEnabled = false, AllTestUsersAreMembers = true },
                            new EntraIdTestGroupSettings { DisplayName = $"{UnitTestsHelper.TestGroupsAccountNamePrefix}025" , SecurityEnabled = true, AllTestUsersAreMembers = true },
                            new EntraIdTestGroupSettings { DisplayName = $"{UnitTestsHelper.TestGroupsAccountNamePrefix}028" , SecurityEnabled = false, AllTestUsersAreMembers = false, },
                            new EntraIdTestGroupSettings { DisplayName = $"{UnitTestsHelper.TestGroupsAccountNamePrefix}038" , SecurityEnabled = false, AllTestUsersAreMembers = true, },
                            new EntraIdTestGroupSettings { DisplayName = $"{UnitTestsHelper.TestGroupsAccountNamePrefix}048" , SecurityEnabled = false, AllTestUsersAreMembers = false, },
                        };
                    foreach (EntraIdTestGroupSettings groupsSetting in _GroupsSettings)
                    {
                        groupsSetting.Id = TestEntitySourceManager.TestGroupsSource.Entities.First(x => x.DisplayName == groupsSetting.DisplayName).Id;
                    }
                    Trace.TraceInformation($"{DateTime.Now:s} [{typeof(EntraIdTestGroupSettings).Name}] Initialized List of {nameof(GroupsSettings)} with {_GroupsSettings.Count} items.");
                }
                return _GroupsSettings;
            }
        }
    }

    public enum UserType
    {
        Member,
        Guest
    }

    public class EntraIdTestUser : TestEntity
    {
        public string UserPrincipalName;
        public UserType UserType;
        public string Mail;
        public string GivenName;        
    }

    public class EntraIdTestUserSettings : EntraIdTestUser
    {
        public bool IsMemberOfAllGroups = false;

        private static object _LockInitUsersWithSpecificSettingsList = new object();
        private static List<EntraIdTestUserSettings> _UsersWithSpecificSettings;
        public static List<EntraIdTestUserSettings> UsersWithSpecificSettings
        {
            get
            {
                if (_UsersWithSpecificSettings != null) { return _UsersWithSpecificSettings; }
                lock (_LockInitUsersWithSpecificSettingsList)
                {
                    if (_UsersWithSpecificSettings != null) { return _UsersWithSpecificSettings; }
                    _UsersWithSpecificSettings = new List<EntraIdTestUserSettings>
                        {
                            new EntraIdTestUserSettings { UserPrincipalName = $"{UnitTestsHelper.TestUsersAccountNamePrefix}001@{UnitTestsHelper.TenantConnection.Name}" , IsMemberOfAllGroups = true },
                            new EntraIdTestUserSettings { UserPrincipalName = $"{UnitTestsHelper.TestUsersAccountNamePrefix}010@{UnitTestsHelper.TenantConnection.Name}" , IsMemberOfAllGroups = true },
                            new EntraIdTestUserSettings { UserPrincipalName = $"{UnitTestsHelper.TestUsersAccountNamePrefix}011@{UnitTestsHelper.TenantConnection.Name}" , IsMemberOfAllGroups = true },
                            new EntraIdTestUserSettings { UserPrincipalName = $"{UnitTestsHelper.TestUsersAccountNamePrefix}012@{UnitTestsHelper.TenantConnection.Name}" , IsMemberOfAllGroups = true },
                            new EntraIdTestUserSettings { UserPrincipalName = $"{UnitTestsHelper.TestUsersAccountNamePrefix}013@{UnitTestsHelper.TenantConnection.Name}" , IsMemberOfAllGroups = true },
                            new EntraIdTestUserSettings { UserPrincipalName = $"{UnitTestsHelper.TestUsersAccountNamePrefix}014@{UnitTestsHelper.TenantConnection.Name}" , IsMemberOfAllGroups = true },
                            new EntraIdTestUserSettings { UserPrincipalName = $"{UnitTestsHelper.TestUsersAccountNamePrefix}015@{UnitTestsHelper.TenantConnection.Name}" , IsMemberOfAllGroups = true },
                        };
                }
                Trace.TraceInformation($"{DateTime.Now:s} [{typeof(EntraIdTestUserSettings).Name}] Initialized List of {nameof(UsersWithSpecificSettings)} with {_UsersWithSpecificSettings.Count} items.");
                return _UsersWithSpecificSettings;
            }
        }
    }

    public class TestEntitySource<T> where T : TestEntity, new()
    {
        private object _LockInitGroupsList = new object();
        private bool listInitialized = false;
        private List<T> _Entities;
        public List<T> Entities
        {
            get
            {
                if (listInitialized) { return _Entities; }
                lock (_LockInitGroupsList)
                {
                    if (_Entities != null) { return _Entities; }
                    _Entities = new List<T>();
                    foreach (T entity in ReadDataSource())
                    {
                        _Entities.Add(entity);
                    }
                    listInitialized = true;
                    Trace.TraceInformation($"{DateTime.Now:s} [{typeof(T).Name}] Initialized List of {nameof(Entities)} with {Entities.Count} items.");
                    return _Entities;
                }
            }
        }
        private Random RandomNumber = new Random();


        private IEnumerable<T> ReadDataSource()
        {
            string csvPath = String.Empty;
            if (typeof(T) == typeof(EntraIdTestUser))
            {
                csvPath = UnitTestsHelper.DataFile_EntraId_TestUsers;
            }
            else
            {
                csvPath = UnitTestsHelper.DataFile_EntraId_TestGroups;
            }
            DataTable dt = DataTable.New.ReadCsv(csvPath);
            foreach (Row row in dt.Rows)
            {
                if (typeof(T) == typeof(EntraIdTestUser))
                {
                    EntraIdTestUser registrationData = new T() as EntraIdTestUser;
                    registrationData.Id = row["id"];
                    registrationData.DisplayName = row["displayName"];
                    registrationData.UserPrincipalName = row["userPrincipalName"];
                    registrationData.UserType = String.Equals(row["userType"], ClaimsProviderConstants.MEMBER_USERTYPE, StringComparison.InvariantCultureIgnoreCase) ? UserType.Member : UserType.Guest;
                    registrationData.Mail = row["mail"];
                    registrationData.GivenName = row["givenName"];
                    yield return registrationData as T;
                }
                else
                {
                    EntraIdTestGroup registrationData = new EntraIdTestGroup();
                    registrationData.Id = row["id"];
                    registrationData.DisplayName = row["displayName"];
                    registrationData.GroupType = row["groupType"];
                    registrationData.SecurityEnabled = Convert.ToBoolean(row["SecurityEnabled"]);
                    yield return registrationData as T;
                }
            }
        }

        public IEnumerable<T> GetSomeEntities(int count, Func<T, bool> filter = null)
        {
            if (count > Entities.Count)
            {
                count = Entities.Count;
            }

            List<int> entitiesIdxs = new List<int>(count);
            for (int i = 0; i < count; i++)
            {
                entitiesIdxs.Add(RandomNumber.Next(0, Entities.Where(filter ?? (x => true)).Count() - 1));
            }

            foreach (int userIdx in entitiesIdxs)
            {
                yield return Entities[userIdx].Clone() as T;
            }
        }        
    }

    public class TestEntitySourceManager
    {
        public static TestEntitySource<EntraIdTestUser> TestUsersSource = new TestEntitySource<EntraIdTestUser>();
        public static TestEntitySource<EntraIdTestGroup> TestGroupsSource = new TestEntitySource<EntraIdTestGroup>();
        public const int MaxNumberOfUsersToTest = 100;
        public const int MaxNumberOfGroupsToTest = 100;

        public static IEnumerable<EntraIdTestUser> GetSomeUsers(int count)
        {
            return TestUsersSource.GetSomeEntities(count, null);
        }

        public static IEnumerable<EntraIdTestGroup> GetSomeGroups(int count, bool securityEnabledOnly)
        {
            Func<EntraIdTestGroup, bool> securityEnabledOnlyFilter = x => x.SecurityEnabled == securityEnabledOnly;
            return TestGroupsSource.GetSomeEntities(count, securityEnabledOnlyFilter);
        }
    }
}