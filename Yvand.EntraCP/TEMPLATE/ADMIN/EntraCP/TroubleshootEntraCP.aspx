<%@ Page Language="C#" AutoEventWireup="true" Inherits="Microsoft.SharePoint.WebControls.LayoutsPageBase" DynamicMasterPageFile="~masterurl/default.master" %>

<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Reflection" %>
<%@ Import Namespace="System.Diagnostics" %>
<%@ Import Namespace="System.Threading.Tasks" %>
<%@ Import Namespace="Yvand.EntraClaimsProvider" %>
<%@ Import Namespace="Yvand.EntraClaimsProvider.Configuration" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Import Namespace="Microsoft.SharePoint.Administration.Claims" %>
<%@ Import Namespace="System.Security.Claims" %>
<%@ Import Namespace="System.IdentityModel.Tokens" %>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">Troubleshoot EntraCP</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">Troubleshoot common issues with EntraCP</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <script runat="server" language="C#">
        public static class Config
        {
            // REPLACE ONLY THOSE VALUES BELOW TO RUN THE TESTS AGAINST YOUR TENANT
            public static string TenantName = "TOREPLACE";
            public static string TenantClientId = "TOREPLACE";
            public static string TenantClientSecret = "TOREPLACE";
            public static string Proxy = "";
            public static string Input = "yvand";

            public static string context = SPContext.Current.Web.Url;
            public static string ClaimsProviderName = "EntraCP";
            public static string IconSuccess = "<span class='ms-status-iconSpan'><img src='/_layouts/15/images/kpinormal-0.gif'></span>";
            public static string IconWarning = "<span class='ms-status-iconSpan'><img src='/_layouts/15/images/kpinormal-1.gif'></span>";
            public static string IconError = "<span class='ms-status-iconSpan'><img src='/_layouts/15/images/kpinormal-2.gif'></span>";
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            try
            {
                TestConnectionToEntraId();

                bool testAssemblyBindingsOk = TestAssemblyBindings(Config.TenantName, Config.TenantClientId, Config.TenantClientSecret, Config.Proxy);
                if (!testAssemblyBindingsOk)
                {
                    LblTestsResult.Text += "<br/>" + Config.IconWarning + "Authentication to tenant skipped skipped since loading the dependencies failed.";
                    LblTestsResult.Text += "<br/>" + Config.IconWarning + "Search of users and groups skipped since loading the dependencies failed.";
                    LblTestsResult.Text += "<br/>" + Config.IconWarning + "Augmentation skipped since test loading the dependencies failed.";
                }
                else if (String.Equals(Config.TenantName, "TOREPLACE", StringComparison.InvariantCultureIgnoreCase))
                {
                    LblTestsResult.Text += "<br/>" + Config.IconWarning + "Authentication to tenant skipped skipped, edit this page in notepad to set the tenant and credentials.";
                    LblTestsResult.Text += "<br/>" + Config.IconWarning + "Search of users and groups skipped, edit this page in notepad to set the tenant and credentials.";
                    LblTestsResult.Text += "<br/>" + Config.IconWarning + "Augmentation skipped, edit this page in notepad to set the tenant and credentials.";
                }
                else
                {
                    EntraIDTenant tenant = TestTenantCredentials(Config.TenantName, Config.TenantClientId, Config.TenantClientSecret, Config.Proxy);
                    if (tenant == null)
                    {
                        LblTestsResult.Text += "<br/>" + Config.IconWarning + String.Format("Search of users and groups skipped (could not establish connection to tenant '{0}').", Config.TenantName);
                        LblTestsResult.Text += "<br/>" + Config.IconWarning + String.Format("Augmentation skipped (could not establish connection to tenant '{0}').", Config.TenantName);
                    }
                    else
                    {
                        ClaimsProviderSettings settings = ClaimsProviderSettings.GetDefaultSettings(Config.ClaimsProviderName);
                        settings.EntraIDTenants.Add(tenant);
                        settings.ProxyAddress = Config.Proxy;
                        EntraCP claimsProvider = new EntraCP(Config.ClaimsProviderName, settings);
                        TestClaimsProviderSearch(claimsProvider, Config.context, Config.Input);
                        TestClaimsProviderAugmentation(claimsProvider, Config.context, Config.Input);
                    }
                }
            }
            catch (Exception ex)
            {
                LblTestsResult.Text += "<br/>" + Config.IconError + String.Format("Something very unexpected happened: {0}", ex.GetType().Name + ": " + ex.Message);
            }

            ListCurrentUserClaims();
        }

        public bool TestConnectionToEntraId()
        {
            string proxyAddress = Config.Proxy;
            WebProxy proxy = String.IsNullOrWhiteSpace(proxyAddress) ? new WebProxy() : new WebProxy(proxyAddress, true);
            WebClient client = new WebClient
            {
                Proxy = proxy,
            };
            string[] urls = new string[] { "https://login.microsoftonline.com", "https://graph.microsoft.com" };
            foreach (string url in urls)
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                try
                {
                    // One difference VS EntraCP is that WebClient follows HTTP redirects, which, from URLs above, will take it to https://www.office.com/login and https://developer.microsoft.com/graph.
                    client.DownloadData(url);
                    //client.DownloadString(url);
                    timer.Stop();
                    string text = String.IsNullOrWhiteSpace(proxyAddress) ? "Connection to '{0}'" : "Connection to '{0}' through proxy '{1}'";
                    LblTestsResult.Text += "<br/>" + Config.IconSuccess + String.Format(text, url, proxyAddress, timer.ElapsedMilliseconds);
                }
                catch (Exception ex)
                {
                    timer.Stop();
                    LblTestsResult.Text += "<br/>" + Config.IconError + String.Format("Connection to '{0}' through proxy '{1}' failed after {2} ms: {3}", url, proxyAddress, timer.ElapsedMilliseconds, ex.GetType().Name + " - " + ex.Message);
                }
            }
            return true;
        }

        /// <summary>
        /// Tests assembly bindings by instanciating objects and calling methods in EntraCP that trigger the load of dependent .NET assemblies
        /// </summary>
        /// <param name="tenantName"></param>
        /// <param name="tenantClientId"></param>
        /// <param name="tenantClientSecret"></param>
        /// <param name="proxy"></param>
        /// <returns></returns>
        public bool TestAssemblyBindings(string tenantName, string tenantClientId, string tenantClientSecret, string proxy)
        {
            bool success = false;
            string errorMessage = "<br/>" + Config.IconError + "Loading of the dependencies failed, check your assembly bindings in the machine.config file. Exception: {0}";
            try
            {
                // Calling constructor of EntraIDTenant may throw FileNotFoundException on Azure.Identity
                EntraIDTenant tenant = new EntraIDTenant(tenantName);
                tenant.SetCredentials(tenantClientId, tenantClientSecret);

                // EntraIDTenant.InitializeAuthentication() will throw an exception if .NET cannot load one of the following assemblies:
                // Azure.Core.dll, System.Diagnostics.DiagnosticSource.dll, Microsoft.IdentityModel.Abstractions.dll, System.Memory.dll, System.Runtime.CompilerServices.Unsafe.dll
                tenant.InitializeAuthentication(ClaimsProviderConstants.DEFAULT_TIMEOUT, proxy);

                // EntraIDTenant.TestConnectionAsync() may throw the following exceptions:
                // System.IO.FileNotFoundException if .NET cannot load assembly Microsoft.IdentityModel.Abstractions.dll
                // Azure.Identity.AuthenticationFailedException if invalid credentials 
                Task<bool> taskTestConnection = Task.Run(async () => await tenant.TestConnectionAsync(proxy));
                taskTestConnection.Wait();
                LblTestsResult.Text += "<br/>" + Config.IconSuccess + "Loading of the dependencies";
                success = true;
            }
            catch (FileNotFoundException ex)
            {
                LblTestsResult.Text += String.Format(errorMessage, ex.Message);
            }
            // An exception in an async task is always wrapped and returned in an AggregateException
            catch (AggregateException ex)
            {
                if (ex.InnerException is FileNotFoundException)
                {
                    FileNotFoundException fnfEx = ex.InnerException as FileNotFoundException;
                    LblTestsResult.Text += String.Format(errorMessage, fnfEx.Message);
                }
                else
                {
                    // Azure.Identity.AuthenticationFailedException is expected if credentials are not valid, not an assembly load issue
                    if (String.Equals(ex.InnerException.GetType().FullName, "Azure.Identity.AuthenticationFailedException", StringComparison.InvariantCultureIgnoreCase))
                    {
                        LblTestsResult.Text += "<br/>" + Config.IconSuccess + "Loading of the dependencies";
                        success = true;
                    }
                    else
                    {
                        LblTestsResult.Text += "<br/>" + Config.IconWarning + String.Format("Loading of the dependencies might not be successful. Exception: {0}", ex.InnerException.Message);
                    }
                }
            }
            catch (TargetInvocationException ex)
            {
                if (ex.InnerException is TypeInitializationException && ex.InnerException.InnerException is FileNotFoundException)
                {
                    FileNotFoundException fnfEx = ex.InnerException.InnerException as FileNotFoundException;
                    LblTestsResult.Text += String.Format(errorMessage, fnfEx.Message);
                }
            }
            catch (Exception ex)
            {
                LblTestsResult.Text += String.Format(errorMessage, ex.GetType().Name + ": " + ex.Message);
            }
            return success;
        }

        public EntraIDTenant TestTenantCredentials(string tenantName, string tenantClientId, string tenantClientSecret, string proxy)
        {
            if (String.Equals(tenantName, "TOREPLACE", StringComparison.InvariantCultureIgnoreCase))
            {
                LblTestsResult.Text += "<br/>" + Config.IconWarning + "Authentication to tenant skipped, edit this page in notepad to set the tenant and credentials.";
                return null;
            }
            EntraIDTenant tenant = null;
            bool success = false;
            try
            {
                tenant = new EntraIDTenant(tenantName);
                tenant.SetCredentials(tenantClientId, tenantClientSecret);
                tenant.InitializeAuthentication(ClaimsProviderConstants.DEFAULT_TIMEOUT, proxy);
                Task<bool> taskTestConnection = Task.Run(async () => await tenant.TestConnectionAsync(proxy));
                taskTestConnection.Wait();
                success = taskTestConnection.Result;
                LblTestsResult.Text += "<br/>" + Config.IconSuccess + String.Format("Authentication to tenant '{0}' using client ID '{1}'", tenant.Name, tenantClientId);
            }
            // An exception in an async task is always wrapped and returned in an AggregateException
            catch (AggregateException ex)
            {
                // Azure.Identity.AuthenticationFailedException is expected if credentials are not valid
                if (String.Equals(ex.InnerException.GetType().FullName, "Azure.Identity.AuthenticationFailedException", StringComparison.InvariantCultureIgnoreCase))
                {
                    LblTestsResult.Text += "<br/>" + Config.IconError + String.Format("Authentication to tenant '{0}' using client ID '{1}' failed due to invalid credentials: {2}", tenant.Name, tenantClientId, ex.InnerException.Message + ex.InnerException.InnerException.Message);
                }
                else
                {
                    LblTestsResult.Text += "<br/>" + Config.IconError + String.Format("Authentication to tenant '{0}' using client ID '{1}' failed: {2}", tenant.Name, tenantClientId, ex.InnerException.Message);
                }
            }
            catch (Exception ex)
            {
                LblTestsResult.Text += "<br/>" + Config.IconError + String.Format("Authentication to tenant '{0}' using client ID '{1}' failed: {2}", tenant.Name, tenantClientId, ex.GetType().Name + ": " + ex.Message);
            }
            return success ? tenant : null;
        }

        public void TestClaimsProviderSearch(EntraCP claimsProvider, string context, string input)
        {
            try
            {
                var searchResult = claimsProvider.Search(new Uri(context), new[] { "User", "Group" }, input, null, 30);
                List<string> searchResultsClaimValue = new List<string>();
                if (searchResult != null)
                {
                    foreach (var children in searchResult.Children)
                    {
                        searchResultsClaimValue.AddRange(children.EntityData.Select(x => x.Claim.Value));
                    }
                }
                if (searchResultsClaimValue.Count == 0)
                {
                    LblTestsResult.Text += "<br/>" + Config.IconWarning + String.Format("Searched '{0}' in Entra ID (in context '{1}'): No result was returned.", input, context);
                }
                else
                {
                    LblTestsResult.Text += "<br/>" + Config.IconSuccess + String.Format("Searched '{0}' in Entra ID (in context '{1}'): {2} results were returned: {3}", input, context, searchResultsClaimValue.Count, String.Join(",", searchResultsClaimValue));
                }
            }
            catch (Exception ex)
            {
                LblTestsResult.Text += "<br/>" + Config.IconError + String.Format("Searching '{0}' in Entra ID (in context '{1}') failed: {2}", input, context, ex.Message);
            }
        }

        public void TestClaimsProviderAugmentation(EntraCP claimsProvider, string context, string input)
        {
            try
            {
                IdentityClaimTypeConfig idClaim = claimsProvider.Settings.ClaimTypes.UserIdentifierConfig;
                string originalIssuer = SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, Utils.GetSPTrustAssociatedWithClaimsProvider("EntraCP").Name);
                SPClaim userClaim = new SPClaim(idClaim.ClaimType, input, idClaim.ClaimValueType, originalIssuer);
                SPClaim[] groups = claimsProvider.GetClaimsForEntity(new Uri(context), userClaim);
                if (groups == null || groups.Length == 0)
                {
                    LblTestsResult.Text += "<br/>" + Config.IconWarning + String.Format("Augmentation of user with identifier '{0}' (in context '{1}'): No group was returned.", input, context);
                }
                else
                {
                    LblTestsResult.Text += "<br/>" + Config.IconSuccess + String.Format("Augmentation of user with identifier '{0}' (in context '{1}'): {2} groups were returned: {3}.", input, context, groups.Length, String.Join(",", groups.Select(x => x.Value)));
                }
            }
            catch (Exception ex)
            {
                LblTestsResult.Text += "<br/>" + Config.IconError + String.Format("Augmentation of with identifier user '{0}' (in context '{1}') failed: {2}", input, context, ex.Message);
            }
        }

        public void ListCurrentUserClaims()
        {
            try
            {
                ClaimsPrincipal claimsPrincipal = Page.User as ClaimsPrincipal;
                if (claimsPrincipal != null)
                {
                    ClaimsIdentity claimsIdentity = claimsPrincipal.Identity as ClaimsIdentity;
                    BootstrapContext bootstrapContext = claimsIdentity.BootstrapContext as BootstrapContext;
                    string sessionLifetime = bootstrapContext == null ? String.Empty : String.Format("is valid from \"{0}\" to \"{1}\" and it", bootstrapContext.SecurityToken.ValidFrom, bootstrapContext.SecurityToken.ValidTo);
                    LblCurrentUserClaims.Text += String.Format("The token of the current user \"{0}\" {1} contains {2} claims:", claimsIdentity.Name, sessionLifetime, claimsIdentity.Claims.Count());
                    foreach (Claim claim in claimsIdentity.Claims)
                    {
                        LblCurrentUserClaimsList.Text += String.Format("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>", claim.Type, claim.Value, claim.OriginalIssuer);
                    }
                }
            }
            catch (Exception ex)
            {
                LblCurrentUserClaims.Text += String.Format("Could not get the claims of the current user: {0}", ex.GetType().Name + ": " + ex.Message);
            }
        }

        protected void BtnAction_Click(object sender, EventArgs e)
        {
        }
    </script>
    <h2>Overview</h2>
    <p>
        This page helps you troubleshoot common issues with EntraCP, in particular the connectivity with Entra ID and .NET assembly binding issues. It is:
    </p>
    <ul>
        <li>Standalone: It does NOT use the EntraCP configuration.</li>
        <li>Written in inline code: You can view and edit its code using a notepad.</li>
        <li>Located in &quot;16\template\admin\EntraCP\TroubleshootEntraCP.aspx&quot;.</li>
    </ul>
    <p>Beside, EntraCP records all its activity in the SharePoint logs. You can use <a href="https://www.microsoft.com/en-us/download/details.aspx?id=44020&msockid=1428a673f6cf6d172683b376f7be6c1f" target="_blank">ULS Viewer</a> to easily inspect them (including real time monitoring). Filter on Product/Area &quot;EntraCP&quot; to only view the messages generated by EntraCP.</p>
    <h2>How-to use it</h2>
    <p>
        It may be used as-is, or you can edit it using a notepad to set valid values to connect to your tenant, and run all the tests.<br />
        You can also copy it anywhere under the folder &quot;16\TEMPLATE\LAYOUTS&quot;, to be able to call it from any SharePoint site. This can be very useful in some scenarios, for example if you want to list the claims of a user.<br />
    </p>
    <h2>Tests</h2>
    <p>
        Tests results:
        <asp:Literal ID="LblTestsResult" runat="server" Text="" />
    </p>
    <h2>Claims of the current user</h2>
    <p>
        <asp:Literal ID="LblCurrentUserClaims" runat="server" Text="" />
        <table>
            <tr>
                <th>Claim type</th>
                <th>Claim value</th>
                <th>Issuer</th>
            </tr>
            <asp:Literal ID="LblCurrentUserClaimsList" runat="server" Text="" />
        </table>
    </p>
    <%--<asp:TextBox ID="TxtUrl" runat="server" CssClass="ms-inputformcontrols" Text="URL..."></asp:TextBox>
    <br />
	<asp:Button ID="BtnAction" runat="server" Text="Boom" OnClick="BtnAction_Click" />--%>
</asp:Content>
