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

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">Troubleshoot EntraCP</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">Troubleshoot EntraCP in the context of current site</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <script runat="server" language="C#">
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // Edit the variables below with your own values
            string tenantName = "ReplaceWithYourOwnValue";
            string tenantClientId = "ReplaceWithYourOwnValue";
            string tenantClientSecret = "ReplaceWithYourOwnValue";
            string proxy = "";
            string input = "yvand";
            string context = SPContext.Current.Web.Url;

            TestConnectionToEntraId(proxy);

            EntraIDTenant tenant = TestTenantConnectionAndAssemblyBindings(tenantName, tenantClientId, tenantClientSecret, proxy);
            if (tenant == null)
            {
                return;
            }

            string claimsProviderName = "EntraCP";
            EntraCPSettings settings = EntraCPSettings.GetDefaultSettings(claimsProviderName);
            settings.EntraIDTenants.Add(tenant);
            settings.ProxyAddress = proxy;
            EntraCP claimsProvider = new EntraCP(claimsProviderName, settings);
            TestClaimsProviderSearch(claimsProvider, context, input);
            TestClaimsProviderAugmentation(claimsProvider, context, input);
        }

        public bool TestConnectionToEntraId(string proxyAddress)
        {
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
                    LblResult.Text += String.Format("<br/>Test connection to '{0}' through proxy '{1}': OK, took {2} ms.", url, proxyAddress, timer.ElapsedMilliseconds);
                }
                catch (Exception ex)
                {
                    timer.Stop();
                    LblResult.Text += String.Format("<br/>Test connection to '{0}' through proxy '{1}' failed after {2} ms: {3}", url, proxyAddress, timer.ElapsedMilliseconds, ex.GetType().Name + " - " + ex.Message);
                }
            }
            return true;
        }

        public EntraIDTenant TestTenantConnectionAndAssemblyBindings(string tenantName, string tenantClientId, string tenantClientSecret, string proxy)
        {
            EntraIDTenant tenant = null;
            bool success = false;
            try
            {
                // Calling constructor of EntraIDTenant may throw FileNotFoundException on Azure.Identity
                tenant = new EntraIDTenant(tenantName);
                tenant.SetCredentials(tenantClientId, tenantClientSecret);

                // EntraIDTenant.InitializeAuthentication() will throw an exception if .NET cannot load one of the following assemblies:
                // Azure.Core.dll, System.Diagnostics.DiagnosticSource.dll, Microsoft.IdentityModel.Abstractions.dll, System.Memory.dll, System.Runtime.CompilerServices.Unsafe.dll
                tenant.InitializeAuthentication(ClaimsProviderConstants.DEFAULT_TIMEOUT, proxy);

                // EntraIDTenant.TestConnectionAsync() will throw exceptions:
                // if .NET cannot load assembly Microsoft.IdentityModel.Abstractions.dll
                // Azure.Identity.AuthenticationFailedException if invalid credentials 
                Task<bool> taskTestConnection = Task.Run(async () => await tenant.TestConnectionAsync(proxy));
                taskTestConnection.Wait();
                success = taskTestConnection.Result;
                LblResult.Text += String.Format("<br/>Test loading of dependencies: OK");
                LblResult.Text += String.Format("<br/>Test connection to tenant '{0}': {1}", tenant.Name, success ? "OK" : "Failed");
            }
            catch (FileNotFoundException ex)
            {
                LblResult.Text += String.Format("<br/>Test loading of dependencies: Failed. Check your assembly bindings in the machine.config file. Exception: '{0}'", ex.Message);
            }
            // An exception in an async task is always wrapped and returned in an AggregateException
            catch (AggregateException ex)
            {
                if (ex.InnerException is FileNotFoundException)
                {
                    LblResult.Text += String.Format("<br/>Test loading of dependencies: Failed. Check your assembly bindings in the machine.config file. Exception: '{0}'", ex.InnerException.Message);
                }
                else
                {
                    LblResult.Text += String.Format("<br/>Test loading of dependencies: OK");
                    // Azure.Identity.AuthenticationFailedException is expected if credentials are not valid
                    if (String.Equals(ex.InnerException.GetType().FullName, "Azure.Identity.AuthenticationFailedException", StringComparison.InvariantCultureIgnoreCase))
                    {
                        LblResult.Text += String.Format("<br/>Test connection to tenant '{0}' failed due to invalid credentials: {1}", tenant.Name, ex.InnerException.Message);
                    }
                    else
                    {
                        LblResult.Text += String.Format("<br/>Test connection to tenant '{0}' failed for an unknown reason: {1}", tenant.Name, ex.InnerException.GetType().Name + " - " + ex.InnerException.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                LblResult.Text += String.Format("<br/>Unexpected error {0}: {1}", ex.GetType().Name, ex.Message);
            }
            return success ? tenant : null;
        }

        public bool TestClaimsProviderSearch(EntraCP claimsProvider, string context, string input)
        {
            try
            {
                var searchResult = claimsProvider.Search(new Uri(context), new[] { "User", "Group" }, input, null, 30);
                int searchResultCount = 0;
                if (searchResult != null)
                {
                    foreach (var children in searchResult.Children)
                    {
                        searchResultCount += children.EntityData.Count;
                    }
                }
                LblResult.Text += String.Format("<br/>Test search with input '{0}' on '{1}': OK with {2} results returned.", input, context, searchResultCount);
                return true;
            }
            catch (Exception ex)
            {
                LblResult.Text += String.Format("<br/>Test search with input '{0}' on '{1}': Failed: {2}", input, context, ex.Message);
            }
            return false;
        }

        public bool TestClaimsProviderAugmentation(EntraCP claimsProvider, string context, string input)
        {
            try
            {
                IdentityClaimTypeConfig idClaim = claimsProvider.Settings.ClaimTypes.IdentityClaim;
                string originalIssuer = SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, Utils.GetSPTrustAssociatedWithClaimsProvider("EntraCP").Name);
                SPClaim claim = new SPClaim(idClaim.ClaimType, input, idClaim.ClaimValueType, originalIssuer);
                SPClaim[] groups = claimsProvider.GetClaimsForEntity(new Uri(context), claim);
                LblResult.Text += String.Format("<br/>Test augmentation for user '{0}' on '{1}': OK with {2} groups returned.", input, context, groups == null ? 0 : groups.Length);
                return true;
            }
            catch (Exception ex)
            {
                LblResult.Text += String.Format("<br/>Test augmentation for user '{0}' on '{1}': Failed: {2}", input, context, ex.Message);
            }
            return false;
        }
    </script>
    This page helps you troubleshoot EntraCP with minimal overhead, directly in the context of SharePoint sites.<br />
    It is written entirely using inline code, so you can easily customize it (and set valid credentials).<br />
    For security reasons, by default it can only be called from the central administration, but you can simply copy it in the LAYOUTS folder, to call it from any SharePoint web application.<br />
    <br />
    <asp:Literal ID="LblResult" runat="server" Text="" />
    <%--<asp:TextBox ID="TxtUrl" runat="server" CssClass="ms-inputformcontrols" Text="URL..."></asp:TextBox>--%>
</asp:Content>
