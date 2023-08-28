<%@ Page Language="C#" AutoEventWireup="true" Inherits="Microsoft.SharePoint.WebControls.LayoutsPageBase" DynamicMasterPageFile="~masterurl/default.master" %>

<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Reflection" %>
<%@ Import Namespace="System.Threading.Tasks" %>
<%@ Import Namespace="Yvand" %>
<%@ Import Namespace="Yvand.Config" %>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">Troubleshoot AzureCP SE</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">Troubleshoot AzureCP SE in the context of current site</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <script runat="server">
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            // Variables proxy and tenant below can be replaced with actual values
            string proxy = String.Empty;
            AzureTenant tenant = new AzureTenant
            {
                Name = "IDoNotExist",
                ClientId = "IDoNotExist",
                ClientSecret = "IDoNotExist",
            };
            try
            {
                // The call to AzureTenant.InitializeAuthentication() tests if .NET can load the following assemblies:
                // Azure.Core.dll
                // System.Diagnostics.DiagnosticSource.dll
                // Microsoft.IdentityModel.Abstractions.dll
                // System.Memory.dll
                // System.Runtime.CompilerServices.Unsafe.dll
                tenant.InitializeAuthentication(10000, String.Empty);

                // The call to AzureTenant.TestConnectionAsync() tests if .NET can load the following assembly: Microsoft.IdentityModel.Abstractions.dll
                Task<bool> taskTestConnection = Task.Run(async () => await tenant.TestConnectionAsync(proxy));
                // If no valid credentials are set, this should throw an Azure.Identity.AuthenticationFailedException
                taskTestConnection.Wait();
                bool success = taskTestConnection.Result;
                LblResult.Text = String.Format("Loading of all dependent assemblies was successful. Connection to tenant successful: {0}", success);
            }
            catch (FileNotFoundException ex)
            {
                LblResult.Text = String.Format(".NET could not load an assembly, please check your assembly bindings in machine.config file, or .config file for current process. Exception details: {0}", ex.Message);
            }
            // An exception in an async task is always wrapped and returned in an AggregateException
            catch (AggregateException ex)
            {
                if (ex.InnerException is FileNotFoundException)
                {
                    LblResult.Text = String.Format(".NET could not load an assembly, please check your assembly bindings in machine.config file, or .config file for current process. Exception details: {0}", ex.InnerException.Message);
                }
                else
                {
                    // Azure.Identity.AuthenticationFailedException is expected if credentials are not valid
                    if (String.Equals(ex.InnerException.GetType().FullName, "Azure.Identity.AuthenticationFailedException", StringComparison.InvariantCultureIgnoreCase))
                    {
                        LblResult.Text = String.Format("Loading of all dependent assemblies was successful.<br />Authentication to the tenant '{0}' failed: {1}", tenant.Name, ex.InnerException.Message);
                    }
                    else
                    {
                        LblResult.Text = ex.InnerException.Message;
                    }
                }
            }
            catch (Exception ex)
            {
                LblResult.Text = String.Format("Unexpected exception {0}: {1}", ex.GetType().Name, ex.Message);
            }
        }
    </script>
    This page primarily verifies if the .NET assembly bindings are correctly set.<br />
    It also tries to connect to Azure AD, but it uses hardcoded, fake credentials by default.<br />
    It has no code behind, it is written entirely using inline code, so you can easily customize it (and set valid credentials).<br />
    For security reasons, by default it can only be called from the central administration, but you can simply copy it in the LAYOUTS folder, to call it from any SharePoint web application.<br />
    <br />
    Result of the tests:<br />
    <asp:Literal ID="LblResult" runat="server" Text="" />
    <%--<asp:TextBox ID="TxtUrl" runat="server" CssClass="ms-inputformcontrols" Text="URL..."></asp:TextBox>--%>
</asp:Content>
