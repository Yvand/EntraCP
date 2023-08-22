<%@ Page Language="C#" AutoEventWireup="true" Inherits="Microsoft.SharePoint.WebControls.LayoutsPageBase" DynamicMasterPageFile="~masterurl/default.master" %>

<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Reflection" %>
<%@ Import Namespace="System.Threading.Tasks" %>
<%@ Import Namespace="Yvand.ClaimsProviders" %>
<%@ Import Namespace="Yvand.ClaimsProviders.Config" %>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">Validate the installation of AzureCP SE</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">Verify if AzureCP SE can run</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <%
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
                    LblResult.Text = String.Format("Loading of all dependent assemblies was successful. Authentication to the tenant '{0}' failed: {1}", tenant.Name, ex.InnerException.Message);
                }
                else
                {
                    LblResult.Text = ex.InnerException.Message;
                }
            }
        }
        catch (Exception ex)
        {
            LblResult.Text = ex.Message;
            LblResult.Text = ex.GetType().Name;
        }
    %>
    Result:<br />
    <asp:Literal ID="LblResult" runat="server" Text="" />
    <%--<asp:TextBox ID="TxtUrl" runat="server" CssClass="ms-inputformcontrols" Text="URL..."></asp:TextBox>--%>
</asp:Content>
