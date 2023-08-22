<%@ Page Language="C#" AutoEventWireup="true" Inherits="Microsoft.SharePoint.WebControls.LayoutsPageBase" DynamicMasterPageFile="~masterurl/default.master" %>

<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="System.Reflection" %>
<%@ Import Namespace="Yvand.ClaimsProviders" %>
<%@ Import Namespace="Yvand.ClaimsProviders.Config" %>
<%@ Import Namespace="Microsoft.SharePoint.Administration" %>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">Validate the installation of AzureCP SE</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">Verify if AzureCP SE can run</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <%
        string url = "https://login.microsoftonline.com";
        LblUrl.Text = url;
        try
        {
            // Ensure that Azure.Core.dll can be loaded
            AzureTenant tenant = new AzureTenant();
            tenant.Name = "fake";
            tenant.ClientId = "fake";
            tenant.ClientSecret = "fake";
            // This call will tests all assemblies, with the unfortunate exception of "Microsoft.IdentityModel.Abstractions"
            tenant.InitializeAuthentication(10000, String.Empty);
            LblResult.Text = "Loading of assemblies OK";

            //string claimsProviderName = "AzureCPSE";
            //IAzureCPSettings entityProviderSettings = new AzureCPSettings
            //{
            //    Version = -1,
            //    AzureTenants = new List<AzureTenant> { tenant },
            //    ClaimTypes = AADEntityProviderSettings.ReturnDefaultClaimTypesConfig(claimsProviderName),
            //};
            //AzureCP claimsProvider = new AzureCP(claimsProviderName);
            //claimsProvider.GetType().GetProperty("Settings").SetMethod.Invoke(claimsProvider, new object[] { entityProviderSettings });
            //claimsProvider.GetType().GetMethod("InitializeInternalRuntimeSettings", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(claimsProvider, new object[] { });
            ////claimsProvider.ValidateSettings(null);
            //Uri uri = Microsoft.SharePoint.Administration.SPAdministrationWebApplication.Local.GetResponseUri(0);
            //claimsProvider.Search(uri, new string[] { "User" }, "fake", null, 30);
        }
        catch (FileNotFoundException ex)
        {
            LblResult.Text = String.Format(".NET could not load an assembly, please check your assembly bindings in machine.config file, or .config file for current process. Exception details: {0}", ex.Message);
        }
        catch (Exception ex)
        {
            LblResult.Text = ex.Message;
        }
    %>
    StatusCode of the connection to "<asp:Literal ID="LblUrl" runat="server" Text="" />":
    <asp:Literal ID="LblResult" runat="server" Text="" />
    <%--<asp:TextBox ID="TxtUrl" runat="server" CssClass="ms-inputformcontrols" Text="URL..."></asp:TextBox>--%>
</asp:Content>
