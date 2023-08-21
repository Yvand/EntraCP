<%@ Page Language="C#" AutoEventWireup="true" Inherits="Microsoft.SharePoint.WebControls.LayoutsPageBase" DynamicMasterPageFile="~masterurl/default.master" %>
<%@ Assembly Name="AzureCPSE, Version=1.0.0.0, Culture=neutral, PublicKeyToken=65dc6b5903b51636" %>
<%@ Import Namespace="System" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Net" %>
<%@ Import Namespace="Yvand.ClaimsProviders.Config" %>

<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">Validate the installation of AzureCP SE</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server" >
Ensure AzureCP SE can run
</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
<%
    string url = "https://login.microsoftonline.com";
    LblUrl.Text = url;
    try
    {
        // Ensure that Azure.Core.dll can be loaded
	    AzureTenant tenant = new AzureTenant();
	    tenant.InitializeAuthentication(10000, String.Empty);
	    LblResult.Text = "Load of Azure.Core OK";
    }
    catch (FileNotFoundException ex)
    {
        LblResult.Text = String.Format("An assembly binding seems to be missing in .config file: {0}", ex.Message);
    }
    catch (Exception ex)
    {
        LblResult.Text = ex.Message;
    }
%>
    StatusCode of the connection to "<asp:Literal ID="LblUrl" runat="server" Text="" />": <asp:Literal ID="LblResult" runat="server" Text="" />
    <%--<asp:TextBox ID="TxtUrl" runat="server" CssClass="ms-inputformcontrols" Text="URL..."></asp:TextBox>--%>
</asp:Content>
