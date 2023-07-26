<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Page Language="C#" AutoEventWireup="true" Inherits="Microsoft.SharePoint.WebControls.LayoutsPageBase" MasterPageFile="~/_admin/admin.master" %>
<%@ Register TagPrefix="AzureCP" TagName="GlobalSettings" src="GlobalSettings.ascx" %>
<%@ Import Namespace="Yvand.ClaimsProviders.Configuration" %>
<%@ Import Namespace="System.Diagnostics" %>
<%@ Import Namespace="System.Reflection" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server" />
<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">AzureCP Configuration</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    <%= String.Format("AzureCP {0} - <a href=\"{1}\" target=\"_blank\">Visit AzureCP site</a>", ClaimsProviderConstants.ClaimsProviderVersion, ClaimsProviderConstants.PUBLICSITEURL) %>
</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <AzureCP:GlobalSettings ID="AzureCPGlobalSettings" Runat="server" ClaimsProviderName="AzureCP" PersistedObjectName="<%# ClaimsProviderConstants.CONFIGURATION_NAME %>" PersistedObjectID="<%# ClaimsProviderConstants.CONFIGURATION_ID %>" />
    </table>
</asp:Content>
