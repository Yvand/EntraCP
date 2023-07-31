<%@ Page Language="C#" AutoEventWireup="true" Inherits="Microsoft.SharePoint.WebControls.LayoutsPageBase" MasterPageFile="~/_admin/admin.master" %>
<%@ Register TagPrefix="AzureCP" TagName="ClaimTypesConfigUC" src="ClaimTypesConfig.ascx" %>
<%@ Import Namespace="Yvand.ClaimsProviders.Configuration" %>
<%@ Import Namespace="Yvand.ClaimsProviders" %>
<%@ Import Namespace="System.Diagnostics" %>
<%@ Import Namespace="System.Reflection" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
</asp:Content>
<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
    Claim types configuration for AzureCPSE
</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    <%= String.Format("<a href=\"{1}\" target=\"_blank\">AzureCP Subscription Edition</a> {0}", ClaimsProviderConstants.ClaimsProviderVersion, ClaimsProviderConstants.PUBLICSITEURL) %>
</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <AzureCP:ClaimTypesConfigUC ID="ClaimsListConfiguration" Runat="server" ClaimsProviderName="<%# AzureCPSE.ClaimsProviderName %>" ConfigurationName="<%# ClaimsProviderConstants.CONFIGURATION_NAME %>" ConfigurationID="<%# ClaimsProviderConstants.CONFIGURATION_ID %>" />
    </table>
</asp:Content>
