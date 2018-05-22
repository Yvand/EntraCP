<%@ Page Language="C#" AutoEventWireup="true" Inherits="Microsoft.SharePoint.WebControls.LayoutsPageBase" MasterPageFile="~/_admin/admin.master" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Register TagPrefix="AzureCP" TagName="ClaimTypesConfigUC" src="ClaimTypesConfig.ascx" %>
<%@ Import Namespace="azurecp" %>
<%@ Import Namespace="System.Diagnostics" %>
<%@ Import Namespace="System.Reflection" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
</asp:Content>
<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
    Claim types configuration for AzureCP
</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea" runat="server">
    <%= String.Format("AzureCP v{0} - <a href=\"{1}\" target=\"_blank\">Open AzureCP site</a>", FileVersionInfo.GetVersionInfo(Assembly.GetAssembly(typeof(AzureCP)).Location).FileVersion, ClaimsProviderConstants.PUBLICSITEURL) %>
</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <table border="0" cellspacing="0" cellpadding="0" width="100%">
        <AzureCP:ClaimTypesConfigUC ID="LdapcpClaimsList" Runat="server" ClaimsProviderName="AzureCP" PersistedObjectName="<%# ClaimsProviderConstants.AZURECPCONFIG_NAME %>" PersistedObjectID="<%# ClaimsProviderConstants.AZURECPCONFIG_ID %>" />
    </table>
</asp:Content>
