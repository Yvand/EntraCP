<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="GlobalSettings.ascx.cs" Inherits="Yvand.EntraClaimsProvider.Administration.GlobalSettingsUserControl" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Register TagPrefix="wssuc" TagName="InputFormSection" Src="~/_controltemplates/InputFormSection.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="InputFormControl" Src="~/_controltemplates/InputFormControl.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="ButtonSection" Src="~/_controltemplates/ButtonSection.ascx" %>
<%@ Register TagPrefix="wssawc" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<script type="text/javascript" src="/_layouts/15/entracp/jquery-1.9.1.min.js"></script>
<style>
    /* Maximaze space available for description text */
    .ms-inputformdescription {
        width: 100%;
    }

    /* corev15.css set it to 0.9em, which makes it too small */
    .ms-descriptiontext {
        font-size: 1em;
    }

    /* Set the size of the right part with all input controls */
    .ms-inputformcontrols {
        width: 750px;
    }

    /* Set the display of the title of each section */
    .ms-standardheader {
        color: #0072c6;
        font-weight: bold;
        font-size: 1.15em;
    }

    /* Only used in td elements in grid view that displays LDAP connections */
    .ms-vb2 {
        vertical-align: middle;
    }

    .Entracp-success {
        color: green;
        font-weight: bold;
    }

    .Entracp-HideCol {
        display: none;
    }

    .divfieldset label {
        display: inline-block;
        line-height: 1.8;
        width: 200px;
    }

    .divfieldset em {
        font-weight: bold;
        font-style: normal;
        color: #f00;
    }

    fieldset {
        border: 1px lightgray solid;
        margin: 0;
        padding: 0;
    }

        fieldset ul {
            margin: 0;
            padding: 0;
        }

        fieldset li {
            list-style: none;
            padding: 5px;
            margin: 0;
        }
</style>

<script type="text/javascript">
    // Builds unique namespace
    window.Entracp = window.Entracp || {};
    window.Entracp.EntracpSettingsPage = window.Entracp.EntracpSettingsPage || {

        Init: function () {
            // Add event handlers to preview the permission's value for both entity types, based on current settings
            // users
            $('#<%= DDLDirectoryPropertyMemberUsers.ClientID %>').on('change', function () {
                window.Entracp.EntracpSettingsPage.UpdatePermissionValuePreview("<%= DDLDirectoryPropertyMemberUsers.ClientID %>", "lblMemberPermissionValuePreview");
            });
            $('#<%= DDLDirectoryPropertyGuestUsers.ClientID %>').on('change', function () {
                window.Entracp.EntracpSettingsPage.UpdatePermissionValuePreview("<%= DDLDirectoryPropertyGuestUsers.ClientID %>", "lblGuestPermissionValuePreview");
            });

            this.UpdatePermissionValuePreview("<%= DDLDirectoryPropertyMemberUsers.ClientID %>", "lblMemberPermissionValuePreview");
            this.UpdatePermissionValuePreview("<%= DDLDirectoryPropertyGuestUsers.ClientID %>", "lblGuestPermissionValuePreview");
        },

        UpdatePermissionValuePreview: function (inputIdentifierAttributeId, lblResultId) {
            // Get the TxtGroupLdapAttribute value
            var entityPermissionValue = $("#" + inputIdentifierAttributeId + " :selected").text();

            // Set the label control to preview a group's value
            var entityPermissionValuePreview = "<" + entityPermissionValue + "_from_EntraID>";
            $("#" + lblResultId).text(entityPermissionValuePreview);
        }
    };
    // Register initialization method to run when DOM is ready and most SP JS functions loaded
    _spBodyOnLoadFunctionNames.push("window.Entracp.EntracpSettingsPage.Init");
</script>

<table width="100%" class="propertysheet" cellspacing="0" cellpadding="0" border="0">
    <tr>
        <td class="ms-descriptionText">
            <asp:Label ID="LabelMessage" runat="server" EnableViewState="False" class="ms-descriptionText" />
        </td>
    </tr>
    <tr>
        <td class="ms-error">
            <asp:Label ID="LabelErrorMessage" runat="server" EnableViewState="False" />
        </td>
    </tr>
    <tr>
        <td class="ms-descriptionText">
            <asp:ValidationSummary ID="ValSummary" HeaderText="<%$SPHtmlEncodedResources:spadmin, ValidationSummaryHeaderText%>"
                DisplayMode="BulletList" ShowSummary="True" runat="server"></asp:ValidationSummary>
        </td>
    </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
    <wssuc:ButtonSection runat="server">
        <Template_Buttons>
            <asp:Button UseSubmitBehavior="false" runat="server" class="ms-ButtonHeightWidth" OnClick="BtnOK_Click" Text="<%$Resources:wss,multipages_okbutton_text%>" ID="BtnOKTop" AccessKey="<%$Resources:wss,okbutton_accesskey%>" />
        </Template_Buttons>
    </wssuc:ButtonSection>
    <wssuc:InputFormSection Title="Registered Microsoft Entra ID tenants" runat="server">
        <Template_Description>
            <wssawc:EncodedLiteral runat="server" Text="Microsoft Entra ID tenants currently registered in EntraCP configuration." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' />
        </Template_Description>
        <Template_InputFormControls>
            <tr>
                <td>
                    <wssawc:SPGridView runat="server" ID="grdAzureTenants" AutoGenerateColumns="false" OnRowDeleting="grdAzureTenants_RowDeleting">
                        <Columns>
                            <asp:BoundField DataField="Id" ItemStyle-CssClass="Entracp-HideCol" HeaderStyle-CssClass="Entracp-HideCol" />
                            <asp:BoundField HeaderText="Tenant" DataField="TenantName" />
                            <asp:BoundField HeaderText="Application ID" DataField="ClientID" />
                            <asp:BoundField HeaderText="Authentication mode" DataField="AuthenticationMode" />
                            <asp:BoundField HeaderText="Extension Attributes Application ID" DataField="ExtensionAttributesApplicationId" />
                            <asp:CommandField HeaderText="Action" ButtonType="Button" DeleteText="Remove" ShowDeleteButton="True" />
                        </Columns>
                    </wssawc:SPGridView>
                </td>
            </tr>
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection Title="Register a new Microsoft Entra ID tenant" runat="server">
        <Template_Description>
            <wssawc:EncodedLiteral runat="server" Text="<p>EntraCP needs its own app registration to connect to your Microsoft Entra ID tenant, with permissions 'GroupMember.Read.All' and 'User.Read.All'.<br />Check <a href='https://entracp.yvand.net/docs/usage/register-application/' target='_blank'>this page</a> to see how to register it properly.<br /><br />EntraCP can authenticate using <a href='https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow#get-a-token' target='_blank'>either a secret or a certificate</a>.</p>" EncodeMethod='NoEncode' />
        </Template_Description>
        <Template_InputFormControls>
            <tr>
                <td>
                    <div class="divfieldset">
                        <fieldset>
                            <legend>Information on the Microsoft Entra ID tenant to register</legend>
                            <ul>
                                <li>
                                    <label for="<%= TxtTenantName.ClientID %>">Tenant name: <em>*</em></label>
                                    <wssawc:InputFormTextBox title="Azure tenant name" class="ms-input" ID="TxtTenantName" Columns="50" runat="server" MaxLength="255" Text="TENANTNAME.onMicrosoft.com" />
                                </li>
                                <li>
                                    <label for="<%= DDLAzureCloudInstance.ClientID %>">Cloud instance: <em>*</em></label>
                                    <asp:DropDownList runat="server" ID="DDLAzureCloudInstance" class="ms-input" Columns="50" />
                                </li>
                                <li>
                                    <label for="<%= TxtClientId.ClientID %>">Application (client) ID: <em>*</em></label>
                                    <wssawc:InputFormTextBox title="Password" class="ms-input" ID="TxtClientId" Columns="50" runat="server" MaxLength="255" />
                                </li>
                                <li>
                                    <p style="margin-bottom: 0px; margin-top: 0px">Specify either a client secret or a client certificate (but not both):</p>
                                </li>
                                <li>
                                    <label for="<%= TxtClientSecret.ClientID %>">Client secret:</label>
                                    <wssawc:InputFormTextBox title="Password" class="ms-input" ID="TxtClientSecret" Columns="50" runat="server" MaxLength="255" TextMode="Password" />
                                </li>
                                <li>
                                    <label for="<%= InputClientCertFile.ClientID %>">Client certificate (.pfx):</label>
                                    <span dir="ltr">
                                        <input id="InputClientCertFile" title="Client certificate file" runat="server" type="file" size="38" class="ms-fileinput" />
                                    </span>
                                </li>
                                <li>
                                    <label for="<%= InputClientCertPassword.ClientID %>">Client certificate password:</label>
                                    <wssawc:InputFormTextBox title="Certificate password" class="ms-input" ID="InputClientCertPassword" Columns="50" runat="server" MaxLength="255" TextMode="Password" />
                                </li>
                                <li>
                                    <label for="<%= ChkMemberUserTypeOnly.ClientID %>">Exclude <a href="https://learn.microsoft.com/en-us/entra/external-id/user-properties" target="_blank">guest users</a></label>
                                    <asp:CheckBox class="ms-input" ID="ChkMemberUserTypeOnly" runat="server" />
                                </li>
                                <li>
                                    <label for="<%= TxtExtensionAttributesApplicationId.ClientID %>">App ID for extension attributes</label>
                                    <wssawc:InputFormTextBox title="Application ID" class="ms-input" ID="TxtExtensionAttributesApplicationId" Columns="50" runat="server" MaxLength="36" />
                                </li>
                            </ul>
                            <div class="divbuttons">
                                <asp:Button runat="server" ID="BtnTestAzureTenantConnection" Text="Test connection to tenant" ToolTip="Make sure this server has access to Internet before you click" OnClick="BtnTestAzureTenantConnection_Click" class="ms-ButtonHeightWidth" />
                                <asp:Button runat="server" ID="BtnAddLdapConnection" Text="Add tenant" OnClick="BtnAddAzureTenant_Click" class="ms-ButtonHeightWidth" />
                            </div>
                            <p style="margin-left: 10px;">
                                <asp:Label ID="LabelErrorTestLdapConnection" runat="server" EnableViewState="False" class="ms-error" />
                                <asp:Label ID="LabelTestTenantConnectionOK" runat="server" EnableViewState="False" />
                            </p>
                        </fieldset>
                    </div>
                </td>
            </tr>
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection runat="server" Title="Configuration for the user identifier claim type">
        <Template_Description>
            <sharepoint:encodedliteral runat="server" text="Specify the settings to search, create and display the permissions for users." encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <br />
            <sharepoint:encodedliteral runat="server" text="Preview of an encoded permission returned by EntraCP, based on current settings:" encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <sharepoint:encodedliteral runat="server" text="- For members:" encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <b><span><%= UserIdentifierEncodedValuePrefix %><span id="lblMemberPermissionValuePreview"></span></span></b>
            <br />
            <sharepoint:encodedliteral runat="server" text="- For guests:" encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <b><span><%= UserIdentifierEncodedValuePrefix %><span id="lblGuestPermissionValuePreview"></span></span></b>
        </Template_Description>
        <Template_InputFormControls>
            <tr>
                <td colspan="2">
                    <div class="divfieldset">
                        <fieldset>
                            <legend>User identifier settings</legend>
                            <ol>
                                <li>
                                    <label>Claim type</label>
                                    <label>
                                        <wssawc:EncodedLiteral runat="server" ID="lblUserIdClaimType" EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' /></label>
                                </li>
                                <li>
                                    <label for="<%= DDLDirectoryPropertyMemberUsers.ClientID %>">Identifier for members <em>*</em></label>
                                    <asp:DropDownList runat="server" ID="DDLDirectoryPropertyMemberUsers" class="ms-input" />
                                </li>
                                <li>
                                    <label for="<%= DDLDirectoryPropertyGuestUsers.ClientID %>">Identifier for guests <em>*</em></label>
                                    <asp:DropDownList runat="server" ID="DDLDirectoryPropertyGuestUsers" class="ms-input" />
                                </li>
                                <li>
                                    <label for="<%= DDLDirectoryPropertyGuestUsers.ClientID %>" title="Property displayed in the results list in the people picker (leave blank to use the user identifier attribute)">Property as display text &#9432;</label>
                                    <asp:DropDownList runat="server" ID="DDLGraphPropertyToDisplay" class="ms-input" />
                                </li>
                            </ol>
                        </fieldset>
                    </div>
                </td>
            </tr>
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection runat="server" Title="Bypass Microsoft Entra ID lookup" Description="Skip Microsoft Entra ID lookup and consider any input as valid.<br/><br/>This can be useful to keep people picker working even if connectivity with the Azure tenant is lost.">
        <Template_InputFormControls>
            <asp:CheckBox runat="server" Name="ChkAlwaysResolveUserInput" ID="ChkAlwaysResolveUserInput" Text="Bypass Microsoft Entra ID lookup" />
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection runat="server" Title="Require exact match" Description="Enable this to return only results that match exactly the user input (case-insensitive).">
        <Template_InputFormControls>
            <asp:CheckBox runat="server" Name="ChkFilterExactMatchOnly" ID="ChkFilterExactMatchOnly" Text="Require exact match" />
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection runat="server" Title="Augmentation">
        <Template_Description>
            <wssawc:EncodedLiteral runat="server" Text="Enable augmentation to let EntraCP get " EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' />
            <a href="https://docs.microsoft.com/en-us/graph/api/user-getmembergroups" target="_blank">
                <wssawc:EncodedLiteral runat="server" Text="all the Microsoft Entra ID groups" EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' /></a>
            <wssawc:EncodedLiteral runat="server" Text="that the user is a member of.<br/><br/>If not enabled, permissions granted to Microsoft Entra ID groups may not work correctly." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' />
        </Template_Description>
        <Template_InputFormControls>
            <asp:CheckBox runat="server" Name="ChkAugmentAADRoles" ID="ChkAugmentAADRoles" Text="Retrieve Microsoft Entra ID groups" />
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection runat="server" Title="Proxy">
        <Template_Description>
            <wssawc:EncodedLiteral runat="server" Text="Configure the proxy if it is needed for EntraCP to connect to Microsoft Graph." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' />
        </Template_Description>
        <Template_InputFormControls>
            <label for="<%= InputProxyAddress.ClientID %>">Proxy address:</label><br />
            <wssawc:InputFormTextBox title="Proxy address" class="ms-input" ID="InputProxyAddress" Columns="50" runat="server" />
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection runat="server" Title="Type of groups">
        <Template_Description>
            <wssawc:EncodedLiteral runat="server" Text="Set if all " EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' />
            <a href="https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0" target="_blank">
                <wssawc:EncodedLiteral runat="server" Text="type of groups" EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' /></a>
            <wssawc:EncodedLiteral runat="server" Text="should be returned, including Office 365 unified groups, or only those that are security-enabled." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' />
        </Template_Description>
        <Template_InputFormControls>
            <asp:CheckBox runat="server" Name="ChkFilterSecurityEnabledGroupsOnly" ID="ChkFilterSecurityEnabledGroupsOnly" Text="Return <a href='https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0' target='_blank'>security-enabled</a> groups only" />
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection runat="server" Title="Reset EntraCP configuration" Description="Restore configuration to its default values. All changes, including in claim types mappings, will be lost.">
        <Template_InputFormControls>
            <asp:Button runat="server" ID="BtnResetConfig" Text="Reset EntraCP configuration" OnClick="BtnResetConfig_Click" class="ms-ButtonHeightWidth" OnClientClick="return confirm('Do you really want to reset EntraCP configuration?');" />
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:ButtonSection runat="server">
        <Template_Buttons>
            <asp:Button UseSubmitBehavior="false" runat="server" class="ms-ButtonHeightWidth" OnClick="BtnOK_Click" Text="<%$Resources:wss,multipages_okbutton_text%>" ID="BtnOK" AccessKey="<%$Resources:wss,okbutton_accesskey%>" />
        </Template_Buttons>
    </wssuc:ButtonSection>
</table>
