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
        width: 210px;
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
            $('#<%= DdlUserIdDirectoryPropertyMembers.ClientID %>').on('change', function () {
                window.Entracp.EntracpSettingsPage.UpdatePermissionValuePreview("<%= DdlUserIdDirectoryPropertyMembers.ClientID %>", "lblMemberPermissionValuePreview");
            });
            $('#<%= DdlUserIdDirectoryPropertyGuests.ClientID %>').on('change', function () {
                window.Entracp.EntracpSettingsPage.UpdatePermissionValuePreview("<%= DdlUserIdDirectoryPropertyGuests.ClientID %>", "lblGuestPermissionValuePreview");
            });

            // Groups
            $('#<%= DdlGroupDirectoryProperty.ClientID %>').on('change', function () {
                window.Entracp.EntracpSettingsPage.UpdatePermissionValuePreview("<%= DdlGroupDirectoryProperty.ClientID %>", "lblGroupsPermissionValuePreview");
            });

            // Set the initial value
            this.UpdatePermissionValuePreview("<%= DdlUserIdDirectoryPropertyMembers.ClientID %>", "lblMemberPermissionValuePreview");
            this.UpdatePermissionValuePreview("<%= DdlUserIdDirectoryPropertyGuests.ClientID %>", "lblGuestPermissionValuePreview");
            this.UpdatePermissionValuePreview("<%= DdlGroupDirectoryProperty.ClientID %>", "lblGroupsPermissionValuePreview");
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
            <wssawc:EncodedLiteral runat="server" Text="<p>EntraCP needs its own app registration to connect to your Microsoft Entra ID tenant, with permissions 'GroupMember.Read.All' and 'User.Read.All'.<br />Read <a href='https://entracp.yvand.net/overview/register-application/' target='_blank'>this article</a> to learn how to register the app in your tenant.<br /><br />EntraCP can authenticate using <a href='https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow#get-a-token' target='_blank'>either a secret or a certificate</a>.</p>" EncodeMethod='NoEncode' />
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
    <wssuc:InputFormSection runat="server" Title="Configuration for the user identifier">
        <Template_Description>
            <sharepoint:encodedliteral runat="server" text="Specify the settings to search, create and display the permissions for users." encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <br />
            <sharepoint:encodedliteral runat="server" text="Preview of an encoded permission, based on current settings:" encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
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
                                    <label for="<%= DdlUserIdDirectoryPropertyMembers.ClientID %>">Identifier for members <em>*</em></label>
                                    <asp:DropDownList runat="server" ID="DdlUserIdDirectoryPropertyMembers" class="ms-input" />
                                </li>
                                <li>
                                    <label for="<%= DdlUserIdDirectoryPropertyGuests.ClientID %>">Identifier for guests <em>*</em></label>
                                    <asp:DropDownList runat="server" ID="DdlUserIdDirectoryPropertyGuests" class="ms-input" />
                                </li>
                                <li>
                                    <label for="<%= DdlUserGraphPropertyToDisplay.ClientID %>" title="Property displayed in the results list in the people picker (leave blank to use the user identifier attribute)">Property as display text &#9432;</label>
                                    <asp:DropDownList runat="server" ID="DdlUserGraphPropertyToDisplay" class="ms-input" />
                                </li>
                            </ol>
                        </fieldset>
                    </div>
                </td>
            </tr>
        </Template_InputFormControls>
    </wssuc:InputFormSection>

    <wssuc:InputFormSection runat="server" Title="Configuration for the group identifier">
        <Template_Description>
            <sharepoint:encodedliteral runat="server" text="Specify the settings to search, create and display the permissions for groups." encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <br />
            <sharepoint:encodedliteral runat="server" text="Preview of an encoded permission, based on current settings:" encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <b><span><%= GroupIdentifierEncodedValuePrefix %><span id="lblGroupsPermissionValuePreview"></span></span></b>
            <br />
            <br />
            <sharepoint:encodedliteral runat="server" text="- Augmentation: If enabled, EntraCP gets the group membership of the users when they sign-in, or whenever SharePoint asks for it. If not enabled, permissions granted to Microsoft Entra ID groups may not work." encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
        </Template_Description>
        <Template_InputFormControls>
            <tr>
                <td colspan="2">
                    <div class="divfieldset">
                        <fieldset>
                            <legend>Group identifier settings</legend>
                            <ol>
                                <li>
                                    <label title="This liste is based on the claim types registered in your SharePoint trust">
                                        <wssawc:EncodedLiteral runat="server" Text="Claim type &#9432;" EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' /><em>*</em></label>
                                    <asp:DropDownList ID="DdlGroupClaimType" runat="server" class="ms-input">
                                        <asp:ListItem Selected="True" Value="None"></asp:ListItem>
                                    </asp:DropDownList>
                                </li>
                                <li>
                                    <label for="<%= DdlGroupDirectoryProperty.ClientID %>">Identifier property <em>*</em></label>
                                    <asp:DropDownList runat="server" ID="DdlGroupDirectoryProperty" class="ms-input" />
                                </li>
                                <li>
                                    <label for="<%= DdlGroupGraphPropertyToDisplay.ClientID %>" title="Property displayed in the results list in the people picker (leave blank to use the group identifier attribute)">Property as display text &#9432;</label>
                                    <asp:DropDownList runat="server" ID="DdlGroupGraphPropertyToDisplay" class="ms-input" />
                                </li>
                                <li>
                                    <label for="<%= ChkFilterSecurityEnabledGroupsOnly.ClientID %>">Enable augmentation</label>
                                    <asp:CheckBox runat="server" Name="ChkAugmentAADRoles" ID="ChkAugmentAADRoles" />
                                </li>
                                <li>
                                    <label for="<%= ChkFilterSecurityEnabledGroupsOnly.ClientID %>">Return only <a href='https://learn.microsoft.com/en-us/graph/api/resources/groups-overview' target='_blank'>security-enabled groups</a></label>
                                    <asp:CheckBox runat="server" Name="ChkFilterSecurityEnabledGroupsOnly" ID="ChkFilterSecurityEnabledGroupsOnly" />
                                </li>
                            </ol>
                        </fieldset>
                    </div>
                </td>
            </tr>
        </Template_InputFormControls>
    </wssuc:InputFormSection>

    <wssuc:InputFormSection runat="server" Title="Bypass Entra ID">
        <Template_Description>
            <sharepoint:encodedliteral runat="server" text="Bypass the Entra ID tenant(s) registered and, depending on the context:" encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <sharepoint:encodedliteral runat="server" text="- Search: Uses the input as the claim's value, and return 1 entity per claim type." encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <sharepoint:encodedliteral runat="server" text="- Validation: Validates the incoming entity as-is." encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <sharepoint:encodedliteral runat="server" text="This setting does not affect the augmentation." encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <br />
            <sharepoint:encodedliteral runat="server" text="It can be used as a mitigation if one or more SharePoint server(s) lost the connection with your Entra ID tenant(s), until it is restored." encodemethod='HtmlEncodeAllowSimpleTextFormatting' />
        </Template_Description>
        <Template_InputFormControls>
            <asp:CheckBox runat="server" Name="ChkAlwaysResolveUserInput" ID="ChkAlwaysResolveUserInput" Text="Bypass the Entra ID tenant(s) registered" />
        </Template_InputFormControls>
    </wssuc:InputFormSection>
    <wssuc:InputFormSection runat="server" Title="Require exact match" Description="Enable this to return only results that match exactly the user input (case-insensitive).">
        <Template_InputFormControls>
            <asp:CheckBox runat="server" Name="ChkFilterExactMatchOnly" ID="ChkFilterExactMatchOnly" Text="Require exact match when typing in the people picker" />
        </Template_InputFormControls>
    </wssuc:InputFormSection>

    <wssuc:InputFormSection runat="server" Title="Proxy">
        <Template_Description>
            <wssawc:EncodedLiteral runat="server" Text="Configure the proxy if it is needed for EntraCP to connect to Microsoft Graph." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting' />
            <br />
            <br />
            <wssawc:EncodedLiteral runat="server" Text="Additional configuration in Windows may still be required. Read <a href='https://entracp.yvand.net/docs/configure-the-proxy/' target='_blank'>this article</a> to fully configure the proxy." EncodeMethod='NoEncode' />
        </Template_Description>
        <Template_InputFormControls>
            <label for="<%= InputProxyAddress.ClientID %>">Proxy address:</label><br />
            <wssawc:InputFormTextBox title="Proxy address" class="ms-input" ID="InputProxyAddress" Columns="50" runat="server" />
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
