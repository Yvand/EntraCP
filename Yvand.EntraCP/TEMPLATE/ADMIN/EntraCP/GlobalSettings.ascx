<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="GlobalSettings.ascx.cs" Inherits="Yvand.Administration.GlobalSettingsUserControl" %>
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

    #divNewLdapConnection label {
        display: inline-block;
        line-height: 1.8;
        width: 250px;
    }

    #divUserIdentifiers label {
        display: inline-block;
        line-height: 1.8;
        width: 250px;
    }

    fieldset {
        border: 1;
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

    #divNewLdapConnection em {
        font-weight: bold;
        font-style: normal;
        color: #f00;
    }
</style>
<script type="text/javascript">
    // Builds unique namespace
    window.Entracp = window.Entracp || {};
    window.Entracp.EntracpSettingsPage = window.Entracp.EntracpSettingsPage || {};

    // Identity permission section
    window.Entracp.EntracpSettingsPage.CheckRbIdentityCustomGraphProperty = function () {
        var control = (document.getElementById("<%= RbIdentityCustomGraphProperty.ClientID %>"));
        if (control != null) {
            control.checked = true;
        }
    }

    // Register initialization method to run when DOM is ready and most SP JS functions loaded
    _spBodyOnLoadFunctionNames.push("window.Entracp.EntracpSettingsPage.Init");

    window.Entracp.EntracpSettingsPage.Init = function () {
        // Variables initialized from server side code

    }
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
    <wssuc:buttonsection runat="server">
        <template_buttons>
			<asp:Button UseSubmitBehavior="false" runat="server" class="ms-ButtonHeightWidth" OnClick="BtnOK_Click" Text="<%$Resources:wss,multipages_okbutton_text%>" id="BtnOKTop" accesskey="<%$Resources:wss,okbutton_accesskey%>"/>
		</template_buttons>
    </wssuc:buttonsection>
    <wssuc:inputformsection title="Registered Microsoft Entra ID tenants" runat="server">
        <template_description>
				<wssawc:EncodedLiteral runat="server" text="Microsoft Entra ID tenants currently registered in EntraCP configuration." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
			</template_description>
        <template_inputformcontrols>
			<tr><td>
			<wssawc:SPGridView runat="server" ID="grdAzureTenants" AutoGenerateColumns="false" OnRowDeleting="grdAzureTenants_RowDeleting">
				<Columns>
					<asp:BoundField DataField="Id" ItemStyle-CssClass="Entracp-HideCol" HeaderStyle-CssClass="Entracp-HideCol"/>
					<asp:BoundField HeaderText="Tenant" DataField="TenantName"/>
					<asp:BoundField HeaderText="Application ID" DataField="ClientID"/>
                    <asp:BoundField HeaderText="Authentication mode" DataField="AuthenticationMode" />
                    <asp:BoundField HeaderText="Extension Attributes Application ID" DataField="ExtensionAttributesApplicationId" />
					<asp:CommandField HeaderText="Action" ButtonType="Button" DeleteText="Remove" ShowDeleteButton="True" />
				</Columns>
			</wssawc:SPGridView>
			</td></tr>
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection title="Register a new Microsoft Entra ID tenant" runat="server">
        <template_description>
			<wssawc:EncodedLiteral runat="server" text="<p>EntraCP needs its own app registration to connect to your Microsoft Entra ID tenant, with permissions 'GroupMember.Read.All' and 'User.Read.All'.<br />Check <a href='https://entracp.yvand.net/docs/usage/register-application/' target='_blank'>this page</a> to see how to register it properly.<br /><br />EntraCP can authenticate using <a href='https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow#get-a-token' target='_blank'>either a secret or a certificate</a>.</p>" EncodeMethod='NoEncode' />
		</template_description>
        <template_inputformcontrols>
			<tr><td>
				<div id="divNewLdapConnection">
				<fieldset>
				<legend>Information on the Microsoft Entra ID tenant to register</legend>
				<ul>
					<li>
						<label for="<%= TxtTenantName.ClientID %>">Tenant name: <em>*</em></label>
						<wssawc:InputFormTextBox title="Azure tenant name" class="ms-input" ID="TxtTenantName" Columns="50" Runat="server" MaxLength="255" Text="TENANTNAME.onMicrosoft.com" />
					</li>
                    <li>
                        <label for="<%= DDLAzureCloudInstance.ClientID %>">Cloud instance: <em>*</em></label>
                        <asp:DropDownList runat="server" ID="DDLAzureCloudInstance" class="ms-input" Columns="50" />
                    </li>
					<li>
						<label for="<%= TxtClientId.ClientID %>">Application (client) ID: <em>*</em></label>
						<wssawc:InputFormTextBox title="Password" class="ms-input" ID="TxtClientId" Columns="50" Runat="server" MaxLength="255" />
					</li>
					<li>
						<p style="margin-bottom: 0px; margin-top: 0px">Specify either a client secret or a client certificate (but not both):</p>
					</li>
					<li>
						<label for="<%= TxtClientSecret.ClientID %>">Client secret:</label>
						<wssawc:InputFormTextBox title="Password" class="ms-input" ID="TxtClientSecret" Columns="50" Runat="server" MaxLength="255" TextMode="Password" />
					</li>
                    <li>
						<label for="<%= InputClientCertFile.ClientID %>">Client certificate (.pfx):</label>
                        <span dir="ltr">
					        <input id="InputClientCertFile" title="Client certificate file" runat="server" type="file" size="38" class="ms-fileinput" />
				        </span>
                    </li>
					<li>
						<label for="<%= InputClientCertPassword.ClientID %>">Client certificate password:</label>
						<wssawc:InputFormTextBox title="Certificate password" class="ms-input" ID="InputClientCertPassword" Columns="50" Runat="server" MaxLength="255" TextMode="Password" />
					</li>
					<li>
						<label for="<%= ChkMemberUserTypeOnly.ClientID %>">Filter out <a href="https://docs.microsoft.com/en-us/azure/active-directory/active-directory-b2b-user-properties" target="_blank">Guest users</a> on this tenant:</label>
						<table border="0" cellpadding="0" cellspacing="0" style="display: inline;">
						<wssawc:InputFormCheckBox class="ms-input" ID="ChkMemberUserTypeOnly" ToolTip="Filter out Guest users" runat="server" />
						</table>
					</li>
                    <li>
                        <label for="<%= TxtExtensionAttributesApplicationId.ClientID %>">Application ID for extension attributes</label>
			            <wssawc:InputFormTextBox title="Application ID" class="ms-input" ID="TxtExtensionAttributesApplicationId" Columns="50" Runat="server" MaxLength="36" />
                    </li>
				</ul>
				<div class="divbuttons">
					<asp:Button runat="server" ID="BtnTestAzureTenantConnection" Text="Test connection to tenant" ToolTip="Make sure this server has access to Internet before you click" onclick="BtnTestAzureTenantConnection_Click" class="ms-ButtonHeightWidth" />
					<asp:Button runat="server" ID="BtnAddLdapConnection" Text="Add tenant" OnClick="BtnAddAzureTenant_Click" class="ms-ButtonHeightWidth" />
				</div>
				<p style="margin-left: 10px;">
					<asp:Label ID="LabelErrorTestLdapConnection" Runat="server" EnableViewState="False" class="ms-error" />
					<asp:Label ID="LabelTestTenantConnectionOK" Runat="server" EnableViewState="False" />
				</p>
				</fieldset>
				</div>
			</td></tr>
		 </template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="User identifier property" description="Set the properties that identify users in Microsoft Entra ID.<br/>EntraCP automatically maps them to the identity claim type you set in the SPTrustedIdentityTokenIssuer.<br/><br/>Be cautious: Changing it may make existing Microsoft Entra ID user permissions invalid.">
        <template_inputformcontrols>
			<div id="divUserIdentifiers">
			<label>User identifier for 'Member' users:</label>
			<asp:DropDownList runat="server" ID="DDLDirectoryPropertyMemberUsers" class="ms-input">
			</asp:DropDownList>
			<br/>
			<label>User identifier for 'Guest' users:</label>
			<asp:DropDownList runat="server" ID="DDLDirectoryPropertyGuestUsers" class="ms-input">
			</asp:DropDownList>
			</div> 
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Display of user identifier results" description="Configure how entities created with the identity claim type appear in the people picker.<br/>It does not affect the actual value of the entity, which is always set with the user identifier property.">
        <template_inputformcontrols>
			<wssawc:InputFormRadioButton id="RbIdentityDefault"
				LabelText="Show the user identifier value"
				Checked="true"
				GroupName="RbIdentityDisplay"
				CausesValidation="false"
				runat="server" >
            </wssawc:InputFormRadioButton>
			<wssawc:InputFormRadioButton id="RbIdentityCustomGraphProperty"
				LabelText="Show the value of another property, e.g the display name:"
				GroupName="RbIdentityDisplay"
				CausesValidation="false"
				runat="server" >
            <wssuc:InputFormControl LabelText="InputFormControlLabelText">
				<Template_control>
                    <asp:DropDownList runat="server" ID="DDLGraphPropertyToDisplay" onclick="window.Entracp.EntracpSettingsPage.CheckRbIdentityCustomGraphProperty()" class="ms-input" />
				</Template_control>
			</wssuc:InputFormControl>
			</wssawc:InputFormRadioButton>
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Bypass Microsoft Entra ID lookup" description="Skip Microsoft Entra ID lookup and consider any input as valid.<br/><br/>This can be useful to keep people picker working even if connectivity with the Azure tenant is lost.">
        <template_inputformcontrols>
            <asp:Checkbox Runat="server" Name="ChkAlwaysResolveUserInput" ID="ChkAlwaysResolveUserInput" Text="Bypass Microsoft Entra ID lookup" />
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Require exact match" description="Enable this to return only results that match exactly the user input (case-insensitive).">
        <template_inputformcontrols>
			<asp:Checkbox Runat="server" Name="ChkFilterExactMatchOnly" ID="ChkFilterExactMatchOnly" Text="Require exact match" />
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Augmentation" >
		<template_description>
			<wssawc:EncodedLiteral runat="server" text="Enable augmentation to let EntraCP get " EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
			<a href="https://docs.microsoft.com/en-us/graph/api/user-getmembergroups" target="_blank"><wssawc:EncodedLiteral runat="server" text="all the Microsoft Entra ID groups" EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/></a>
			<wssawc:EncodedLiteral runat="server" text="that the user is a member of.<br/><br/>If not enabled, permissions granted to Microsoft Entra ID groups may not work correctly." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
		</template_description>
        <template_inputformcontrols>
			<asp:Checkbox Runat="server" Name="ChkAugmentAADRoles" ID="ChkAugmentAADRoles" Text="Retrieve Microsoft Entra ID groups" />
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Proxy" >
		<template_description>
			<wssawc:EncodedLiteral runat="server" text="Configure the proxy if it is needed for EntraCP to connect to Microsoft Graph." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
		</template_description>
        <template_inputformcontrols>
            <label for="<%= InputProxyAddress.ClientID %>">Proxy address:</label><br/>
            <wssawc:InputFormTextBox title="Proxy address" class="ms-input" ID="InputProxyAddress" Columns="50" Runat="server" />
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Type of groups">
        <template_description>
			<wssawc:EncodedLiteral runat="server" text="Set if all " EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
			<a href="https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0" target="_blank"><wssawc:EncodedLiteral runat="server" text="type of groups" EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/></a>
			<wssawc:EncodedLiteral runat="server" text="should be returned, including Office 365 unified groups, or only those that are security-enabled." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
		</template_description>
        <template_inputformcontrols>
			<asp:Checkbox Runat="server" Name="ChkFilterSecurityEnabledGroupsOnly" ID="ChkFilterSecurityEnabledGroupsOnly" Text="Return <a href='https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0' target='_blank'>security-enabled</a> groups only" />
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Reset EntraCP configuration" description="Restore configuration to its default values. All changes, including in claim types mappings, will be lost.">
        <template_inputformcontrols>
			<asp:Button runat="server" ID="BtnResetConfig" Text="Reset EntraCP configuration" onclick="BtnResetConfig_Click" class="ms-ButtonHeightWidth" OnClientClick="return confirm('Do you really want to reset EntraCP configuration?');" />
		</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:buttonsection runat="server">
        <template_buttons>
			<asp:Button UseSubmitBehavior="false" runat="server" class="ms-ButtonHeightWidth" OnClick="BtnOK_Click" Text="<%$Resources:wss,multipages_okbutton_text%>" id="BtnOK" accesskey="<%$Resources:wss,okbutton_accesskey%>"/>
		</template_buttons>
    </wssuc:buttonsection>
</table>
