<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="AzureCPGlobalSettings.ascx.cs" Inherits="azurecp.ControlTemplates.AzureCPGlobalSettings" %>
<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Register TagPrefix="wssuc" TagName="InputFormSection" Src="~/_controltemplates/InputFormSection.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="InputFormControl" Src="~/_controltemplates/InputFormControl.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="ButtonSection" Src="~/_controltemplates/ButtonSection.ascx" %>
<%@ Register TagPrefix="wssawc" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<script type="text/javascript" src="/_layouts/15/azurecp/jquery-1.9.1.min.js"></script>
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

    .Azurecp-success {
        color: green;
        font-weight: bold;
    }

    .Azurecp-HideCol {
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
    window.Azurecp = window.Azurecp || {};
    window.Azurecp.AzurecpSettingsPage = window.Azurecp.AzurecpSettingsPage || {};

    // Identity permission section
    window.Azurecp.AzurecpSettingsPage.CheckRbIdentityCustomGraphProperty = function () {
        var control = (document.getElementById("<%= RbIdentityCustomGraphProperty.ClientID %>"));
        if (control != null) {
            control.checked = true;
        }
    }

    // Register initialization method to run when DOM is ready and most SP JS functions loaded
    _spBodyOnLoadFunctionNames.push("window.Azurecp.AzurecpSettingsPage.Init");

    window.Azurecp.AzurecpSettingsPage.Init = function () {
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
    <wssuc:inputformsection title="Existing Azure Active Directory tenants" runat="server">
        <template_description>
				<wssawc:EncodedLiteral runat="server" text="Azure AD tenants registered." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
			</template_description>
        <template_inputformcontrols>
				<tr><td>
				<wssawc:SPGridView runat="server" ID="grdAzureTenants" AutoGenerateColumns="false" OnRowDeleting="grdAzureTenants_RowDeleting">
					<Columns>
						<asp:BoundField DataField="Id" ItemStyle-CssClass="Azurecp-HideCol" HeaderStyle-CssClass="Azurecp-HideCol"/>
						<asp:BoundField HeaderText="Tenant name" DataField="TenantName"/>
						<asp:BoundField HeaderText="Application ID" DataField="ClientID"/>
                        <asp:BoundField HeaderText="Filter out Guest users" DataField="MemberUserTypeOnly" />
						<asp:CommandField HeaderText="Action" ButtonType="Button" DeleteText="Remove" ShowDeleteButton="True" />
					</Columns>
				</wssawc:SPGridView>
				</td></tr>
			</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection title="New Azure Active Directory tenant" runat="server">
        <template_description>
				<wssawc:EncodedLiteral runat="server" text="Register a new Azure AD tenant." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
			</template_description>
        <template_inputformcontrols>
			<tr><td>
				<div id="divNewLdapConnection">
				<fieldset>
				<legend>Details about new Azure AD tenant</legend>
				<ul>
					<li>
						<label for="<%= TxtTenantName.ClientID %>">Tenant name: <em>*</em></label>
						<wssawc:InputFormTextBox title="Azure tenant name" class="ms-input" ID="TxtTenantName" Columns="50" Runat="server" MaxLength="255" Text="TENANTNAME.onMicrosoft.com" />
					</li>
					<li>
						<label for="<%= TxtClientId.ClientID %>">Application ID: <em>*</em></label>
						<wssawc:InputFormTextBox title="Password" class="ms-input" ID="TxtClientId" Columns="50" Runat="server" MaxLength="255" />
					</li>
					<li>
						<label for="<%= TxtClientSecret.ClientID %>">Application secret: <em>*</em></label>
						<wssawc:InputFormTextBox title="Password" class="ms-input" ID="TxtClientSecret" Columns="50" Runat="server" MaxLength="255" TextMode="Password" />
					</li>
					<li>
						<label for="<%=ChkMemberUserTypeOnly.ClientID %>">Filter out <a href="https://docs.microsoft.com/en-us/azure/active-directory/active-directory-b2b-user-properties" target="_blank">Guest users</a> on this tenant:</label>
						<table border="0" cellpadding="0" cellspacing="0" style="display: inline;">
						<wssawc:InputFormCheckBox class="ms-input" ID="ChkMemberUserTypeOnly" ToolTip="Filter out Guest users" runat="server" />
						</table>
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
	<wssuc:inputformsection runat="server" title="User identifier property" description="Set the properties that identify users in Azure Active Directory.<br/><br/>AzureCP automatically maps those properties with the identity claim type set in the SharePoint TrustedLoginProvider">
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
    <wssuc:inputformsection runat="server" title="Display of user identifier results" description="Configure how entities created with identity claim type are shown in the people picker.<br/>It does not change the actual value of the entity, that is the user identifier.">
        <template_inputformcontrols>
				<wssawc:InputFormRadioButton id="RbIdentityDefault"
					LabelText="Display the UserPrincipalName"
					Checked="true"
					GroupName="RbIdentityDisplay"
					CausesValidation="false"
					runat="server" >
                </wssawc:InputFormRadioButton>
				<wssawc:InputFormRadioButton id="RbIdentityCustomGraphProperty"
					LabelText="Display another property"
					GroupName="RbIdentityDisplay"
					CausesValidation="false"
					runat="server" >
                <wssuc:InputFormControl LabelText="InputFormControlLabelText">
					<Template_control>
						<wssawc:EncodedLiteral runat="server" text="You can choose to display a specific property (e.g the display name):<br/>" EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
                        <asp:DropDownList runat="server" ID="DDLGraphPropertyToDisplay" onclick="window.Azurecp.AzurecpSettingsPage.CheckRbIdentityCustomGraphProperty()" class="ms-input" />
					</Template_control>
				</wssuc:InputFormControl>
				</wssawc:InputFormRadioButton>
			</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Bypass Azure AD lookup" description="Completely bypass Azure AD lookup and consider any input as valid.<br/><br/>This can be useful to keep people picker working even if connectivity with Azure tenant is lost.">
        <template_inputformcontrols>
                <asp:Checkbox Runat="server" Name="ChkAlwaysResolveUserInput" ID="ChkAlwaysResolveUserInput" Text="Bypass Azure AD lookup" />
			</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Require exact match" description="Set to only return results that exactly match the user input (case-insensitive).">
        <template_inputformcontrols>
				<asp:Checkbox Runat="server" Name="ChkFilterExactMatchOnly" ID="ChkFilterExactMatchOnly" Text="Require exact match" />
			</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Augmentation" description="Enable augmentation to let AzureCP get group membership of Azure AD users.<br/><br/>If not enabled, permissions granted on Azure AD groups may not work.">
        <template_inputformcontrols>
				<asp:Checkbox Runat="server" Name="ChkAugmentAADRoles" ID="ChkAugmentAADRoles" Text="Retrieve Azure AD groups" />
			</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:inputformsection runat="server" title="Reset AzureCP configuration" description="Restore configuration to its default values. Every changes, including claim types configuration, will be reset.">
        <template_inputformcontrols>
				<asp:Button runat="server" ID="BtnResetAzureCPConfig" Text="Reset AzureCP configuration" onclick="BtnResetAzureCPConfig_Click" class="ms-ButtonHeightWidth" OnClientClick="return confirm('Do you really want to reset AzureCP configuration?');" />
			</template_inputformcontrols>
    </wssuc:inputformsection>
    <wssuc:buttonsection runat="server">
        <template_buttons>
			    <asp:Button UseSubmitBehavior="false" runat="server" class="ms-ButtonHeightWidth" OnClick="BtnOK_Click" Text="<%$Resources:wss,multipages_okbutton_text%>" id="BtnOK" accesskey="<%$Resources:wss,okbutton_accesskey%>"/>
		    </template_buttons>
    </wssuc:buttonsection>
</table>
