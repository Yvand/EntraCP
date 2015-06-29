<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>

<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AzureCPSettings.aspx.cs" Inherits="azurecp.AzureCPSettings" MasterPageFile="~/_admin/admin.master" %>

<%@ Register TagPrefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register TagPrefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Register TagPrefix="wssuc" TagName="InputFormSection" Src="~/_controltemplates/InputFormSection.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="InputFormControl" Src="~/_controltemplates/InputFormControl.ascx" %>
<%@ Register TagPrefix="wssuc" TagName="ButtonSection" Src="~/_controltemplates/ButtonSection.ascx" %>
<%@ Register TagPrefix="wssawc" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<asp:Content ID="PageHead" ContentPlaceHolderID="PlaceHolderAdditionalPageHead" runat="server">
    <script type="text/javascript" src="/_layouts/15/azurecp/jquery-1.9.1.min.js"></script>
    <style>
        /* Set the size of the right part with all input controls */
        .ms-inputformcontrols {
            width: 650px;
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
            width: 100px;
        }

        #divNewLdapConnection fieldset {
            border: 0;
            margin: 0;
            padding: 0;
        }

            #divNewLdapConnection fieldset ol {
                margin: 0;
                padding: 0;
            }

            #divNewLdapConnection fieldset li {
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
</asp:Content>
<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
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
        <wssuc:ButtonSection runat="server">
            <template_buttons>
			<asp:Button UseSubmitBehavior="false" runat="server" class="ms-ButtonHeightWidth" OnClick="BtnOK_Click" Text="<%$Resources:wss,multipages_okbutton_text%>" id="BtnOKTop" accesskey="<%$Resources:wss,okbutton_accesskey%>"/>
		</template_buttons>
        </wssuc:ButtonSection>
        <wssuc:InputFormSection Title="Current Azure tenants" runat="server">
            <template_description>
				<wssawc:EncodedLiteral runat="server" text="Current Azure tenants." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
			</template_description>
            <template_inputformcontrols>
				<tr><td>
				<wssawc:SPGridView runat="server" ID="grdAzureTenants" AutoGenerateColumns="false" OnRowDeleting="grdAzureTenants_RowDeleting">
					<Columns>
						<asp:BoundField DataField="Id" ItemStyle-CssClass="Azurecp-HideCol" HeaderStyle-CssClass="Azurecp-HideCol"/>
						<asp:BoundField HeaderText="TenantName" DataField="TenantName"/>
						<asp:BoundField HeaderText="ClientID" DataField="ClientID"/>
						<asp:CommandField HeaderText="Action" ButtonType="Button" DeleteText="Remove" ShowDeleteButton="True" />
					</Columns>
				</wssawc:SPGridView>
				</td></tr>
			</template_inputformcontrols>
        </wssuc:InputFormSection>
        <wssuc:InputFormSection Title="New Azure tenant" runat="server">
            <template_description>
				<wssawc:EncodedLiteral runat="server" text="Create a new connection to an Azure tenant." EncodeMethod='HtmlEncodeAllowSimpleTextFormatting'/>
			</template_description>
            <template_inputformcontrols>
				<tr><td>
				
				<div id="divNewLdapConnection">
				<fieldset>
				<legend>Informations to add new Azure tenant</legend>
				<ol>
					<li>
						<label for="<%= TxtTenantName.ClientID %>">Tenant <a href="http://msdn.microsoft.com/en-us/library/system.directoryservices.directoryentry.path(v=vs.110).aspx" target="_blank">name</a>: <em>*</em></label>
						<wssawc:InputFormTextBox title="Azure tenant name" class="ms-input" ID="TxtTenantName" Columns="50" Runat="server" MaxLength=255 Text="TENANTNAME.onMicrosoft.com" />
					</li>
					<li>
						<label for="<%= TxtTenantId.ClientID %>">Tenant ID: <em>*</em></label>
						<wssawc:InputFormTextBox title="Username" class="ms-input" ID="TxtTenantId" Columns="50" Runat="server" MaxLength=255 />
					</li>
					<li>
						<label for="<%= TxtClientId.ClientID %>">Client ID: <em>*</em></label>
						<wssawc:InputFormTextBox title="Password" class="ms-input" ID="TxtClientId" Columns="50" Runat="server" MaxLength=255 />
					</li>
					<li>
						<label for="<%= TxtClientSecret.ClientID %>">Client Secret: <em>*</em></label>
						<wssawc:InputFormTextBox title="Password" class="ms-input" ID="TxtClientSecret" Columns="50" Runat="server" MaxLength=255 TextMode="Password" />
					</li>
				</ol>
				</fieldset>
				</div>
					
				<div class="divbuttons">
					<asp:Button runat="server" ID="BtnTestAzureTenantConnection" Text="Test tenant connection" onclick="BtnTestAzureTenantConnection_Click" class="ms-ButtonHeightWidth" />
					<asp:Button runat="server" ID="BtnAddLdapConnection" Text="Add tenant" OnClick="BtnAddAzureTenant_Click" class="ms-ButtonHeightWidth" />
				</div>
				<p>
					<asp:Label ID="LabelErrorTestLdapConnection" Runat="server" EnableViewState="False" class="ms-error" />
					<asp:Label ID="LabelTestTenantConnectionOK" Runat="server" EnableViewState="False" />
				</p>
			</td></tr>
		 </template_inputformcontrols>
        </wssuc:InputFormSection>
        <wssuc:InputFormSection runat="server" Title="Display of permissions created with identity claim" Description="Customize the display text of permissions created with identity claim. Identity claim is defined in the TrustedLoginProvider.<br/> It does not impact the actual value of the permission that will always be the property associated with the identity claim.">
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
        </wssuc:InputFormSection>
        <wssuc:InputFormSection runat="server" Title="Bypass Azure AD lookup" Description="Bypass Azure AD lookup and validate user input on each claim type configured (quite similar to standard SharePoint behavior).<br/>It may be useful to enable if there is a temporary issue to reach the tenant, as it will still be possible to grant new permissions and validate existing ones.">
            <template_inputformcontrols>
                <asp:Checkbox Runat="server" Name="ChkAlwaysResolveUserInput" ID="ChkAlwaysResolveUserInput" Text="Bypass Azure AD lookup" />
			</template_inputformcontrols>
        </wssuc:InputFormSection>
        <wssuc:InputFormSection runat="server" Title="Require exact match" Description="Set to only return results that exactly match the user input (case-insensitive).">
            <template_inputformcontrols>
				<asp:Checkbox Runat="server" Name="ChkFilterExactMatchOnly" ID="ChkFilterExactMatchOnly" Text="Require exact match" />
			</template_inputformcontrols>
        </wssuc:InputFormSection>
        <wssuc:InputFormSection runat="server" Title="Retrieve Azure AD groups" Description="If enabled, every time a user authenticates, his Azure tenant will be queried to retrieve his groups and add them in his SAML token, so that SharePoint can process permissions on Azure groups.">
            <template_inputformcontrols>
				<asp:Checkbox Runat="server" Name="ChkAugmentAADRoles" ID="ChkAugmentAADRoles" Text="Retrieve Azure AD groups" />
			</template_inputformcontrols>
        </wssuc:InputFormSection>
        <wssuc:InputFormSection runat="server" Title="Reset AzureCP configuration" Description="This will delete the AzureCP persisted object in configuration database and recreate one with default values. Every custom settings, including customized claim types, will be deleted.">
            <template_inputformcontrols>
				<asp:Button runat="server" ID="BtnResetAzureCPConfig" Text="Reset AzureCP configuration" onclick="BtnResetAzureCPConfig_Click" class="ms-ButtonHeightWidth" OnClientClick="return confirm('Do you really want to reset AzureCP configuration?');" />
			</template_inputformcontrols>
        </wssuc:InputFormSection>
        <wssuc:ButtonSection runat="server">
            <template_buttons>
			    <asp:Button UseSubmitBehavior="false" runat="server" class="ms-ButtonHeightWidth" OnClick="BtnOK_Click" Text="<%$Resources:wss,multipages_okbutton_text%>" id="BtnOK" accesskey="<%$Resources:wss,okbutton_accesskey%>"/>
		    </template_buttons>
        </wssuc:ButtonSection>
    </table>
</asp:Content>
<asp:Content ID="PageTitle" ContentPlaceHolderID="PlaceHolderPageTitle" runat="server">
    AzureCP Configuration
</asp:Content>
<asp:Content ID="PageTitleInTitleArea" ContentPlaceHolderID="PlaceHolderPageTitleInTitleArea"
    runat="server">
    <asp:Label ID="LblTitle" runat="server" />
</asp:Content>
