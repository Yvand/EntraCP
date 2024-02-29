using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.JSGrid;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using Yvand.EntraClaimsProvider.Configuration;
using Yvand.EntraClaimsProvider.Logging;
using Logger = Yvand.EntraClaimsProvider.Logging.Logger;

namespace Yvand.EntraClaimsProvider.Administration
{
    public partial class GlobalSettingsUserControl : EntraCPUserControl
    {
        readonly string TextErrorNewTenantFieldsMissing = "Some mandatory fields are missing.";
        readonly string TextErrorTestAzureADConnection = "Unable to get access token for tenant '{0}': {1}";
        readonly string TextConnectionSuccessful = "Connection successful.";
        readonly string TextErrorNewTenantCreds = "Specify either a client secret or a client certificate, but not both.";
        readonly string TextErrorExtensionAttributesApplicationId = "Please specify a valid Client ID for AD Connect.";
        readonly string TextSummaryPersistedObjectInformation = "Found configuration '{0}' v{1} (Persisted Object ID: '{2}')";

        protected void Page_Load(object sender, EventArgs e)
        {
            Initialize();
        }

        /// <summary>
        /// Initialize controls as needed if prerequisites are ok, otherwise deactivate controls and show error message
        /// </summary>
        protected void Initialize()
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood)
            {
                this.LabelErrorMessage.Text = base.MostImportantError;
                this.BtnOK.Enabled = false;
                this.BtnOKTop.Enabled = false;
                return;
            }

            LabelMessage.Text = String.Format(TextSummaryPersistedObjectInformation, Configuration.Name, Configuration.Version, Configuration.Id);
            PopulateConnectionsGrid();
            if (!this.IsPostBack)
            {
                BuildGraphPropertyDDL();
                PopulateFields();
            }
        }

        void PopulateConnectionsGrid()
        {
            if (Settings.EntraIDTenants != null)
            {
                PropertyCollectionBinder pcb = new PropertyCollectionBinder();
                foreach (EntraIDTenant tenant in Settings.EntraIDTenants)
                {
                    if (tenant == null)
                    {
                        continue;
                    }
                    pcb.AddRow(tenant.Identifier, tenant.Name, tenant.ClientId, tenant.AuthenticationMode, tenant.ExtensionAttributesApplicationId);
                }
                pcb.BindGrid(grdAzureTenants);
            }
        }

        private void PopulateFields()
        {
            this.lblUserIdClaimType.Text = Settings.ClaimTypes.UserIdentifierConfig.ClaimType;
            if (IdentityCTConfig.EntityPropertyToUseAsDisplayText == DirectoryObjectProperty.NotSet)
            {
                //this.RbIdentityDefault.Checked = true;
                this.DDLGraphPropertyToDisplay.Items.FindByValue("NotSet").Selected = true;
            }
            else
            {
                //this.RbIdentityCustomGraphProperty.Checked = true;
                this.DDLGraphPropertyToDisplay.Items.FindByValue(((int)IdentityCTConfig.EntityPropertyToUseAsDisplayText).ToString()).Selected = true;
            }
            this.DDLDirectoryPropertyMemberUsers.Items.FindByValue(((int)IdentityCTConfig.EntityProperty).ToString()).Selected = true;
            this.DDLDirectoryPropertyGuestUsers.Items.FindByValue(((int)IdentityCTConfig.DirectoryObjectPropertyForGuestUsers).ToString()).Selected = true;
            this.ChkAlwaysResolveUserInput.Checked = Settings.AlwaysResolveUserInput;
            this.ChkFilterExactMatchOnly.Checked = Settings.FilterExactMatchOnly;
            this.ChkAugmentAADRoles.Checked = Settings.EnableAugmentation;
            this.ChkFilterSecurityEnabledGroupsOnly.Checked = Settings.FilterSecurityEnabledGroupsOnly;
            this.InputProxyAddress.Text = Settings.ProxyAddress;

            AzureCloudInstance[] azureCloudInstanceValues = (AzureCloudInstance[])Enum.GetValues(typeof(AzureCloudInstance));
            foreach (var azureCloudInstanceValue in azureCloudInstanceValues)
            {
                if (azureCloudInstanceValue == AzureCloudInstance.None) { continue; }
                this.DDLAzureCloudInstance.Items.Add(new System.Web.UI.WebControls.ListItem(azureCloudInstanceValue.ToString(), azureCloudInstanceValue.ToString()));
            }
            this.DDLAzureCloudInstance.SelectedValue = AzureCloudInstance.AzurePublic.ToString();
        }

        private void BuildGraphPropertyDDL()
        {
            this.DDLGraphPropertyToDisplay.Items.Add(new System.Web.UI.WebControls.ListItem("(Same as the identifier property)", "NotSet"));
            DirectoryObjectProperty[] aadPropValues = (DirectoryObjectProperty[])Enum.GetValues(typeof(DirectoryObjectProperty));
            IEnumerable<DirectoryObjectProperty> aadPropValuesSorted = aadPropValues.OrderBy(v => v.ToString());
            foreach (DirectoryObjectProperty prop in aadPropValuesSorted)
            {
                // Ensure property exists for the User object type
                if (Utils.GetDirectoryObjectPropertyValue(new User(), prop.ToString()) == null) { continue; }

                // Ensure property is of type System.String
                PropertyInfo pi = typeof(User).GetProperty(prop.ToString());
                if (pi == null) { continue; }
                if (pi.PropertyType != typeof(System.String)) { continue; }

                this.DDLGraphPropertyToDisplay.Items.Add(new System.Web.UI.WebControls.ListItem(prop.ToString(), ((int)prop).ToString()));
                this.DDLDirectoryPropertyMemberUsers.Items.Add(new System.Web.UI.WebControls.ListItem(prop.ToString(), ((int)prop).ToString()));
                this.DDLDirectoryPropertyGuestUsers.Items.Add(new System.Web.UI.WebControls.ListItem(prop.ToString(), ((int)prop).ToString()));
            }
        }

        protected void grdAzureTenants_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) { return; }
            if (Settings.EntraIDTenants == null) { return; }

            GridViewRow rowToDelete = grdAzureTenants.Rows[e.RowIndex];
            Guid Id = new Guid(rowToDelete.Cells[0].Text);
            EntraIDTenant tenantToRemove = Settings.EntraIDTenants.FirstOrDefault(x => x.Identifier == Id);
            if (tenantToRemove != null)
            {
                Settings.EntraIDTenants.Remove(tenantToRemove);
                CommitChanges();
                Logger.Log($"Microsoft Entra ID tenant '{tenantToRemove.Name}' was successfully removed from configuration '{ConfigurationName}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Configuration);
                LabelMessage.Text = String.Format(TextSummaryPersistedObjectInformation, Configuration.Name, Configuration.Version, Configuration.Id);
                PopulateConnectionsGrid();
            }
        }

        protected bool UpdateConfiguration(bool commitChanges)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) { return false; }

            if (this.DDLGraphPropertyToDisplay.SelectedValue != "NotSet")
            {
                Settings.ClaimTypes.UserIdentifierConfig.EntityPropertyToUseAsDisplayText = (DirectoryObjectProperty)Convert.ToInt32(this.DDLGraphPropertyToDisplay.SelectedValue);
            }
            else
            {
                Settings.ClaimTypes.UserIdentifierConfig.EntityPropertyToUseAsDisplayText = DirectoryObjectProperty.NotSet;
            }

            DirectoryObjectProperty newUserIdentifier = (DirectoryObjectProperty)Convert.ToInt32(this.DDLDirectoryPropertyMemberUsers.SelectedValue);
            if (newUserIdentifier != DirectoryObjectProperty.NotSet)
            {
                Settings.ClaimTypes.UpdateUserIdentifier(newUserIdentifier);
            }

            DirectoryObjectProperty newIdentifierForGuestUsers = (DirectoryObjectProperty)Convert.ToInt32(this.DDLDirectoryPropertyGuestUsers.SelectedValue);
            if (newIdentifierForGuestUsers != DirectoryObjectProperty.NotSet)
            {
                Settings.ClaimTypes.UpdateIdentifierForGuestUsers(newIdentifierForGuestUsers);
            }

            Settings.AlwaysResolveUserInput = this.ChkAlwaysResolveUserInput.Checked;
            Settings.FilterExactMatchOnly = this.ChkFilterExactMatchOnly.Checked;
            Settings.EnableAugmentation = this.ChkAugmentAADRoles.Checked;
            Settings.FilterSecurityEnabledGroupsOnly = this.ChkFilterSecurityEnabledGroupsOnly.Checked;
            Settings.ProxyAddress = this.InputProxyAddress.Text;

            if (commitChanges) { CommitChanges(); }
            return true;
        }

        protected void BtnTestAzureTenantConnection_Click(Object sender, EventArgs e)
        {
            this.ValidateAzureTenantConnection();
        }

        protected void ValidateAzureTenantConnection()
        {
            if (String.IsNullOrWhiteSpace(this.TxtTenantName.Text) || String.IsNullOrWhiteSpace(this.TxtClientId.Text))
            {
                this.LabelErrorTestLdapConnection.Text = TextErrorNewTenantFieldsMissing;
                return;
            }

            if ((InputClientCertFile.PostedFile == null && String.IsNullOrWhiteSpace(this.TxtClientSecret.Text)) ||
                (InputClientCertFile.PostedFile != null && InputClientCertFile.PostedFile.ContentLength == 0 && String.IsNullOrWhiteSpace(TxtClientSecret.Text)) ||
                (InputClientCertFile.PostedFile != null && InputClientCertFile.PostedFile.ContentLength != 0 && !String.IsNullOrWhiteSpace(TxtClientSecret.Text)))
            {
                this.LabelErrorTestLdapConnection.Text = TextErrorNewTenantCreds;
                return;
            }
            AzureCloudInstance cloudInstance = (AzureCloudInstance)Enum.Parse(typeof(AzureCloudInstance), this.DDLAzureCloudInstance.SelectedValue);

            EntraIDTenant newTenant = new EntraIDTenant
            {
                Name = this.TxtTenantName.Text,
                AzureAuthority = ClaimsProviderConstants.AzureCloudEndpoints.FirstOrDefault(item => item.Key == cloudInstance).Value,
            };

            if (String.IsNullOrWhiteSpace(this.TxtClientSecret.Text))
            {
                X509Certificate2 cert = null;
                if (ValidateUploadedCertFile(InputClientCertFile, this.InputClientCertPassword.Text, out cert) == false)
                {
                    return;
                }
                newTenant.SetCredentials(this.TxtClientId.Text, cert);
            }
            else
            {
                newTenant.SetCredentials(this.TxtClientId.Text, this.TxtClientSecret.Text);
            }

            try
            {
                //Task<bool> taskTestConnection = newTenant.TestConnectionAsync(Settings.ProxyAddress);
                Task<bool> taskTestConnection = Task.Run(async () => await newTenant.TestConnectionAsync(Settings.ProxyAddress));
                taskTestConnection.Wait();
                bool success = taskTestConnection.Result;
                if (success)
                {
                    this.LabelTestTenantConnectionOK.Text = TextConnectionSuccessful;
                }
                else
                {
                    this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, newTenant.Name, String.Empty);
                }
            }
            catch (AggregateException ex)
            {
                this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, newTenant.Name, ex.InnerException.Message);
            }
            catch (Exception ex)
            {
                this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, newTenant.Name, ex.Message);
            }
            //});
        }

        protected void BtnOK_Click(Object sender, EventArgs e)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) { return; }
            if (UpdateConfiguration(true))
            {
                Response.Redirect("/Security.aspx", false);
            }
            else
            {
                LabelErrorMessage.Text = base.MostImportantError;
            }
        }

        protected void BtnResetConfig_Click(Object sender, EventArgs e)
        {
            EntraIDProviderConfiguration.DeleteGlobalConfiguration(ConfigurationID);
            Response.Redirect(Request.RawUrl, false);
        }

        protected void BtnAddAzureTenant_Click(object sender, EventArgs e)
        {
            AddTenantConnection();
        }

        /// <summary>
        /// Add new Microsoft Entra ID tenant in persisted object
        /// </summary>
        void AddTenantConnection()
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) { return; }
            if (String.IsNullOrWhiteSpace(this.TxtTenantName.Text) || String.IsNullOrWhiteSpace(this.TxtClientId.Text))
            {
                this.LabelErrorTestLdapConnection.Text = TextErrorNewTenantFieldsMissing;
                return;
            }

            if (!string.IsNullOrWhiteSpace(this.TxtExtensionAttributesApplicationId.Text))
            {
                try
                {
                    Guid extensionAttributesApplicationId = Guid.Parse(this.TxtExtensionAttributesApplicationId.Text);
                }
                catch (Exception)
                {

                    this.LabelErrorTestLdapConnection.Text = TextErrorExtensionAttributesApplicationId;
                }

            }

            if ((InputClientCertFile.PostedFile == null && String.IsNullOrWhiteSpace(this.TxtClientSecret.Text)) ||
                (InputClientCertFile.PostedFile != null && InputClientCertFile.PostedFile.ContentLength == 0 && String.IsNullOrWhiteSpace(TxtClientSecret.Text)) ||
                (InputClientCertFile.PostedFile != null && InputClientCertFile.PostedFile.ContentLength != 0 && !String.IsNullOrWhiteSpace(TxtClientSecret.Text)))
            {
                this.LabelErrorTestLdapConnection.Text = TextErrorNewTenantCreds;
                return;
            }

            X509Certificate2 cert = null;
            if (String.IsNullOrWhiteSpace(this.TxtClientSecret.Text))
            {
                if (ValidateUploadedCertFile(InputClientCertFile, this.InputClientCertPassword.Text, out cert) == false)
                {
                    return;
                }
            }

            Uri cloudInstance = ClaimsProviderConstants.AzureCloudEndpoints.FirstOrDefault(item => item.Key == (AzureCloudInstance)Enum.Parse(typeof(AzureCloudInstance), this.DDLAzureCloudInstance.SelectedValue)).Value;


            if (Settings.EntraIDTenants == null)
            {
                Settings.EntraIDTenants = new List<EntraIDTenant>();
            }

            var newTenant = new EntraIDTenant
            {
                Name = this.TxtTenantName.Text,
                ExcludeGuestUsers = this.ChkMemberUserTypeOnly.Checked,
                AzureAuthority = cloudInstance,
                ExtensionAttributesApplicationId = string.IsNullOrWhiteSpace(this.TxtExtensionAttributesApplicationId.Text) ? Guid.Empty : Guid.Parse(this.TxtExtensionAttributesApplicationId.Text)
            };

            if (String.IsNullOrWhiteSpace(this.TxtClientSecret.Text))
            {
                newTenant.SetCredentials(this.TxtClientId.Text, cert);
            }
            else
            {
                newTenant.SetCredentials(this.TxtClientId.Text, this.TxtClientSecret.Text);
            }
            this.Settings.EntraIDTenants.Add(newTenant);

            CommitChanges();
            Logger.Log($"Microsoft Entra ID tenant '{this.TxtTenantName.Text}' was successfully added to configuration '{ConfigurationName}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Configuration);
            LabelMessage.Text = String.Format(TextSummaryPersistedObjectInformation, Configuration.Name, Configuration.Version, Configuration.Id);

            PopulateConnectionsGrid();
            this.TxtTenantName.Text = "TENANTNAME.onMicrosoft.com";
            this.TxtClientId.Text = String.Empty;
            this.TxtClientSecret.Text = String.Empty;
            this.InputClientCertPassword.Text = String.Empty;
            this.TxtExtensionAttributesApplicationId.Text = String.Empty;
            this.DDLAzureCloudInstance.SelectedValue = AzureCloudInstance.AzurePublic.ToString();
        }

        private bool ValidateUploadedCertFile(
            HtmlInputFile inputFile,
            string certificatePassword,
            out X509Certificate2 cert)
        {
            cert = null;
            if (inputFile.PostedFile == null ||
                inputFile.PostedFile.ContentLength == 0)
            {
                this.LabelErrorTestLdapConnection.Text = $"No certificate was passed.";
                return false;
            }

            // Ensure that fileName is just the file name (no directories), then check that fileName is legal.
            string fileName = string.Empty;
            try
            {
                fileName = Path.GetFileName(inputFile.PostedFile.FileName);
            }
            catch (ArgumentException ex)
            {
                this.LabelErrorTestLdapConnection.Text = $"Invalid file path. Error message: {ex.Message}";
                return false;
            }
            if (!SPUrlUtility.IsLegalFileName(fileName))
            {
                this.LabelErrorTestLdapConnection.Text = $"The file name is not legal.";
                return false;
            }

            try
            {
                byte[] buffer = new byte[inputFile.PostedFile.ContentLength];
                inputFile.PostedFile.InputStream.Read(buffer, 0, buffer.Length);
                // The certificate must be exportable so it can be saved in the persisted object
                cert = EntraIDTenant.ImportPfxCertificate(buffer, certificatePassword);
                if (cert == null)
                {
                    this.LabelErrorTestLdapConnection.Text = $"Certificate does not contain the private key.";
                    return false;
                }

                // Try to export the certificate with its private key to validate that it succeeds
                cert.Export(X509ContentType.Pfx, "Yvan");
            }
            catch (CryptographicException ex)
            {
                this.LabelErrorTestLdapConnection.Text = $"Invalid certificate. Error message: {ex.Message}";
                return false;
            }
            return true;
        }
    }

    public class PropertyCollectionBinder
    {
        protected DataTable PropertyCollection = new DataTable();
        public PropertyCollectionBinder()
        {
            PropertyCollection.Columns.Add("Id", typeof(Guid));
            PropertyCollection.Columns.Add("TenantName", typeof(string));
            PropertyCollection.Columns.Add("ClientID", typeof(string));
            //PropertyCollection.Columns.Add("MemberUserTypeOnly", typeof(bool));
            PropertyCollection.Columns.Add("AuthenticationMode", typeof(string));
            PropertyCollection.Columns.Add("ExtensionAttributesApplicationId", typeof(Guid));
        }

        public void AddRow(Guid Id, string TenantName, string ClientID, string AuthenticationMode, Guid ExtensionAttributesApplicationId)
        {
            DataRow newRow = PropertyCollection.Rows.Add();
            newRow["Id"] = Id;
            newRow["TenantName"] = TenantName;
            newRow["ClientID"] = ClientID;
            //newRow["MemberUserTypeOnly"] = MemberUserTypeOnly;
            newRow["AuthenticationMode"] = AuthenticationMode;
            newRow["ExtensionAttributesApplicationId"] = ExtensionAttributesApplicationId;
        }

        public void BindGrid(SPGridView grid)
        {
            grid.DataSource = PropertyCollection.DefaultView;
            grid.DataBind();
        }
    }
}
