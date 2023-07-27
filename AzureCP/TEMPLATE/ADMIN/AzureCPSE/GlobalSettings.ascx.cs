using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
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
using Yvand.ClaimsProviders;
using Yvand.ClaimsProviders.Configuration;
using Yvand.ClaimsProviders.Configuration.AzureAD;
using static Yvand.ClaimsProviders.ClaimsProviderLogging;

namespace Yvand.ClaimsProviders.Administration
{
    public partial class GlobalSettings : AzureCPUserControl
    {
        readonly string TextErrorNewTenantFieldsMissing = "Some mandatory fields are missing.";
        readonly string TextErrorTestAzureADConnection = "Unable to get access token for tenant '{0}': {1}";
        readonly string TextConnectionSuccessful = "Connection successful.";
        readonly string TextErrorNewTenantCreds = "Specify either a client secret or a client certificate, but not both.";
        readonly string TextErrorExtensionAttributesApplicationId = "Please specify a valid Client ID for AD Connect.";

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

            PopulateConnectionsGrid();
            if (!this.IsPostBack)
            {
                BuildGraphPropertyDDL();
                PopulateFields();
            }
        }

        void PopulateConnectionsGrid()
        {
            if (Configuration.AzureTenants != null)
            {
                PropertyCollectionBinder pcb = new PropertyCollectionBinder();
                foreach (AzureTenant tenant in Configuration.AzureTenants)
                {
                    pcb.AddRow(tenant.Identifier, tenant.Name, tenant.ClientId, tenant.CloudInstance.ToString(), tenant.ExtensionAttributesApplicationId);
                }
                pcb.BindGrid(grdAzureTenants);
            }
        }

        private void PopulateFields()
        {
            if (Configuration.IdentityClaimTypeConfig.EntityPropertyToUseAsDisplayText == DirectoryObjectProperty.NotSet)
            {
                this.RbIdentityDefault.Checked = true;
            }
            else
            {
                this.RbIdentityCustomGraphProperty.Checked = true;
                this.DDLGraphPropertyToDisplay.Items.FindByValue(((int)Configuration.IdentityClaimTypeConfig.EntityPropertyToUseAsDisplayText).ToString()).Selected = true;
            }
            this.DDLDirectoryPropertyMemberUsers.Items.FindByValue(((int)Configuration.IdentityClaimTypeConfig.EntityProperty).ToString()).Selected = true;
            this.DDLDirectoryPropertyGuestUsers.Items.FindByValue(((int)Configuration.IdentityClaimTypeConfig.DirectoryObjectPropertyForGuestUsers).ToString()).Selected = true;
            this.ChkAlwaysResolveUserInput.Checked = Configuration.AlwaysResolveUserInput;
            this.ChkFilterExactMatchOnly.Checked = Configuration.FilterExactMatchOnly;
            this.ChkAugmentAADRoles.Checked = Configuration.EnableAugmentation;
            this.ChkFilterSecurityEnabledGroupsOnly.Checked = Configuration.FilterSecurityEnabledGroupsOnly;
            this.InputProxyAddress.Text = Configuration.ProxyAddress;

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
            DirectoryObjectProperty[] aadPropValues = (DirectoryObjectProperty[])Enum.GetValues(typeof(DirectoryObjectProperty));
            IEnumerable<DirectoryObjectProperty> aadPropValuesSorted = aadPropValues.OrderBy(v => v.ToString());
            foreach (DirectoryObjectProperty prop in aadPropValuesSorted)
            {
                // Ensure property exists for the User object type
                if (AzureCPSE.GetPropertyValue(new User(), prop.ToString()) == null) { continue; }

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
            if (Configuration.AzureTenants == null) { return; }

            GridViewRow rowToDelete = grdAzureTenants.Rows[e.RowIndex];
            Guid Id = new Guid(rowToDelete.Cells[0].Text);
            AzureTenant tenantToRemove = Configuration.AzureTenants.FirstOrDefault(x => x.Identifier == Id);
            if (tenantToRemove != null)
            {
                Configuration.AzureTenants.Remove(tenantToRemove);
                CommitChanges();
                ClaimsProviderLogging.Log($"Azure AD tenant '{tenantToRemove.Name}' was successfully removed from configuration '{ConfigurationName}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Configuration);
                PopulateConnectionsGrid();
            }
        }

        protected bool UpdateConfiguration(bool commitChanges)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) { return false; }

            if (this.RbIdentityCustomGraphProperty.Checked)
            {
                Configuration.IdentityClaimTypeConfig.EntityPropertyToUseAsDisplayText = (DirectoryObjectProperty)Convert.ToInt32(this.DDLGraphPropertyToDisplay.SelectedValue);
            }
            else
            {
                Configuration.IdentityClaimTypeConfig.EntityPropertyToUseAsDisplayText = DirectoryObjectProperty.NotSet;
            }

            DirectoryObjectProperty newUserIdentifier = (DirectoryObjectProperty)Convert.ToInt32(this.DDLDirectoryPropertyMemberUsers.SelectedValue);
            if (newUserIdentifier != DirectoryObjectProperty.NotSet)
            {
                Configuration.ClaimTypes.UpdateUserIdentifier(newUserIdentifier);
            }

            DirectoryObjectProperty newIdentifierForGuestUsers = (DirectoryObjectProperty)Convert.ToInt32(this.DDLDirectoryPropertyGuestUsers.SelectedValue);
            if (newIdentifierForGuestUsers != DirectoryObjectProperty.NotSet)
            {
                Configuration.ClaimTypes.UpdateIdentifierForGuestUsers(newIdentifierForGuestUsers);
            }

            Configuration.AlwaysResolveUserInput = this.ChkAlwaysResolveUserInput.Checked;
            Configuration.FilterExactMatchOnly = this.ChkFilterExactMatchOnly.Checked;
            Configuration.EnableAugmentation = this.ChkAugmentAADRoles.Checked;
            Configuration.FilterSecurityEnabledGroupsOnly = this.ChkFilterSecurityEnabledGroupsOnly.Checked;
            Configuration.ProxyAddress = this.InputProxyAddress.Text;

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

            string tenantName = this.TxtTenantName.Text;
            string clientId = this.TxtClientId.Text;
            string clientSecret = this.TxtClientSecret.Text;
            AzureCloudInstance cloudInstance = (AzureCloudInstance)Enum.Parse(typeof(AzureCloudInstance), this.DDLAzureCloudInstance.SelectedValue);
            // The whole flow of setting the certificate and testing it in AADAppOnlyAuthenticationProvider needs to be done as app pool account
            // Otherwise AADAppOnlyAuthenticationProvider throws CryptographicException: Keyset does not exist (which means it could not access the private key) 
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                X509Certificate2 cert = null;
                if (String.IsNullOrWhiteSpace(this.TxtClientSecret.Text))
                {
                    if (ValidateUploadedCertFile(InputClientCertFile, this.InputClientCertPassword.Text, out cert) == false)
                    { return; }
                }
                try
                {
                    //AADAppOnlyAuthenticationProvider testConnection;
                    //if (String.IsNullOrWhiteSpace(this.TxtClientSecret.Text))
                    //{
                    //    testConnection = new AADAppOnlyAuthenticationProvider(cloudInstance, tenantName, clientId, cert, String.Empty, ClaimsProviderConstants.DEFAULT_TIMEOUT);
                    //}
                    //else
                    //{
                    //    testConnection = new AADAppOnlyAuthenticationProvider(cloudInstance, tenantName, clientId, clientSecret, String.Empty, ClaimsProviderConstants.DEFAULT_TIMEOUT);
                    //}
                    //Task<bool> testConnectionTask = testConnection.GetAccessToken(true);
                    //testConnectionTask.Wait();
                    this.LabelTestTenantConnectionOK.Text = TextConnectionSuccessful;
                }
                catch (MsalServiceException ex)
                {
                    this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, tenantName, ex.Message);
                }
                catch (Exception ex)
                {
                    this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, tenantName, ex.Message);
                }
            });
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

        protected void BtnResetAzureCPConfig_Click(Object sender, EventArgs e)
        {
            //AzureADEntityProviderConfiguration.DeleteConfiguration(PersistedObjectName);
            EntityProviderBase<AzureADEntityProviderConfiguration>.DeleteGlobalConfiguration(ConfigurationName);
            Response.Redirect(Request.RawUrl, false);
        }

        protected void BtnAddAzureTenant_Click(object sender, EventArgs e)
        {
            AddTenantConnection();
        }

        /// <summary>
        /// Add new Azure AD tenant in persisted object
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


            if (Configuration.AzureTenants == null)
            {
                Configuration.AzureTenants = new List<AzureTenant>();
            }
            this.Configuration.AzureTenants.Add(
                new AzureTenant
                {
                    Name = this.TxtTenantName.Text,
                    ClientId = this.TxtClientId.Text,
                    ClientSecret = this.TxtClientSecret.Text,
                    ExcludeGuestUsers = this.ChkMemberUserTypeOnly.Checked,
                    ClientCertificateWithPrivateKey = cert,
                    AzureAuthority = cloudInstance,
                    ExtensionAttributesApplicationId = string.IsNullOrWhiteSpace(this.TxtExtensionAttributesApplicationId.Text) ? Guid.Empty : Guid.Parse(this.TxtExtensionAttributesApplicationId.Text)
                });

            CommitChanges();
            ClaimsProviderLogging.Log($"Azure AD tenant '{this.TxtTenantName.Text}' was successfully added to configuration '{ConfigurationName}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Configuration);

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
                cert = AzureTenant.ImportPfxCertificateBlob(buffer, certificatePassword, X509KeyStorageFlags.Exportable);
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
            PropertyCollection.Columns.Add("CloudInstance", typeof(string));
            PropertyCollection.Columns.Add("ExtensionAttributesApplicationId", typeof(Guid));
        }

        public void AddRow(Guid Id, string TenantName, string ClientID, string CloudInstance, Guid ExtensionAttributesApplicationId)
        {
            DataRow newRow = PropertyCollection.Rows.Add();
            newRow["Id"] = Id;
            newRow["TenantName"] = TenantName;
            newRow["ClientID"] = ClientID;
            //newRow["MemberUserTypeOnly"] = MemberUserTypeOnly;
            newRow["CloudInstance"] = CloudInstance;
            newRow["ExtensionAttributesApplicationId"] = ExtensionAttributesApplicationId;
        }

        public void BindGrid(SPGridView grid)
        {
            grid.DataSource = PropertyCollection.DefaultView;
            grid.DataBind();
        }
    }
}
