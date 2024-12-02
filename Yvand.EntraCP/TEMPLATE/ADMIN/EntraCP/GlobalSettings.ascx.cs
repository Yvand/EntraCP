using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.JSGrid;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
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
using ListItem = System.Web.UI.WebControls.ListItem;
using Logger = Yvand.EntraClaimsProvider.Logging.Logger;

namespace Yvand.EntraClaimsProvider.Administration
{
    public partial class GlobalSettingsUserControl : EntraCPUserControl
    {
        protected int DefaultTenantDataCacheLifetimeInMinutes = ClaimsProviderConstants.DefaultTenantDataCacheLifetimeInMinutes;

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
            if (!this.IsPostBack)
            {
                PopulateConnectionsGrid();
                PopulateGraphPropertiesLists();
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
                    pcb.AddRow(tenant.Identifier, tenant.Name, tenant.ClientId, tenant.AuthenticationMode, tenant.AzureCloud.ToString());
                }
                pcb.BindGrid(grdAzureTenants);
            }
        }

        private void PopulateFields()
        {
            // User identifier settings
            this.lblUserIdClaimType.Text = Settings.ClaimTypes.UserIdentifierConfig.ClaimType;
            if (IdentityCTConfig.EntityPropertyToUseAsDisplayText == DirectoryObjectProperty.NotSet)
            {
                this.DdlUserGraphPropertyToDisplay.Items.FindByValue("NotSet").Selected = true;
            }
            else
            {
                this.DdlUserGraphPropertyToDisplay.Items.FindByValue(((int)IdentityCTConfig.EntityPropertyToUseAsDisplayText).ToString()).Selected = true;
            }
            this.DdlUserIdDirectoryPropertyMembers.Items.FindByValue(((int)IdentityCTConfig.EntityProperty).ToString()).Selected = true;
            this.DdlUserIdDirectoryPropertyGuests.Items.FindByValue(((int)IdentityCTConfig.DirectoryObjectPropertyForGuestUsers).ToString()).Selected = true;

            // Group identifier settings
            var possibleGroupClaimTypes = SPTrust.ClaimTypeInformation
                .Where(x => !String.Equals(x.MappedClaimType, Settings.ClaimTypes.UserIdentifierConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                .Select(x => x.MappedClaimType);
            foreach (string possibleGroupClaimType in possibleGroupClaimTypes)
            {
                ListItem possibleGroupClaimTypeItem = new ListItem(possibleGroupClaimType);
                this.DdlGroupClaimType.Items.Add(possibleGroupClaimTypeItem);
            }

            ClaimTypeConfig groupCtc = Settings.ClaimTypes.GroupIdentifierConfig;
            if (groupCtc != null)
            {
                this.DdlGroupClaimType.SelectedValue = groupCtc.ClaimType;
                this.DdlGroupDirectoryProperty.Items.FindByValue(((int)groupCtc.EntityProperty).ToString()).Selected = true;
                if (groupCtc.EntityPropertyToUseAsDisplayText == DirectoryObjectProperty.NotSet)
                {
                    this.DdlGroupGraphPropertyToDisplay.Items.FindByValue("NotSet").Selected = true;
                }
                else
                {
                    this.DdlGroupGraphPropertyToDisplay.Items.FindByValue(((int)groupCtc.EntityPropertyToUseAsDisplayText).ToString()).Selected = true;
                }
            }
            this.ChkAugmentAADRoles.Checked = Settings.EnableAugmentation;
            this.ChkFilterSecurityEnabledGroupsOnly.Checked = Settings.FilterSecurityEnabledGroupsOnly;

            // Other settings
            this.ChkAlwaysResolveUserInput.Checked = Settings.AlwaysResolveUserInput;
            this.ChkFilterExactMatchOnly.Checked = Settings.FilterExactMatchOnly;
            this.InputProxyAddress.Text = Settings.ProxyAddress;
            this.InputRestrictSearchableUsersByGroups.Text = Settings.RestrictSearchableUsersByGroups;
            this.InputTenantDataCacheLifetimeInMinutes.Text = Settings.TenantDataCacheLifetimeInMinutes.ToString();

            AzureCloudName[] azureCloudNames = (AzureCloudName[])Enum.GetValues(typeof(AzureCloudName));
            foreach (var azureCloudName in azureCloudNames)
            {
                this.DDLAzureCloudInstance.Items.Add(new ListItem(azureCloudName.ToString(), azureCloudName.ToString()));
            }
            this.DDLAzureCloudInstance.SelectedValue = AzureCloudName.AzureGlobal.ToString();
        }

        private void PopulateGraphPropertiesLists()
        {
            this.DdlUserGraphPropertyToDisplay.Items.Add(new ListItem("(Same as the identifier property)", "NotSet"));
            this.DdlGroupGraphPropertyToDisplay.Items.Add(new ListItem("(Same as the identifier property)", "NotSet"));

            DirectoryObjectProperty[] aadPropValues = (DirectoryObjectProperty[])Enum.GetValues(typeof(DirectoryObjectProperty));
            IEnumerable<DirectoryObjectProperty> aadPropValuesSorted = aadPropValues.OrderBy(v => v.ToString());
            foreach (DirectoryObjectProperty prop in aadPropValuesSorted)
            {
                // Test property exists in type User, to populate lists of user properties
                if (Utils.GetDirectoryObjectPropertyValue(new User(), prop.ToString()) != null)
                {
                    // Ensure property is of type System.String
                    PropertyInfo pi = typeof(User).GetProperty(prop.ToString());
                    if (pi != null && pi.PropertyType == typeof(String))
                    {
                        this.DdlUserIdDirectoryPropertyMembers.Items.Add(new ListItem(prop.ToString(), ((int)prop).ToString()));
                        this.DdlUserIdDirectoryPropertyGuests.Items.Add(new ListItem(prop.ToString(), ((int)prop).ToString()));
                        this.DdlUserGraphPropertyToDisplay.Items.Add(new ListItem(prop.ToString(), ((int)prop).ToString()));
                    }
                }

                // Test property exists in type Group, to populate lists of group properties
                if (Utils.GetDirectoryObjectPropertyValue(new Group(), prop.ToString()) != null)
                {
                    // Ensure property is of type System.String
                    PropertyInfo pi = typeof(Group).GetProperty(prop.ToString());
                    if (pi != null && pi.PropertyType == typeof(String))
                    {
                        this.DdlGroupDirectoryProperty.Items.Add(new ListItem(prop.ToString(), ((int)prop).ToString()));
                        this.DdlGroupGraphPropertyToDisplay.Items.Add(new ListItem(prop.ToString(), ((int)prop).ToString()));
                    }
                }
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
                Logger.Log($"Microsoft Entra ID tenant '{tenantToRemove.Name}' was successfully removed from configuration '{ConfigurationName}'", TraceSeverity.Medium, TraceCategory.Configuration);
                LabelMessage.Text = String.Format(TextSummaryPersistedObjectInformation, Configuration.Name, Configuration.Version, Configuration.Id);
                PopulateConnectionsGrid();
            }
        }

        protected bool UpdateConfiguration(bool commitChanges)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) { return false; }

            // User identifier settings
            Settings.ClaimTypes.UpdateUserIdentifier((DirectoryObjectProperty)Convert.ToInt32(this.DdlUserIdDirectoryPropertyMembers.SelectedValue));
            Settings.ClaimTypes.UpdateIdentifierForGuestUsers((DirectoryObjectProperty)Convert.ToInt32(this.DdlUserIdDirectoryPropertyGuests.SelectedValue));
            if (this.DdlUserGraphPropertyToDisplay.SelectedValue != "NotSet")
            {
                Settings.ClaimTypes.UserIdentifierConfig.EntityPropertyToUseAsDisplayText = (DirectoryObjectProperty)Convert.ToInt32(this.DdlUserGraphPropertyToDisplay.SelectedValue);
            }
            else
            {
                Settings.ClaimTypes.UserIdentifierConfig.EntityPropertyToUseAsDisplayText = DirectoryObjectProperty.NotSet;
            }

            // Group identifier settings
            Settings.ClaimTypes.UpdateGroupIdentifier((DirectoryObjectProperty)Convert.ToInt32(this.DdlGroupDirectoryProperty.SelectedValue));
            if (this.DdlGroupGraphPropertyToDisplay.SelectedValue != "NotSet")
            {
                Settings.ClaimTypes.GroupIdentifierConfig.EntityPropertyToUseAsDisplayText = (DirectoryObjectProperty)Convert.ToInt32(this.DdlGroupGraphPropertyToDisplay.SelectedValue);
            }
            else
            {
                Settings.ClaimTypes.GroupIdentifierConfig.EntityPropertyToUseAsDisplayText = DirectoryObjectProperty.NotSet;
            }
            Settings.EnableAugmentation = this.ChkAugmentAADRoles.Checked;
            Settings.FilterSecurityEnabledGroupsOnly = this.ChkFilterSecurityEnabledGroupsOnly.Checked;

            // Other settings
            Settings.AlwaysResolveUserInput = this.ChkAlwaysResolveUserInput.Checked;
            Settings.FilterExactMatchOnly = this.ChkFilterExactMatchOnly.Checked;
            Settings.ProxyAddress = this.InputProxyAddress.Text;
            Settings.RestrictSearchableUsersByGroups = this.InputRestrictSearchableUsersByGroups.Text;
            Settings.TenantDataCacheLifetimeInMinutes = Convert.ToInt32(this.InputTenantDataCacheLifetimeInMinutes.Text);

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

            EntraIDTenant newTenant = new EntraIDTenant
            {
                Name = this.TxtTenantName.Text,
                AzureCloud = (AzureCloudName)Enum.Parse(typeof(AzureCloudName), this.DDLAzureCloudInstance.SelectedValue),
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

            if (Settings.EntraIDTenants == null)
            {
                Settings.EntraIDTenants = new List<EntraIDTenant>();
            }

            var newTenant = new EntraIDTenant
            {
                Name = this.TxtTenantName.Text,
                AzureCloud = (AzureCloudName)Enum.Parse(typeof(AzureCloudName), this.DDLAzureCloudInstance.SelectedValue),
                ExcludeGuestUsers = this.ChkMemberUserTypeOnly.Checked,
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
            Logger.Log($"Microsoft Entra ID tenant '{this.TxtTenantName.Text}' was successfully added to configuration '{ConfigurationName}'", TraceSeverity.Medium, TraceCategory.Configuration);
            LabelMessage.Text = String.Format(TextSummaryPersistedObjectInformation, Configuration.Name, Configuration.Version, Configuration.Id);

            PopulateConnectionsGrid();
            this.TxtTenantName.Text = "TENANTNAME.onMicrosoft.com";
            this.TxtClientId.Text = String.Empty;
            this.TxtClientSecret.Text = String.Empty;
            this.InputClientCertPassword.Text = String.Empty;
            this.TxtExtensionAttributesApplicationId.Text = String.Empty;
            this.DDLAzureCloudInstance.SelectedValue = AzureCloudName.AzureGlobal.ToString();
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

        protected void grdAzureTenants_RowEditing(object sender, GridViewEditEventArgs e)
        {
            grdAzureTenants.EditIndex = e.NewEditIndex;
            PopulateConnectionsGrid();
        }

        protected void grdAzureTenants_RowCancelingEdit(object sender, GridViewCancelEditEventArgs e)
        {
            grdAzureTenants.EditIndex = -1;
            PopulateConnectionsGrid();
        }

        protected void grdAzureTenants_RowUpdating(object sender, GridViewUpdateEventArgs e)
        {
            Guid tenantId = Guid.Parse(e.NewValues["Id"].ToString());
            string newClientId = e.NewValues["ClientID"].ToString();
            TextBox newSecretTextBox = (TextBox)grdAzureTenants.Rows[e.RowIndex].FindControl("EditTenantNewSecret");
            string newSecret = newSecretTextBox.Text;

            EntraIDTenant tenant = Settings.EntraIDTenants.First(x => x.Identifier == tenantId);
            bool tenantUpdated = false;
            if (String.IsNullOrWhiteSpace(newSecret))
            {
                if (tenant.ClientId != newClientId)
                {
                    tenant.ClientId = newClientId;
                    tenantUpdated = true;
                }
            }
            else
            {
                tenant.SetCredentials(newClientId, newSecret);
                tenantUpdated = true;
            }
            if (tenantUpdated)
            {
                CommitChanges();
            }
            grdAzureTenants.EditIndex = -1;
            PopulateConnectionsGrid();
        }

        protected void grdAzureTenants_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            // Ask user for confirmation when cliking on button Delete - https://stackoverflow.com/questions/9026884/asp-net-gridview-delete-row-only-on-confirmation
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                Button deleteButton = (Button)e.Row.Cells[4].Controls[2];
                if (deleteButton != null && String.Equals(deleteButton.Text, "Delete", StringComparison.OrdinalIgnoreCase))
                {
                    deleteButton.OnClientClick = "if(!confirm('Are you sure you want to delete this tenant?')) return;";
                }
            }
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
            PropertyCollection.Columns.Add("AzureCloud", typeof(string));
        }

        public void AddRow(Guid Id, string TenantName, string ClientID, string AuthenticationMode, string AzureCloud)
        {
            DataRow newRow = PropertyCollection.Rows.Add();
            newRow["Id"] = Id;
            newRow["TenantName"] = TenantName;
            newRow["ClientID"] = ClientID;
            //newRow["MemberUserTypeOnly"] = MemberUserTypeOnly;
            newRow["AuthenticationMode"] = AuthenticationMode;
            newRow["AzureCloud"] = AzureCloud;
        }

        public void BindGrid(SPGridView grid)
        {
            grid.DataSource = PropertyCollection.DefaultView;
            grid.DataBind();
        }
    }
}
