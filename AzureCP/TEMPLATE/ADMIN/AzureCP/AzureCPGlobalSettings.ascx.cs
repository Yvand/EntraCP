using Microsoft.Graph;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using static azurecp.ClaimsProviderLogging;

namespace azurecp.ControlTemplates
{
    public partial class AzureCPGlobalSettings : AzureCPUserControl
    {
        string TextErrorAzureTenantFieldsMissing = "Some mandatory fields are missing.";
        string TextErrorTestAzureADConnection = "Unable to get access token for tenant '{0}': {1}";
        string TextConnectionSuccessful = "Connection successful.";

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
                this.BtnOK.Enabled = this.BtnOKTop.Enabled = false;
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
            if (PersistedObject.AzureTenants != null)
            {
                PropertyCollectionBinder pcb = new PropertyCollectionBinder();
                foreach (AzureTenant tenant in PersistedObject.AzureTenants)
                {
                    pcb.AddRow(tenant.Id, tenant.TenantName, tenant.ClientId, tenant.ExcludeMemberUsers);
                }
                pcb.BindGrid(grdAzureTenants);
            }
        }

        private void PopulateFields()
        {
            if (IdentityCTConfig.DirectoryObjectPropertyToShowAsDisplayText == AzureADObjectProperty.NotSet)
            {
                this.RbIdentityDefault.Checked = true;
            }
            else
            {
                this.RbIdentityCustomGraphProperty.Checked = true;
                this.DDLGraphPropertyToDisplay.Items.FindByValue(((int)IdentityCTConfig.DirectoryObjectPropertyToShowAsDisplayText).ToString()).Selected = true;
            }
            this.DDLDirectoryPropertyMemberUsers.Items.FindByValue(((int)IdentityCTConfig.DirectoryObjectProperty).ToString()).Selected = true;
            this.DDLDirectoryPropertyGuestUsers.Items.FindByValue(((int)IdentityCTConfig.DirectoryObjectPropertyForGuestUsers).ToString()).Selected = true;
            this.ChkAlwaysResolveUserInput.Checked = PersistedObject.AlwaysResolveUserInput;
            this.ChkFilterExactMatchOnly.Checked = PersistedObject.FilterExactMatchOnly;
            this.ChkAugmentAADRoles.Checked = PersistedObject.EnableAugmentation;
            this.ChkFilterSecurityEnabledGroupsOnly.Checked = PersistedObject.FilterSecurityEnabledGroupsOnly;
        }

        private void BuildGraphPropertyDDL()
        {
            foreach (AzureADObjectProperty prop in Enum.GetValues(typeof(AzureADObjectProperty)))
            {
                // Ensure property exists for the User object type
                if (AzureCP.GetPropertyValue(new User(), prop.ToString()) == null) continue;

                // Ensure property is of type System.String
                PropertyInfo pi = typeof(User).GetProperty(prop.ToString());
                if (pi == null) continue;
                if (pi.PropertyType != typeof(System.String)) continue;

                //System.Web.UI.WebControls.ListItem listItem = new System.Web.UI.WebControls.ListItem(prop.ToString(), ((int)prop).ToString());
                this.DDLGraphPropertyToDisplay.Items.Add(new System.Web.UI.WebControls.ListItem(prop.ToString(), ((int)prop).ToString()));
                this.DDLDirectoryPropertyMemberUsers.Items.Add(new System.Web.UI.WebControls.ListItem(prop.ToString(), ((int)prop).ToString()));
                this.DDLDirectoryPropertyGuestUsers.Items.Add(new System.Web.UI.WebControls.ListItem(prop.ToString(), ((int)prop).ToString()));
            }
        }

        protected void grdAzureTenants_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) return;
            if (PersistedObject.AzureTenants == null) return;

            GridViewRow rowToDelete = grdAzureTenants.Rows[e.RowIndex];
            Guid Id = new Guid(rowToDelete.Cells[0].Text);
            AzureTenant tenantToRemove = PersistedObject.AzureTenants.FirstOrDefault(x => x.Id == Id);
            if (tenantToRemove != null)
            {
                PersistedObject.AzureTenants.Remove(tenantToRemove);
                CommitChanges();
                ClaimsProviderLogging.Log($"Azure AD tenant '{tenantToRemove.TenantName}' was successfully removed from configuration '{PersistedObjectName}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Configuration);
                PopulateConnectionsGrid();
            }
        }

        protected bool UpdateConfiguration(bool commitChanges)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) return false;

            if (this.RbIdentityCustomGraphProperty.Checked)
            {
                IdentityCTConfig.DirectoryObjectPropertyToShowAsDisplayText = (AzureADObjectProperty)Convert.ToInt32(this.DDLGraphPropertyToDisplay.SelectedValue);
            }
            else
            {
                IdentityCTConfig.DirectoryObjectPropertyToShowAsDisplayText = AzureADObjectProperty.NotSet;
            }

            AzureADObjectProperty newUserIdentifier = (AzureADObjectProperty)Convert.ToInt32(this.DDLDirectoryPropertyMemberUsers.SelectedValue);
            if (newUserIdentifier != AzureADObjectProperty.NotSet)
                PersistedObject.ClaimTypes.UpdateUserIdentifier(newUserIdentifier);

            AzureADObjectProperty newIdentifierForGuestUsers = (AzureADObjectProperty)Convert.ToInt32(this.DDLDirectoryPropertyGuestUsers.SelectedValue);
            if (newIdentifierForGuestUsers != AzureADObjectProperty.NotSet)
                PersistedObject.ClaimTypes.UpdateIdentifierForGuestUsers(newIdentifierForGuestUsers);

            PersistedObject.AlwaysResolveUserInput = this.ChkAlwaysResolveUserInput.Checked;
            PersistedObject.FilterExactMatchOnly = this.ChkFilterExactMatchOnly.Checked;
            PersistedObject.EnableAugmentation = this.ChkAugmentAADRoles.Checked;
            PersistedObject.FilterSecurityEnabledGroupsOnly = this.ChkFilterSecurityEnabledGroupsOnly.Checked;

            if (commitChanges) CommitChanges();
            return true;
        }

        protected void BtnTestAzureTenantConnection_Click(Object sender, EventArgs e)
        {
            this.ValidateAzureTenantConnection();
        }

        protected void ValidateAzureTenantConnection()
        {
            if (this.TxtTenantName.Text == String.Empty || this.TxtClientId.Text == String.Empty || this.TxtClientSecret.Text == String.Empty)
            {
                this.LabelErrorTestLdapConnection.Text = TextErrorAzureTenantFieldsMissing;
                return;
            }

            string tenantName = this.TxtTenantName.Text;
            string clientId = this.TxtClientId.Text;
            string clientSecret = this.TxtClientSecret.Text;
            try
            {
                AADAppOnlyAuthenticationProvider testConnection = new AADAppOnlyAuthenticationProvider(ClaimsProviderConstants.AuthorityUriTemplate, tenantName, clientId, clientSecret, String.Empty, ClaimsProviderConstants.DEFAULT_TIMEOUT);
                Task<bool> testConnectionTask = testConnection.GetAccessToken(true);
                testConnectionTask.Wait();
                this.LabelTestTenantConnectionOK.Text = TextConnectionSuccessful;
            }
            catch (AdalServiceException ex)
            {
                this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, tenantName, ex.Message);
            }
            catch (Exception ex)
            {
                this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, tenantName, ex.Message);
            }
        }

        protected void BtnOK_Click(Object sender, EventArgs e)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) return;
            if (UpdateConfiguration(true)) Response.Redirect("/Security.aspx", false);
            else LabelErrorMessage.Text = base.MostImportantError;
        }

        protected void BtnResetAzureCPConfig_Click(Object sender, EventArgs e)
        {
            AzureCPConfig.DeleteConfiguration(PersistedObjectName);
            Response.Redirect(Request.RawUrl, false);
        }

        private TableCell GetTableCell(string Value)
        {
            TableCell tc = new TableCell();
            tc.Text = Value;
            return tc;
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
            if (ValidatePrerequisite() != ConfigStatus.AllGood) return;
            if (this.TxtTenantName.Text == String.Empty || this.TxtClientId.Text == String.Empty || this.TxtClientSecret.Text == String.Empty)
            {
                this.LabelErrorTestLdapConnection.Text = TextErrorAzureTenantFieldsMissing;
                return;
            }

            if (PersistedObject.AzureTenants == null) PersistedObject.AzureTenants = new List<AzureTenant>();
            this.PersistedObject.AzureTenants.Add(
                new AzureTenant
                {
                    TenantName = this.TxtTenantName.Text,
                    ClientId = TxtClientId.Text,
                    ClientSecret = this.TxtClientSecret.Text,
                    ExcludeMemberUsers = this.ChkMemberUserTypeOnly.Checked,
                });

            CommitChanges();
            ClaimsProviderLogging.Log($"Azure AD tenant '{this.TxtTenantName.Text}' was successfully added in configuration '{PersistedObjectName}'", TraceSeverity.Medium, EventSeverity.Information, TraceCategory.Configuration);

            PopulateConnectionsGrid();
            this.TxtTenantName.Text = "TENANTNAME.onMicrosoft.com";
            this.TxtClientId.Text = this.TxtClientSecret.Text = String.Empty;
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
            PropertyCollection.Columns.Add("MemberUserTypeOnly", typeof(bool));
        }

        public void AddRow(Guid Id, string TenantName, string ClientID, bool MemberUserTypeOnly)
        {
            DataRow newRow = PropertyCollection.Rows.Add();
            newRow["Id"] = Id;
            newRow["TenantName"] = TenantName;
            newRow["ClientID"] = ClientID;
            newRow["MemberUserTypeOnly"] = MemberUserTypeOnly;
        }

        public void BindGrid(SPGridView grid)
        {
            grid.DataSource = PropertyCollection.DefaultView;
            grid.DataBind();
        }
    }
}
