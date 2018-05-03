using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using System.Web.UI.WebControls;
using System.Linq;
using Microsoft.Graph;
using static azurecp.ClaimsProviderLogging;
using System.Web.UI;

namespace azurecp.ControlTemplates
{
    public partial class AzureCPGlobalSettings : AzureCPUserControl
    {
        string TextErrorAzureTenantFieldsMissing = "Some mandatory fields are missing.";

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
                    pcb.AddRow(tenant.Id, tenant.TenantName, tenant.ClientId, tenant.MemberUserTypeOnly);
                }
                pcb.BindGrid(grdAzureTenants);
            }
        }

        private void PopulateFields()
        {
            if (IdentityClaim.DirectoryObjectPropertyToShowAsDisplayText == AzureADObjectProperty.NotSet)
            {
                this.RbIdentityDefault.Checked = true;
            }
            else
            {
                this.RbIdentityCustomGraphProperty.Checked = true;
                this.DDLGraphPropertyToDisplay.Items.FindByValue(((int)IdentityClaim.DirectoryObjectPropertyToShowAsDisplayText).ToString()).Selected = true;
            }
            this.ChkAlwaysResolveUserInput.Checked = PersistedObject.AlwaysResolveUserInput;
            this.ChkFilterExactMatchOnly.Checked = PersistedObject.FilterExactMatchOnly;
            this.ChkAugmentAADRoles.Checked = PersistedObject.EnableAugmentation;
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

                this.DDLGraphPropertyToDisplay.Items.Add(new System.Web.UI.WebControls.ListItem(prop.ToString(), ((int)prop).ToString()));
            }
        }

        protected void grdAzureTenants_RowDeleting(object sender, GridViewDeleteEventArgs e)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) return;
            if (PersistedObject.AzureTenants == null) return;
            GridViewRow rowToDelete = grdAzureTenants.Rows[e.RowIndex];

            Guid Id = new Guid(rowToDelete.Cells[0].Text);
            PersistedObject.AzureTenants.Remove(PersistedObject.AzureTenants.Find(x => x.Id == Id));

            // Update object in database
            //UpdatePersistedObject();
            CommitChanges();
            ClaimsProviderLogging.Log(
                    String.Format("Removed an Azure tenant in PersistedObject {0}", ClaimsProviderConstants.AZURECPCONFIG_NAME),
                    TraceSeverity.Medium,
                    EventSeverity.Information,
                    TraceCategory.Configuration);

            PopulateConnectionsGrid();
        }

        /// <summary>
        /// Update global configuration of AzureCP, except LDAP connections
        /// </summary>
        protected void UpdateTrustConfiguration()
        {
            //if (!this.AllowPersistedObjectUpdate) return;
            //if (null == PersistedObject)
            //{
            //    AzureCPLogging.Log(
            //        String.Format("PersistedObject {0} should not be null.", Constants.AZURECPCONFIG_NAME),
            //        TraceSeverity.Unexpected,
            //        EventSeverity.Error,
            //        TraceCategory.Configuration);
            //    return;
            //}

            //if (null == CurrentTrustedLoginProvider)
            //{
            //    AzureCPLogging.Log(
            //        "Trust associated with AzureCP could not be found.",
            //        TraceSeverity.Unexpected,
            //        EventSeverity.Error,
            //        TraceCategory.Configuration);
            //    return;
            //}

            // Handle identity claim type
            if (this.RbIdentityCustomGraphProperty.Checked)
            {
                IdentityClaim.DirectoryObjectPropertyToShowAsDisplayText = (AzureADObjectProperty)Convert.ToInt32(this.DDLGraphPropertyToDisplay.SelectedValue);
            }
            else
            {
                IdentityClaim.DirectoryObjectPropertyToShowAsDisplayText = AzureADObjectProperty.NotSet;
            }

            PersistedObject.AlwaysResolveUserInput = this.ChkAlwaysResolveUserInput.Checked;
            PersistedObject.FilterExactMatchOnly = this.ChkFilterExactMatchOnly.Checked;
            PersistedObject.EnableAugmentation = this.ChkAugmentAADRoles.Checked;

            //UpdatePersistedObject();
            //AzureCPLogging.Log(
            //    String.Format("Updated PersistedObject {0}", Constants.AZURECPCONFIG_NAME),
            //    TraceSeverity.Medium,
            //    EventSeverity.Information,
            //    TraceCategory.Configuration);
        }

        protected override bool UpdatePersistedObjectProperties(bool commitChanges)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood) return false;
            //UpdateLdapSettings();
            //UpdateAugmentationSettings();
            //UpdateGeneralSettings();
            UpdateTrustConfiguration();
            if (commitChanges) CommitChanges();
            return true;
        }
        ///// <summary>
        ///// Commit changes made to persisted object in configuration database
        ///// </summary>
        //void UpdatePersistedObject()
        //{
        //    // Update object in database
        //    SPSecurity.RunWithElevatedPrivileges(delegate ()
        //    {
        //        this.Web.AllowUnsafeUpdates = true;
        //        PersistedObject.Update();
        //        this.Web.AllowUnsafeUpdates = false;
        //    });
        //    ViewState["PersistedObjectVersion"] = PersistedObject.Version;
        //}

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

            //ActiveDirectoryClient activeDirectoryClient;
            //try
            //{
            //    string tenantName = this.TxtTenantName.Text;
            //    string tenantId = this.TxtTenantId.Text;
            //    string clientId = this.TxtClientId.Text;
            //    string clientSecret = this.TxtClientSecret.Text;

            //    // Get access token
            //    activeDirectoryClient = AuthenticationHelper.GetActiveDirectoryClientAsApplication(tenantName, tenantId, clientId, clientSecret);
            //    // Get information on tenant
            //    ITenantDetail tenant = activeDirectoryClient.TenantDetails
            //        .Where(tDetail => tDetail.ObjectId.Equals(tenantId))
            //        .ExecuteAsync()
            //        .Result.CurrentPage.FirstOrDefault();
            //    if (tenant != null)
            //    {
            //        this.LabelTestTenantConnectionOK.Text = TextConnectionSuccessful;
            //        this.LabelTestTenantConnectionOK.Text += "<br>" + tenant.DisplayName;
            //    }
            //    else
            //    {
            //        this.LabelErrorTestLdapConnection.Text = TextErrorTestAzureADConnectionTenantNotFound = "Tenant was not found.";
            //    }
            //    activeDirectoryClient = null;
            //}
            //catch (AuthenticationException ex)
            //{
            //    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
            //    //InnerException Message will contain the HTTP error status codes mentioned in the link above
            //    this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, ex.Message);
            //    if (ex.InnerException != null)
            //    {
            //        this.LabelErrorTestLdapConnection.Text += String.Format("<br>Error detail: {0}", ex.InnerException.Message);
            //    }
            //    AzureCPLogging.LogException("AzureCP", "while testing connectivity", AzureCPLogging.Categories.Configuration, ex);
            //}
            //catch (Exception ex)
            //{
            //    //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
            //    //InnerException Message will contain the HTTP error status codes mentioned in the link above
            //    this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, ex.Message);
            //    if (ex.InnerException != null)
            //    {
            //        this.LabelErrorTestLdapConnection.Text += String.Format("<br>Error detail: {0}", ex.InnerException.Message);
            //    }
            //    AzureCPLogging.LogException("AzureCP", "while testing connectivity", AzureCPLogging.Categories.Configuration, ex);
            //}

            //try
            //{

            //    // get OAuth AccessToken using Client Credentials
            //    string tenantName = this.TxtTenantName.Text;
            //    string authString = "https://login.windows.net/" + tenantName;

            //    AuthenticationContext authenticationContext = new AuthenticationContext(authString, false);

            //    // Config for OAuth client credentials 
            //    string clientId = this.TxtClientId.Text;
            //    string clientSecret = this.TxtClientSecret.Text;
            //    ClientCredential clientCred = new ClientCredential(clientId, clientSecret);
            //    string resource = "https://graph.windows.net";

            //    AuthenticationResult authenticationResult = authenticationContext.AcquireToken(resource, clientCred);
            //    string accessToken = authenticationResult.AccessToken;

            //    GraphConnection graphConnection;
            //    Guid ClientRequestId = Guid.NewGuid();
            //    GraphSettings graphSettings = new GraphSettings();
            //    graphSettings.ApiVersion = "2013-11-08";
            //    graphSettings.GraphDomainName = "graph.windows.net";
            //    graphConnection = new GraphConnection(accessToken, ClientRequestId, graphSettings);

            //    this.LabelTestTenantConnectionOK.Text = TextConnectionSuccessful;
            //}
            //catch (Exception ex)
            //{
            //    this.LabelErrorTestLdapConnection.Text = String.Format(TextErrorTestAzureADConnection, ex.Message);
            //}
        }

        protected void BtnOK_Click(Object sender, EventArgs e)
        {
            //this.UpdateTrustConfiguration();
            //Response.Redirect("/", false);

            if (ValidatePrerequisite() != ConfigStatus.AllGood) return;
            if (UpdatePersistedObjectProperties(true)) Response.Redirect("/Security.aspx", false);
            else LabelErrorMessage.Text = MostImportantError;
        }

        protected void BtnResetAzureCPConfig_Click(Object sender, EventArgs e)
        {
            AzureCPConfig.DeleteAzureCPConfig(PersistedObjectName);
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
        /// Add new LDAP connection to collection in persisted object
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
                    MemberUserTypeOnly = this.ChkMemberUserTypeOnly.Checked,
                });

            // Update object in database
            //UpdatePersistedObject();
            CommitChanges();
            ClaimsProviderLogging.Log(
                   String.Format("Added a new Azure tenant in PersistedObject {0}", ClaimsProviderConstants.AZURECPCONFIG_NAME),
                   TraceSeverity.Medium,
                   EventSeverity.Information,
                   TraceCategory.Configuration);

            PopulateConnectionsGrid();
            this.TxtClientId.Text = this.TxtClientSecret.Text = String.Empty;
            this.TxtTenantName.Text = "TENANTNAME.onMicrosoft.com";
        }

        //public static Dictionary<int, string> EnumToList(Type t)
        //{
        //    Dictionary<int, string> list = new Dictionary<int, string>();
        //    foreach (var v in Enum.GetValues(t))
        //    {
        //        string name = Enum.GetName(t, (int)v);
        //        // Encryption and SecureSocketsLayer have same value and it will violate uniqueness of key if attempt to add both to Dictionary
        //        if (String.Equals(name, "Encryption", StringComparison.InvariantCultureIgnoreCase) && list.ContainsValue("Encryption")) continue;
        //        list.Add((int)v, name);
        //    }
        //    return list;
        //}
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
