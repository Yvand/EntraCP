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
using static azurecp.AzureCPLogging;
using System.Web.UI;

namespace azurecp.ControlTemplates
{
    public partial class AzureCPGlobalSettings : AzureCPUserControl
    {
        //public string PersistedObjectName = Constants.AZURECPCONFIG_NAME;
        //public Guid PersistedObjectID = new Guid(Constants.AZURECPCONFIG_ID);
        //SPTrustedLoginProvider CurrentTrustedLoginProvider;
        //AzureCPConfig PersistedObject;
        //AzureADObject IdentityClaim;
        //bool AllowPersistedObjectUpdate = true;

        string TextErrorNoTrustAssociation = "AzureCP is currently not associated with any TrustedLoginProvider. It is mandatory because it cannot create permission for a trust if it is not associated to it.<br/>Visit <a href=\"https://github.com/Yvand/AzureCP\" target=\"_blank\">https://github.com/Yvand/AzureCP</a> for documentation.<br/>Settings on this page will not be available as long as AzureCP will not associated to a trut.";
        string TextErrorAzureTenantFieldsMissing = "Some mandatory fields are missing.";
        //string TextErrorTestAzureADConnection = "Unable to connect to Azure tenant<br/>It may be expected if w3wp process of central admin has intentionally no access to Azure.<br/>{0}";
        //string TextErrorTestAzureADConnectionTenantNotFound = "Tenant was not found.";
        //string TextConnectionSuccessful = "Connection successful.";
        //string TextErrorNoIdentityClaimType = "The TrustedLoginProvider {0} is set with identity claim type \"{1}\" but it is not in the claims list of AzureCP.<br/>Please visit AzureCP page \"claims mapping\" in Security tab to set it and return to this page afterwards.";
        //string TextErrorPersistedObjectStale = "Modification is cancelled because persisted object was modified since last load of the page. Please refresh the page and try again.";

        protected void Page_Load(object sender, EventArgs e)
        {
            // Get trust currently associated with AzureCP, if any
            //CurrentTrustedLoginProvider = AzureCP.GetSPTrustAssociatedWithCP(AzureCP._ProviderInternalName);
            //if (null == CurrentTrustedLoginProvider)
            //{
            //    // Claim provider is currently not associated with any trust.
            //    // Display a message in the page and disable controls
            //    this.LabelErrorMessage.Text = TextErrorNoTrustAssociation;
            //    this.BtnOK.Enabled = this.BtnOKTop.Enabled = this.BtnAddLdapConnection.Enabled = this.BtnTestAzureTenantConnection.Enabled = false;
            //    this.AllowPersistedObjectUpdate = false;
            //    return;
            //}

            //SPSecurity.RunWithElevatedPrivileges(delegate ()
            //{
            //    // Get SPPersisted Object and create it if it doesn't exist
            //    PersistedObject = AzureCPConfig.GetConfiguration(PersistedObjectName);
            //    if (PersistedObject == null)
            //    {
            //        this.Web.AllowUnsafeUpdates = true;
            //        PersistedObject = AzureCPConfig.CreatePersistedObject(PersistedObjectID.ToString(), PersistedObjectName);
            //        this.Web.AllowUnsafeUpdates = false;
            //    }
            //});

            if (ValidatePrerequisite() != ConfigStatus.AllGood)
            {
                this.LabelErrorMessage.Text = base.MostImportantError;
                this.BtnOK.Enabled = this.BtnOKTop.Enabled = false;
                return;
            }

            //this.IdentityClaim = PersistedObject.AzureADObjects.Find(x => String.Equals(CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase) && !x.CreateAsIdentityClaim);
            //if (null == this.IdentityClaim)
            //{
            //    // Identity claim type is missing in the attributes list
            //    this.LabelErrorMessage.Text = String.Format(TextErrorNoIdentityClaimType, CurrentTrustedLoginProvider.DisplayName, CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType);
            //    this.BtnOK.Enabled = this.BtnOKTop.Enabled = this.BtnAddLdapConnection.Enabled = this.BtnTestAzureTenantConnection.Enabled = false;
            //    return;
            //}

            //if (ViewState["PersistedObjectVersion"] == null)
            //    ViewState.Add("PersistedObjectVersion", PersistedObject.Version);
            //if ((long)ViewState["PersistedObjectVersion"] != PersistedObject.Version)
            //{
            //    // PersistedObject changed since last time. Should not allow any update
            //    this.LabelErrorMessage.Text = TextErrorPersistedObjectStale;
            //    this.AllowPersistedObjectUpdate = false;
            //    return;
            //}

            if (!this.IsPostBack) Initialize();
        }

        protected void Initialize()
        {
            BuildGraphPropertyDDL();
            PopulateLdapConnectionGrid();
            PopulateFields();
        }

        void PopulateLdapConnectionGrid()
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
            if (IdentityClaim.GraphPropertyToDisplay == GraphProperty.None)
            {
                this.RbIdentityDefault.Checked = true;
            }
            else
            {
                this.RbIdentityCustomGraphProperty.Checked = true;
                this.DDLGraphPropertyToDisplay.Items.FindByValue(((int)IdentityClaim.GraphPropertyToDisplay).ToString()).Selected = true;
            }
            this.ChkAlwaysResolveUserInput.Checked = PersistedObject.AlwaysResolveUserInput;
            this.ChkFilterExactMatchOnly.Checked = PersistedObject.FilterExactMatchOnly;
            this.ChkAugmentAADRoles.Checked = PersistedObject.AugmentAADRoles;
        }

        private void BuildGraphPropertyDDL()
        {
            foreach (GraphProperty prop in Enum.GetValues(typeof(GraphProperty)))
            {
                // Ensure property exists for the User object type
                if (AzureCP.GetGraphPropertyValue(new User(), prop.ToString()) == null) continue;

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
            AzureCPLogging.Log(
                    String.Format("Removed an Azure tenant in PersistedObject {0}", Constants.AZURECPCONFIG_NAME),
                    TraceSeverity.Medium,
                    EventSeverity.Information,
                    TraceCategory.Configuration);

            PopulateLdapConnectionGrid();
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
                IdentityClaim.GraphPropertyToDisplay = (GraphProperty)Convert.ToInt32(this.DDLGraphPropertyToDisplay.SelectedValue);
            }
            else
            {
                IdentityClaim.GraphPropertyToDisplay = GraphProperty.None;
            }

            PersistedObject.AlwaysResolveUserInput = this.ChkAlwaysResolveUserInput.Checked;
            PersistedObject.FilterExactMatchOnly = this.ChkFilterExactMatchOnly.Checked;
            PersistedObject.AugmentAADRoles = this.ChkAugmentAADRoles.Checked;

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
            if (this.TxtTenantName.Text == String.Empty || this.TxtTenantId.Text == String.Empty || this.TxtClientId.Text == String.Empty || this.TxtClientSecret.Text == String.Empty)
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

            if (this.TxtTenantName.Text == String.Empty || this.TxtTenantId.Text == String.Empty || this.TxtClientId.Text == String.Empty || this.TxtClientSecret.Text == String.Empty)
            {
                this.LabelErrorTestLdapConnection.Text = TextErrorAzureTenantFieldsMissing;
                return;
            }

            if (PersistedObject.AzureTenants == null) PersistedObject.AzureTenants = new List<AzureTenant>();
            this.PersistedObject.AzureTenants.Add(
                new AzureTenant
                {
                    TenantName = this.TxtTenantName.Text,
                    TenantId = this.TxtTenantId.Text,
                    ClientId = TxtClientId.Text,
                    ClientSecret = this.TxtClientSecret.Text,
                    MemberUserTypeOnly = this.ChkMemberUserTypeOnly.Checked,
                });

            // Update object in database
            //UpdatePersistedObject();
            CommitChanges();
            AzureCPLogging.Log(
                   String.Format("Added a new Azure tenant in PersistedObject {0}", Constants.AZURECPCONFIG_NAME),
                   TraceSeverity.Medium,
                   EventSeverity.Information,
                   TraceCategory.Configuration);

            PopulateLdapConnectionGrid();
            this.TxtTenantId.Text = this.TxtClientId.Text = this.TxtClientSecret.Text = String.Empty;
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
