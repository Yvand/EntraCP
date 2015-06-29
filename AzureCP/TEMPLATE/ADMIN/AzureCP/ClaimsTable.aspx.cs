using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace azurecp
{
    public partial class ClaimsTablePage : LayoutsPageBase
    {
        SPTrustedLoginProvider CurrentTrustedLoginProvider;
        AzureCPConfig PersistedObject;
        List<KeyValuePair<int, AzureADObject>> ClaimsMapping;
        bool AllowPersistedObjectUpdate = true;
        public bool ShowNewItemForm = false;
        public bool HideAllContent = false;
        public string TrustName = String.Empty;

        string TextErrorNoTrustAssociation = "AzureCP is currently not associated with any TrustedLoginProvider. It is mandatory because it cannot create permission for a trust if it is not associated to it.<br/>Visit <a href=\"http://ldapcp.codeplex.com/\" target=\"_blank\">http://ldapcp.codeplex.com/</a> to see how to associate it.<br/>Settings on this page will not be available as long as AzureCP will not associated to a trut.";
        string TextErrorFieldsMissing = "Some mandatory fields are missing.";
        string TextErrorDuplicateClaimType = "This claim type already exists in the list, you cannot create duplicates.";
        string TextErrorUpdateEmptyClaimType = "You tried to update item {0} with an empty claim type, which is not allowed.";
        string TextErrorUpdateItemDuplicate = "You tried to update item {0} with a {1} that already exists ({2}). Duplicates are not allowed.";
        string TextErrorUpdateIdentityClaimTypeChanged = "You cannot change claim type of identity claim.";
        string TextErrorUpdateIdentityClaimEntityTypeNotUser = "Identity claim must be set to SPClaimEntityTypes.User.";
        string TextErrorPersistedObjectStale = "Modification is cancelled because persisted object was modified since page was loaded. Please refresh the page and try again.";
        string TextErrorNoIdentityClaimType = "The TrustedLoginProvider {0} is set with identity claim type \"{1}\" but it is not in the claims list below. AzureCP will not work until you add this claim type in this list.";
        string TextErrorNewMetadataAlreadyUsed = "Metadata {0} is already used. Duplicates are not allowed.";

        string HtmlCellClaimType = "<span name=\"span_claimtype_{1}\" id=\"span_claimtype_{1}\">{0}</span><input name=\"input_claimtype_{1}\" id=\"input_claimtype_{1}\" style=\"display: none; width: 90%;\" value=\"{0}\"></input>";
        string HtmlCellGraphProperty = "<span name=\"span_graphproperty_{1}\" id=\"span_graphproperty_{1}\">{0}</span><select name=\"list_graphproperty_{1}\" id=\"list_graphproperty_{1}\" style=\"display:none;\" value=\"{0}\">{2}</select>";
        string HtmlCellGraphPropertyToDisplay = "<span name=\"span_GraphPropertyToDisplay_{1}\" id=\"span_GraphPropertyToDisplay_{1}\">{0}</span><select name=\"list_GraphPropertyToDisplay_{1}\" id=\"list_GraphPropertyToDisplay_{1}\" style=\"display:none;\" value=\"{0}\">{2}</select>";
        string HtmlCellMetadata = "<span name=\"span_Metadata_{1}\" id=\"span_Metadata_{1}\">{0}</span><select name=\"list_Metadata_{1}\" id=\"list_Metadata_{1}\" style=\"display:none;\">{2}</select>";
        string HtmlCellPrefixToBypassLookup = "<span name=\"span_PrefixToBypassLookup_{1}\" id=\"span_PrefixToBypassLookup_{1}\">{0}</span><input name=\"input_PrefixToBypassLookup_{1}\" id=\"input_PrefixToBypassLookup_{1}\" style=\"display:none;\" value=\"{0}\"></input>";
        string HtmlCellClaimEntityType = "<span name=\"span_ClaimEntityType_{1}\" id=\"span_ClaimEntityType_{1}\">{0}</span><select name=\"list_ClaimEntityType_{1}\" id=\"list_ClaimEntityType_{1}\" style=\"display:none;\">{2}</select>";

        string HtmlEditLink = "<a name=\"editLink_{0}\" id=\"editLink_{0}\" href=\"javascript:Azurecp.ClaimsTablePage.EditItem('{0}')\">Edit</a>";
        string HtmlCancelEditLink = "<a name=\"cancelLink_{0}\" id=\"cancelLink_{0}\" href=\"javascript:Azurecp.ClaimsTablePage.CancelEditItem('{0}')\" style=\"display:none;\">Cancel</a>";

        protected void Page_Load(object sender, EventArgs e)
        {
            // Get trust currently associated with AzureCP, if any
            CurrentTrustedLoginProvider = AzureCP.GetSPTrustAssociatedWithCP(AzureCP._ProviderInternalName);
            if (null == CurrentTrustedLoginProvider)
            {
                // Claim provider is currently not associated with any trust.
                // Display a message in the page and disable controls
                this.LabelErrorMessage.Text = TextErrorNoTrustAssociation;
                this.HideAllContent = true;
                this.BtnCreateNewItem.Visible = false;
                return;
            }

            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                // Get SPPersisted Object and create it if it doesn't exist
                PersistedObject = AzureCPConfig.GetFromConfigDB();
                if (PersistedObject == null)
                {
                    this.Web.AllowUnsafeUpdates = true;
                    PersistedObject = AzureCPConfig.CreatePersistedObject();
                    this.Web.AllowUnsafeUpdates = false;
                }
            });

            if (ViewState["PersistedObjectVersion"] == null)
                ViewState.Add("PersistedObjectVersion", PersistedObject.Version);
            if ((long)ViewState["PersistedObjectVersion"] != PersistedObject.Version)
            {
                // PersistedObject changed since last time. Should not allow any update
                this.LabelErrorMessage.Text = TextErrorPersistedObjectStale;
                this.AllowPersistedObjectUpdate = false;
                return;
            }
            TrustName = CurrentTrustedLoginProvider.Name;

            if (!this.IsPostBack)
            {
                New_DdlPermissionMetadata.Items.Add(String.Empty);
                foreach (object field in typeof(PeopleEditorEntityDataKeys).GetFields())
                {
                    New_DdlPermissionMetadata.Items.Add(((System.Reflection.FieldInfo)field).Name);
                }

                New_DdlGraphProperty.Items.Add(String.Empty);
                New_DdlGraphPropertyToDisplay.Items.Add(String.Empty);
                foreach (object field in typeof(GraphProperty).GetFields())
                {
                    string prop = ((System.Reflection.FieldInfo)field).Name;
                    if (AzureCP.GetGraphPropertyValue(new User(), prop) == null) continue;
                    //if (AzureCP.GetGraphPropertyValue(new Group(), prop) == null) continue;
                    //if (AzureCP.GetGraphPropertyValue(new Role(), prop) == null) continue;

                    New_DdlGraphProperty.Items.Add(prop);
                    New_DdlGraphPropertyToDisplay.Items.Add(prop);
                }
            }

            BuildAttributesListTable(this.IsPostBack);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pendingUpdate">true if there is a post back, which means an update is being made</param>
        private void BuildAttributesListTable(bool pendingUpdate)
        {
            // Copy claims list in a key value pair so that each item has a unique ID that can be used later for update/delete operations
            ClaimsMapping = new List<KeyValuePair<int, AzureADObject>>();
            int i = 0;
            foreach (AzureADObject attr in this.PersistedObject.AzureADObjects)
            {
                ClaimsMapping.Add(new KeyValuePair<int, AzureADObject>(i++, attr));
            }

            bool identityClaimPresent = false;

            TblClaimsMapping.Rows.Clear();
            TableRow tr = new TableRow();
            TableHeaderCell th;
            th = GetTableHeaderCell("Actions");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Claim type");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Property to query");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Property to display");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("<a href='http://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.webcontrols.peopleeditorentitydatakeys_members(v=office.15).aspx' target='_blank'>Metadata</a>");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("<a href='http://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.administration.claims.spclaimentitytypes_members(v=office.15).aspx' target='_blank'>Claim entity type</a>");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Prefix to bypass lookup");
            tr.Cells.Add(th);
            this.TblClaimsMapping.Rows.Add(tr);

            foreach (var attr in this.ClaimsMapping)
            {
                tr = new TableRow();
                bool allowEditItem = String.IsNullOrEmpty(attr.Value.ClaimType) ? false : true;

                // LinkButton must always be created otherwise event receiver will not fire on postback
                TableCell tc = new TableCell();
                if (allowEditItem) tc.Controls.Add(new LiteralControl(String.Format(HtmlEditLink, attr.Key) + "&nbsp;&nbsp;"));
                // But we don't allow to delete identity claim
                if (!String.Equals(attr.Value.ClaimType, CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    LinkButton LnkDeleteItem = new LinkButton();
                    LnkDeleteItem.ID = "DeleteItemLink_" + attr.Key;
                    LnkDeleteItem.Command += LnkDeleteItem_Command;
                    LnkDeleteItem.CommandArgument = attr.Key.ToString();
                    LnkDeleteItem.Text = "Delete";
                    LnkDeleteItem.OnClientClick = "javascript:return confirm('This will delete this item. Do you want to continue?');";
                    if (pendingUpdate) LnkDeleteItem.Visible = false;
                    tc.Controls.Add(LnkDeleteItem);
                }
                LinkButton LnkUpdateItem = new LinkButton();
                LnkUpdateItem.ID = "UpdateItemLink_" + attr.Key;
                LnkUpdateItem.Command += LnkUpdateItem_Command;
                LnkUpdateItem.CommandArgument = attr.Key.ToString();
                LnkUpdateItem.Text = "Save";
                LnkUpdateItem.Style.Add("display", "none");
                if (pendingUpdate) LnkUpdateItem.Visible = false;

                tc.Controls.Add(LnkUpdateItem);
                tc.Controls.Add(new LiteralControl("&nbsp;&nbsp;" + String.Format(HtmlCancelEditLink, attr.Key)));
                tr.Cells.Add(tc);

                // This is just to avoid building the table if we know that there is a pending update, which means it will be rebuilt again after update is complete
                if (!pendingUpdate)
                {
                    string html;
                    TableCell c = null;
                    if (!String.IsNullOrEmpty(attr.Value.ClaimType))
                    {
                        html = String.Format(HtmlCellClaimType, attr.Value.ClaimType, attr.Key);
                        c = GetTableCell(html);
                        allowEditItem = true;
                        if (String.Equals(CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType, attr.Value.ClaimType, StringComparison.InvariantCultureIgnoreCase) && !attr.Value.CreateAsIdentityClaim)
                        {
                            tr.CssClass = "azurecp-rowidentityclaim";
                            identityClaimPresent = true;
                        }
                        else if (CurrentTrustedLoginProvider.ClaimTypeInformation.FirstOrDefault(x => String.Equals(x.MappedClaimType, attr.Value.ClaimType, StringComparison.InvariantCultureIgnoreCase)) == null)
                        {
                            tr.CssClass = "azurecp-rowClaimTypeNotUsedInTrust";
                        }
                        else if (attr.Value.ClaimEntityType != SPClaimEntityTypes.User)
                        {
                            tr.CssClass = "azurecp-rowRoleClaimType";
                        }
                        else
                        {
                            tr.CssClass = "azurecp-rowClaimTypeOk";
                        }
                    }
                    else
                    {
                        c = GetTableCell(attr.Value.CreateAsIdentityClaim ? "linked to identity claim" : "Used as metadata for the permission created");
                    }
                    tr.Cells.Add(c);

                    string htmlCellGraphProperty;
                    string htmlCellGraphPropertyToDisplay;
                    BuildGraphPropertyDDLs(attr, out htmlCellGraphProperty, out htmlCellGraphPropertyToDisplay);

                    html = htmlCellGraphProperty;
                    tr.Cells.Add(GetTableCell(html));

                    html = htmlCellGraphPropertyToDisplay;
                    tr.Cells.Add(GetTableCell(html));

                    MemberInfo[] members;
                    members = typeof(PeopleEditorEntityDataKeys).GetFields(BindingFlags.Static | BindingFlags.Public);
                    html = BuildDDLFromTypeMembers(HtmlCellMetadata, attr, "EntityDataKey", members, true);
                    tr.Cells.Add(GetTableCell(html));

                    members = typeof(SPClaimEntityTypes).GetProperties(BindingFlags.Static | BindingFlags.Public);
                    html = BuildDDLFromTypeMembers(HtmlCellClaimEntityType, attr, "ClaimEntityType", members, false);
                    tr.Cells.Add(GetTableCell(html));

                    html = String.Format(HtmlCellPrefixToBypassLookup, attr.Value.PrefixToBypassLookup, attr.Key);
                    tr.Cells.Add(GetTableCell(html));
                }
                TblClaimsMapping.Rows.Add(tr);
            }

            if (!identityClaimPresent && !pendingUpdate)
            {
                LabelErrorMessage.Text = String.Format(this.TextErrorNoIdentityClaimType, CurrentTrustedLoginProvider.DisplayName, CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType);
            }
        }

        private string BuildDDLFromTypeMembers(string htmlCell, KeyValuePair<int, AzureADObject> attr, string propertyToCheck, MemberInfo[] members, bool addEmptyChoice)
        {
            string option = "<option value=\"{0}\" {1}>{2}</option>";
            string selected = String.Empty;
            bool metadataFound = false;
            StringBuilder options = new StringBuilder();

            // GetValue returns null if object doesn't have a value on this property, using "as string" avoids to throw a NullReference in this case.
            string attrValue = typeof(AzureADObject).GetProperty(propertyToCheck).GetValue(attr.Value) as string;

            // Build DDL based on members retrieved from the supplied type
            foreach (MemberInfo member in members)
            {
                if (String.Equals(attrValue, member.Name, StringComparison.InvariantCultureIgnoreCase))
                {
                    selected = "selected";
                    metadataFound = true;
                }
                else selected = String.Empty;
                options.Append(String.Format(option, member.Name, selected, member.Name));
            }

            if (addEmptyChoice)
            {
                selected = metadataFound ? String.Empty : "selected";
                options = options.Insert(0, String.Format(option, String.Empty, selected, String.Empty));
            }
            return String.Format(htmlCell, attrValue, attr.Key, options.ToString());
        }

        private void BuildGraphPropertyDDLs(KeyValuePair<int, AzureADObject> azureObject, out string htmlCellGraphProperty, out string htmlCellGraphPropertyToDisplay)
        {
            string option = "<option value=\"{0}\" {1}>{2}</option>";
            string graphPropertySelected = String.Empty;
            string graphPropertyToDisplaySelected = String.Empty;
            StringBuilder graphPropertyOptions = new StringBuilder();
            StringBuilder graphPropertyToDisplayOptions = new StringBuilder();
            bool graphPropertyToDisplayFound = false;

            foreach (GraphProperty prop in Enum.GetValues(typeof(GraphProperty)))
            {
                // Ensure property exists for the current object type
                if (azureObject.Value.ClaimEntityType == SPClaimEntityTypes.User)
                {
                    if (AzureCP.GetGraphPropertyValue(new User(), prop.ToString()) == null) continue;
                }
                else
                {
                    if (AzureCP.GetGraphPropertyValue(new Group(), prop.ToString()) == null) continue;
                    //if (AzureCP.GetGraphPropertyValue(new Role(), prop.ToString()) == null) continue;
                }

                graphPropertySelected = azureObject.Value.GraphProperty == prop ? "selected" : String.Empty;

                if (azureObject.Value.GraphPropertyToDisplay == prop)
                {
                    graphPropertyToDisplaySelected = "selected";
                    graphPropertyToDisplayFound = true;
                }
                else graphPropertyToDisplaySelected = String.Empty;

                // Utils.GetPropertyName throws an ArgumentException if GraphProperty == GraphProperty.None
                // Another problem is that Utils.GetPropertyName(prop) returns string with 1st character in lowercase
                //string strProp;
                //if (prop == GraphProperty.None) strProp = "None";
                //else strProp = Utils.GetPropertyName(prop);

                graphPropertyOptions.Append(String.Format(option, prop.ToString(), graphPropertySelected, prop.ToString()));
                graphPropertyToDisplayOptions.Append(String.Format(option, prop.ToString(), graphPropertyToDisplaySelected, prop.ToString()));
            }

            // Insert at 1st position GraphProperty.None in GraphPropertyToDisplay DDL and select it if needed
            string selectNone = graphPropertyToDisplayFound ? String.Empty : "selected";
            graphPropertyToDisplayOptions = graphPropertyToDisplayOptions.Insert(0, String.Format(option, GraphProperty.None, selectNone, GraphProperty.None));

            htmlCellGraphProperty = String.Format(HtmlCellGraphProperty, azureObject.Value.GraphProperty, azureObject.Key, graphPropertyOptions.ToString());
            //string graphPropertyToDisplaySpanDisplay = azureObject.Value.GraphPropertyToDisplay == GraphProperty.None ? String.Empty : azureObject.Value.GraphPropertyToDisplay.ToString();
            htmlCellGraphPropertyToDisplay = String.Format(HtmlCellGraphPropertyToDisplay, azureObject.Value.GraphPropertyToDisplay, azureObject.Key, graphPropertyToDisplayOptions.ToString());
        }

        private TableHeaderCell GetTableHeaderCell(string Value)
        {
            TableHeaderCell tc = new TableHeaderCell();
            tc.Text = Value;
            return tc;
        }
        private TableCell GetTableCell(string Value)
        {
            TableCell tc = new TableCell();
            tc.Text = Value;
            return tc;
        }

        protected void UpdatePersistedObject()
        {
            if (null == PersistedObject)
            {
                AzureCP.LogToULS(
                    String.Format("PersistedObject {0} should not be null.", Constants.AZURECPCONFIG_NAME),
                    TraceSeverity.Unexpected,
                    EventSeverity.Error,
                    AzureCPLogging.Categories.Configuration);
                return;
            }

            if (null == CurrentTrustedLoginProvider)
            {
                AzureCP.LogToULS(
                    "Trust associated with AzureCP could not be found.",
                    TraceSeverity.Unexpected,
                    EventSeverity.Error,
                    AzureCPLogging.Categories.Configuration);
                return;
            }

            // Update object in database
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                this.Web.AllowUnsafeUpdates = true;
                PersistedObject.Update();
                this.Web.AllowUnsafeUpdates = false;

                AzureCP.LogToULS(
                    String.Format("Objects list of AzureCP was successfully updated in PersistedObject {0}.", Constants.AZURECPCONFIG_NAME),
                    TraceSeverity.Medium,
                    EventSeverity.Information,
                    AzureCPLogging.Categories.Configuration);
            });
            ViewState["PersistedObjectVersion"] = PersistedObject.Version;
        }

        void LnkDeleteItem_Command(object sender, CommandEventArgs e)
        {
            if (!this.AllowPersistedObjectUpdate) return;

            string itemId = e.CommandArgument.ToString();
            AzureADObject attr = ClaimsMapping.Find(x => x.Key == Convert.ToInt32(itemId)).Value;
            PersistedObject.AzureADObjects.Remove(attr);
            this.UpdatePersistedObject();
            this.BuildAttributesListTable(false);
        }

        void LnkUpdateItem_Command(object sender, CommandEventArgs e)
        {
            if (!this.AllowPersistedObjectUpdate) return;

            string itemId = e.CommandArgument.ToString();

            NameValueCollection formData = Request.Form;
            if (String.IsNullOrEmpty(formData["input_claimtype_" + itemId]) || String.IsNullOrEmpty(formData["list_graphproperty_" + itemId]))
            {
                this.LabelErrorMessage.Text = TextErrorFieldsMissing;
                return;
            }

            // Get object to update
            int azureObjectId = Convert.ToInt32(itemId);
            AzureADObject azureObject = ClaimsMapping.Find(x => x.Key == azureObjectId).Value;

            // Check if changes are OK
            // Check if claim type is not empty and not already used
            string newClaimType = formData["input_claimtype_" + itemId];
            if (newClaimType == String.Empty)
            {
                this.LabelErrorMessage.Text = TextErrorUpdateEmptyClaimType;
                BuildAttributesListTable(false);
                return;
            }
            List<KeyValuePair<int, AzureADObject>> otherAzureObjects = ClaimsMapping.FindAll(x => x.Key != azureObjectId);
            KeyValuePair<int, AzureADObject> matchFound;
            matchFound = otherAzureObjects.FirstOrDefault(x => String.Equals(x.Value.ClaimType, newClaimType, StringComparison.InvariantCultureIgnoreCase));

            // Check if new claim type is not already used
            if (!matchFound.Equals(default(KeyValuePair<int, AzureADObject>)))
            {
                this.LabelErrorMessage.Text = String.Format(TextErrorUpdateItemDuplicate, azureObject.ClaimType, "claim type", newClaimType);
                BuildAttributesListTable(false);
                return;
            }

            // Check if new entity data key is not already used (we don't care about this check if it's empty)
            string newEntityDataKey = formData["list_Metadata_" + itemId];
            if (newEntityDataKey != String.Empty)
            {
                matchFound = otherAzureObjects.FirstOrDefault(x => String.Equals(x.Value.EntityDataKey, newEntityDataKey, StringComparison.InvariantCultureIgnoreCase));
                if (!matchFound.Equals(default(KeyValuePair<int, AzureADObject>)))
                {
                    this.LabelErrorMessage.Text = String.Format(TextErrorUpdateItemDuplicate, azureObject.ClaimType, "permission metadata", newEntityDataKey);
                    BuildAttributesListTable(false);
                    return;
                }
            }

            string newClaimEntityType = formData["list_ClaimEntityType_" + itemId];
            // Specific checks if current claim type is identity claim type
            if (String.Equals(azureObject.ClaimType, CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase))
            {
                // We don't allow to change claim type
                if (!String.Equals(azureObject.ClaimType, newClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    this.LabelErrorMessage.Text = TextErrorUpdateIdentityClaimTypeChanged;
                    BuildAttributesListTable(false);
                    return;
                }

                // ClaimEntityType must be "SPClaimEntityTypes.User"
                if (!String.Equals(SPClaimEntityTypes.User, newClaimEntityType, StringComparison.InvariantCultureIgnoreCase))
                {
                    this.LabelErrorMessage.Text = TextErrorUpdateIdentityClaimEntityTypeNotUser;
                    BuildAttributesListTable(false);
                    return;
                }
            }

            azureObject.ClaimType = newClaimType;
            azureObject.ClaimEntityType = newClaimEntityType;
            azureObject.PrefixToBypassLookup = formData["input_PrefixToBypassLookup_" + itemId];
            azureObject.EntityDataKey = newEntityDataKey;

            GraphProperty prop;
            bool convertSuccess = Enum.TryParse<GraphProperty>(formData["list_graphproperty_" + itemId], out prop);
            azureObject.GraphProperty = convertSuccess ? prop : azureObject.GraphProperty;

            convertSuccess = Enum.TryParse<GraphProperty>(formData["list_GraphPropertyToDisplay_" + itemId], out prop);
            azureObject.GraphPropertyToDisplay = convertSuccess ? prop : azureObject.GraphPropertyToDisplay;

            this.UpdatePersistedObject();
            this.BuildAttributesListTable(false);
        }

        protected void BtnReset_Click(object sender, EventArgs e)
        {
            AzureCPConfig.ResetClaimsList();
            Response.Redirect(Request.Url.ToString());
        }

        protected void BtnCreateNewItem_Click(object sender, EventArgs e)
        {
            AzureADObject azureObject = new AzureADObject();
            if (RdbNewItemClassicClaimType.Checked)
            {
                if (String.IsNullOrEmpty(New_TxtClaimType.Text))
                {
                    this.LabelErrorMessage.Text = TextErrorFieldsMissing;
                    ShowNewItemForm = true;
                    BuildAttributesListTable(false);
                    return;
                }

                azureObject.ClaimType = New_TxtClaimType.Text;

                if (PersistedObject.AzureADObjects.FirstOrDefault(x => String.Equals(x.ClaimType, azureObject.ClaimType, StringComparison.InvariantCultureIgnoreCase)) != null)
                {
                    this.LabelErrorMessage.Text = TextErrorDuplicateClaimType;
                    ShowNewItemForm = true;
                    BuildAttributesListTable(false);
                    return;
                }
            }
            else if (RdbNewItemPermissionMetadata.Checked)
            {
                if (String.IsNullOrEmpty(New_DdlPermissionMetadata.SelectedValue))
                {
                    this.LabelErrorMessage.Text = TextErrorFieldsMissing;
                    ShowNewItemForm = true;
                    BuildAttributesListTable(false);
                    return;
                }
            }
            else azureObject.CreateAsIdentityClaim = true;

            if (!String.IsNullOrEmpty(New_DdlPermissionMetadata.SelectedValue) && !ClaimsMapping.FirstOrDefault(x => String.Equals(x.Value.EntityDataKey, New_DdlPermissionMetadata.SelectedValue, StringComparison.InvariantCultureIgnoreCase)).Equals(default(KeyValuePair<int, AzureADObject>)))
            {
                this.LabelErrorMessage.Text = String.Format(TextErrorNewMetadataAlreadyUsed, New_DdlPermissionMetadata.SelectedValue);
                ShowNewItemForm = true;
                BuildAttributesListTable(false);
                return;
            }

            GraphProperty prop;
            bool convertSuccess = Enum.TryParse<GraphProperty>(New_DdlGraphProperty.SelectedValue, out prop);
            if (!convertSuccess || prop == GraphProperty.None)
            {
                this.LabelErrorMessage.Text = TextErrorFieldsMissing;
                ShowNewItemForm = true;
                BuildAttributesListTable(false);
                return;
            }
            azureObject.GraphProperty = prop;
            convertSuccess = Enum.TryParse<GraphProperty>(New_DdlGraphPropertyToDisplay.SelectedValue, out prop);
            azureObject.GraphPropertyToDisplay = convertSuccess ? prop : GraphProperty.None;
            azureObject.ClaimEntityType = SPClaimEntityTypes.User;
            azureObject.EntityDataKey = New_DdlPermissionMetadata.SelectedValue;

            PersistedObject.AzureADObjects.Add(azureObject);
            UpdatePersistedObject();
            BuildAttributesListTable(false);
        }
    }
}
