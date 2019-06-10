using Microsoft.Graph;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace azurecp.ControlTemplates
{
    public partial class ClaimTypesConfigUserControl : AzureCPUserControl
    {
        List<KeyValuePair<int, ClaimTypeConfig>> ClaimsMapping;
        protected bool ShowNewItemForm = false;
        public bool HideAllContent = false;
        public string TrustName = String.Empty; // This must be a field to be accessible from marup code, it cannot be a property

        string TextErrorFieldsMissing = "Some mandatory fields are missing.";
        string TextErrorUpdateEmptyClaimType = "Claim type must be set.";

        string HtmlCellClaimType = "<span name=\"span_claimtype_{1}\" id=\"span_claimtype_{1}\">{0}</span><input name=\"input_claimtype_{1}\" id=\"input_claimtype_{1}\" style=\"display: none; width: 90%;\" value=\"{0}\"></input>";
        string HtmlCellGraphProperty = "<span name=\"span_graphproperty_{1}\" id=\"span_graphproperty_{1}\">{0}</span><select name=\"list_graphproperty_{1}\" id=\"list_graphproperty_{1}\" style=\"display:none;\" value=\"{0}\">{2}</select>";
        string HtmlCellGraphPropertyToDisplay = "<span name=\"span_GraphPropertyToDisplay_{1}\" id=\"span_GraphPropertyToDisplay_{1}\">{0}</span><select name=\"list_GraphPropertyToDisplay_{1}\" id=\"list_GraphPropertyToDisplay_{1}\" style=\"display:none;\" value=\"{0}\">{2}</select>";
        string HtmlCellMetadata = "<span name=\"span_Metadata_{1}\" id=\"span_Metadata_{1}\">{0}</span><select name=\"list_Metadata_{1}\" id=\"list_Metadata_{1}\" style=\"display:none;\">{2}</select>";
        string HtmlCellPrefixToBypassLookup = "<span name=\"span_PrefixToBypassLookup_{1}\" id=\"span_PrefixToBypassLookup_{1}\">{0}</span><input name=\"input_PrefixToBypassLookup_{1}\" id=\"input_PrefixToBypassLookup_{1}\" style=\"display:none;\" value=\"{0}\"></input>";
        string HtmlCellDirectoryObjectType = "<span name=\"span_ClaimEntityType_{1}\" id=\"span_ClaimEntityType_{1}\">{0}</span><select name=\"list_ClaimEntityType_{1}\" id=\"list_ClaimEntityType_{1}\" style=\"display:none;\">{2}</select>";

        string HtmlEditLink = "<a name=\"editLink_{0}\" id=\"editLink_{0}\" href=\"javascript:Azurecp.ClaimsTablePage.EditItem('{0}')\">Edit</a>";
        string HtmlCancelEditLink = "<a name=\"cancelLink_{0}\" id=\"cancelLink_{0}\" href=\"javascript:Azurecp.ClaimsTablePage.CancelEditItem('{0}')\" style=\"display:none;\">Cancel</a>";

        protected void Page_Load(object sender, EventArgs e)
        {
            Initialize();
        }

        /// <summary>
        /// Initialize controls as needed if prerequisites are ok, otherwise deactivate controls and show error message
        /// </summary>
        protected void Initialize()
        {
            ConfigStatus status = ValidatePrerequisite();
            if (status != ConfigStatus.AllGood && status != ConfigStatus.NoIdentityClaimType)
            {
                this.LabelErrorMessage.Text = base.MostImportantError;
                this.HideAllContent = true;
                this.BtnCreateNewItem.Visible = false;
                return;
            }

            TrustName = CurrentTrustedLoginProvider.Name;
            if (!this.IsPostBack)
            {
                // NEW ITEM FORM
                // Populate LDAPObjectType DDL
                foreach (var value in Enum.GetValues(typeof(DirectoryObjectType)))
                {
                    DdlNewDirectoryObjectType.Items.Add(value.ToString());
                }

                // Populate picker entity metadata DDL
                DdlNewEntityMetadata.Items.Add(String.Empty);
                foreach (object field in typeof(PeopleEditorEntityDataKeys).GetFields())
                {
                    DdlNewEntityMetadata.Items.Add(((System.Reflection.FieldInfo)field).Name);
                }

                DdlNewGraphProperty.Items.Add(String.Empty);
                DdlNewGraphPropertyToDisplay.Items.Add(String.Empty);
                foreach (object field in typeof(AzureADObjectProperty).GetFields())
                {
                    string prop = ((System.Reflection.FieldInfo)field).Name;
                    if (AzureCP.GetPropertyValue(new User(), prop) == null) continue;
                    //if (AzureCP.GetGraphPropertyValue(new Group(), prop) == null) continue;
                    //if (AzureCP.GetGraphPropertyValue(new Role(), prop) == null) continue;

                    DdlNewGraphProperty.Items.Add(prop);
                    DdlNewGraphPropertyToDisplay.Items.Add(prop);
                }
            }
            BuildAttributesListTable(this.IsPostBack);
        }

        /// <summary>
        /// Build table that shows mapping between claim types and Azure AD objects
        /// </summary>
        /// <param name="pendingUpdate">true if there is a post back, which means an update is being made</param>
        private void BuildAttributesListTable(bool pendingUpdate)
        {
            // Copy claims list in a key value pair so that each item has a unique ID that can be used later for update/delete operations
            ClaimsMapping = new List<KeyValuePair<int, ClaimTypeConfig>>();
            int i = 0;
            foreach (ClaimTypeConfig attr in this.PersistedObject.ClaimTypes)
            {
                ClaimsMapping.Add(new KeyValuePair<int, ClaimTypeConfig>(i++, attr));
            }

            bool identityClaimPresent = false;

            TblClaimsMapping.Rows.Clear();

            // FIRST ROW HEADERS
            TableRow tr = new TableRow();
            TableHeaderCell th;
            th = GetTableHeaderCell("Actions");
            th.RowSpan = 2;
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Claim type");
            th.RowSpan = 2;
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Azure AD object details");
            th.ColumnSpan = 3;
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Optional settings");
            th.ColumnSpan = 2;
            tr.Cells.Add(th);
            this.TblClaimsMapping.Rows.Add(tr);

            // SECONDE ROW HEADERS
            tr = new TableRow();
            th = new TableHeaderCell();
            th = GetTableHeaderCell("Object type");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Property to query");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("Property to display");
            tr.Cells.Add(th);
            th = GetTableHeaderCell("<a href='http://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.webcontrols.peopleeditorentitydatakeys_members(v=office.15).aspx' target='_blank'>PickerEntity metadata</a>");
            tr.Cells.Add(th);

            th = GetTableHeaderCell("Prefix to bypass lookup");
            tr.Cells.Add(th);
            this.TblClaimsMapping.Rows.Add(tr);

            foreach (var attr in this.ClaimsMapping)
            {
                tr = new TableRow();
                bool allowEditItem = String.IsNullOrEmpty(attr.Value.ClaimType) ? false : true;

                // ACTIONS
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
                    // CLAIM TYPE
                    string html;
                    TableCell c = null;
                    if (!String.IsNullOrEmpty(attr.Value.ClaimType))
                    {
                        html = String.Format(HtmlCellClaimType, attr.Value.ClaimType, attr.Key);
                        c = GetTableCell(html);
                        allowEditItem = true;
                        if (String.Equals(CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType, attr.Value.ClaimType, StringComparison.InvariantCultureIgnoreCase) && !attr.Value.UseMainClaimTypeOfDirectoryObject)
                        {
                            tr.CssClass = "azurecp-rowidentityclaim";
                            identityClaimPresent = true;
                        }
                        else if (CurrentTrustedLoginProvider.ClaimTypeInformation.FirstOrDefault(x => String.Equals(x.MappedClaimType, attr.Value.ClaimType, StringComparison.InvariantCultureIgnoreCase)) == null)
                        {
                            tr.CssClass = "azurecp-rowClaimTypeNotUsedInTrust";
                        }
                        else if (attr.Value.EntityType == DirectoryObjectType.Group)
                        {
                            tr.CssClass = "azurecp-rowMainGroupClaimType";
                        }
                    }
                    else
                    {
                        if (!attr.Value.UseMainClaimTypeOfDirectoryObject)
                        {
                            c = GetTableCell("Azure AD object property linked to a PickerEntity metadata");
                        }
                        else
                        {
                            c = GetTableCell($"Azure AD property linked to the main mapping for object {attr.Value.EntityType}");
                            if (attr.Value.EntityType == DirectoryObjectType.User)
                            {
                                tr.CssClass = "azurecp-rowUserProperty";
                            }
                            else
                            {
                                tr.CssClass = "azurecp-rowGroupProperty";
                            }
                        }
                    }
                    tr.Cells.Add(c);

                    // DIRECTORY OBJECT SETTINGS
                    string htmlCellGraphProperty;
                    string htmlCellGraphPropertyToDisplay;
                    string htmlCellDirectoryObjectType;
                    BuildGraphPropertyDDLs(attr, out htmlCellGraphProperty, out htmlCellGraphPropertyToDisplay, out htmlCellDirectoryObjectType);

                    tr.Cells.Add(GetTableCell(htmlCellDirectoryObjectType));

                    html = htmlCellGraphProperty;
                    tr.Cells.Add(GetTableCell(html));

                    html = htmlCellGraphPropertyToDisplay;
                    tr.Cells.Add(GetTableCell(html));

                    // OPTIONAL SETTINGS
                    MemberInfo[] members;
                    members = typeof(PeopleEditorEntityDataKeys).GetFields(BindingFlags.Static | BindingFlags.Public);
                    html = BuildDDLFromTypeMembers(HtmlCellMetadata, attr, "EntityDataKey", members, true);
                    tr.Cells.Add(GetTableCell(html));

                    html = String.Format(HtmlCellPrefixToBypassLookup, attr.Value.PrefixToBypassLookup, attr.Key);
                    tr.Cells.Add(GetTableCell(html));
                }
                TblClaimsMapping.Rows.Add(tr);
            }

            if (!identityClaimPresent && !pendingUpdate)
            {
                LabelErrorMessage.Text = String.Format(TextErrorNoIdentityClaimType, CurrentTrustedLoginProvider.DisplayName, CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType);
            }
        }

        private string BuildDDLFromTypeMembers(string htmlCell, KeyValuePair<int, ClaimTypeConfig> attr, string propertyToCheck, MemberInfo[] members, bool addEmptyChoice)
        {
            string option = "<option value=\"{0}\" {1}>{2}</option>";
            string selected = String.Empty;
            bool metadataFound = false;
            StringBuilder options = new StringBuilder();

            // GetValue returns null if object doesn't have a value on this property, using "as string" avoids to throw a NullReference in this case.
            string attrValue = typeof(ClaimTypeConfig).GetProperty(propertyToCheck).GetValue(attr.Value) as string;

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

        private void BuildGraphPropertyDDLs(KeyValuePair<int, ClaimTypeConfig> azureObject, out string htmlCellGraphProperty, out string htmlCellGraphPropertyToDisplay, out string htmlCellDirectoryObjectType)
        {
            string option = "<option value=\"{0}\" {1}>{2}</option>";
            string graphPropertySelected = String.Empty;
            string graphPropertyToDisplaySelected = String.Empty;
            StringBuilder graphPropertyOptions = new StringBuilder();
            StringBuilder graphPropertyToDisplayOptions = new StringBuilder();
            StringBuilder directoryObjectTypeOptions = new StringBuilder();
            bool graphPropertyToDisplayFound = false;

            // Build EntityType list
            string selectedText = azureObject.Value.EntityType == DirectoryObjectType.User ? "selected" : String.Empty;
            directoryObjectTypeOptions.Append(String.Format(option, DirectoryObjectType.User.ToString(), selectedText, DirectoryObjectType.User.ToString()));
            selectedText = azureObject.Value.EntityType == DirectoryObjectType.Group ? "selected" : String.Empty;
            directoryObjectTypeOptions.Append(String.Format(option, DirectoryObjectType.Group.ToString(), selectedText, DirectoryObjectType.Group.ToString()));

            // Build DirectoryObjectProperty and DirectoryObjectPropertyToShowAsDisplayText lists
            foreach (AzureADObjectProperty prop in Enum.GetValues(typeof(AzureADObjectProperty)))
            {
                // Ensure property exists for the current object type
                if (azureObject.Value.EntityType == DirectoryObjectType.User)
                {
                    if (AzureCP.GetPropertyValue(new User(), prop.ToString()) == null) continue;
                }
                else
                {
                    if (AzureCP.GetPropertyValue(new Group(), prop.ToString()) == null) continue;
                }

                graphPropertySelected = azureObject.Value.DirectoryObjectProperty == prop ? "selected" : String.Empty;

                if (azureObject.Value.DirectoryObjectPropertyToShowAsDisplayText == prop)
                {
                    graphPropertyToDisplaySelected = "selected";
                    graphPropertyToDisplayFound = true;
                }
                else graphPropertyToDisplaySelected = String.Empty;

                graphPropertyOptions.Append(String.Format(option, prop.ToString(), graphPropertySelected, prop.ToString()));
                graphPropertyToDisplayOptions.Append(String.Format(option, prop.ToString(), graphPropertyToDisplaySelected, prop.ToString()));
            }

            // Insert at 1st position AzureADObjectProperty.NotSet in GraphPropertyToDisplay DDL and select it if needed
            string selectNotSet = graphPropertyToDisplayFound ? String.Empty : "selected";
            graphPropertyToDisplayOptions = graphPropertyToDisplayOptions.Insert(0, String.Format(option, AzureADObjectProperty.NotSet, selectNotSet, AzureADObjectProperty.NotSet));

            htmlCellGraphProperty = String.Format(HtmlCellGraphProperty, azureObject.Value.DirectoryObjectProperty, azureObject.Key, graphPropertyOptions.ToString());
            string graphPropertyToDisplaySpanDisplay = azureObject.Value.DirectoryObjectPropertyToShowAsDisplayText == AzureADObjectProperty.NotSet ? String.Empty : azureObject.Value.DirectoryObjectPropertyToShowAsDisplayText.ToString();
            htmlCellGraphPropertyToDisplay = String.Format(HtmlCellGraphPropertyToDisplay, graphPropertyToDisplaySpanDisplay, azureObject.Key, graphPropertyToDisplayOptions.ToString());
            htmlCellDirectoryObjectType = String.Format(HtmlCellDirectoryObjectType, azureObject.Value.EntityType, azureObject.Key, directoryObjectTypeOptions.ToString());
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

        void LnkDeleteItem_Command(object sender, CommandEventArgs e)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood && Status != ConfigStatus.NoIdentityClaimType) return;

            string itemId = e.CommandArgument.ToString();
            ClaimTypeConfig ctConfig = ClaimsMapping.Find(x => x.Key == Convert.ToInt32(itemId)).Value;
            PersistedObject.ClaimTypes.Remove(ctConfig);
            CommitChanges();
            this.BuildAttributesListTable(false);
        }

        void LnkUpdateItem_Command(object sender, CommandEventArgs e)
        {
            if (ValidatePrerequisite() != ConfigStatus.AllGood && Status != ConfigStatus.NoIdentityClaimType) return;

            string itemId = e.CommandArgument.ToString();
            ClaimTypeConfig existingCTConfig = ClaimsMapping.Find(x => x.Key == Convert.ToInt32(itemId)).Value;

            // Get new values
            NameValueCollection formData = Request.Form;
            string newClaimType = formData["input_claimtype_" + itemId];
            string newDirectoryObjectType = formData["list_ClaimEntityType_" + itemId];
            Enum.TryParse(newDirectoryObjectType, out DirectoryObjectType typeSelected);

            if (String.IsNullOrEmpty(newClaimType))
            {
                this.LabelErrorMessage.Text = TextErrorUpdateEmptyClaimType;
                BuildAttributesListTable(false);
                return;
            }

            ClaimTypeConfig newCTConfig = existingCTConfig.CopyConfiguration();
            newCTConfig.ClaimType = newClaimType;
            newCTConfig.EntityType = typeSelected;
            newCTConfig.PrefixToBypassLookup = formData["input_PrefixToBypassLookup_" + itemId];
            newCTConfig.EntityDataKey = formData["list_Metadata_" + itemId];

            AzureADObjectProperty prop;
            bool convertSuccess = Enum.TryParse<AzureADObjectProperty>(formData["list_graphproperty_" + itemId], out prop);
            if (convertSuccess) newCTConfig.DirectoryObjectProperty = prop;
            convertSuccess = Enum.TryParse<AzureADObjectProperty>(formData["list_GraphPropertyToDisplay_" + itemId], out prop);
            if (convertSuccess) newCTConfig.DirectoryObjectPropertyToShowAsDisplayText = prop;

            try
            {
                // ClaimTypeConfigCollection.Update() may thrown an exception if new ClaimTypeConfig is not valid for any reason
                PersistedObject.ClaimTypes.Update(existingCTConfig.ClaimType, newCTConfig);
            }
            catch (Exception ex)
            {
                this.LabelErrorMessage.Text = ex.Message;
                BuildAttributesListTable(false);
                return;
            }
            CommitChanges();
            this.BuildAttributesListTable(false);
        }

        protected void BtnReset_Click(object sender, EventArgs e)
        {
            PersistedObject.ResetClaimTypesList();
            PersistedObject.Update();
            Response.Redirect(Request.Url.ToString());
        }

        /// <summary>
        /// Add a new claim type configuration
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void BtnCreateNewItem_Click(object sender, EventArgs e)
        {
            string newClaimType = TxtNewClaimType.Text.Trim();
            AzureADObjectProperty newDirectoryObjectProp;
            Enum.TryParse<AzureADObjectProperty>(DdlNewGraphProperty.SelectedValue, out newDirectoryObjectProp);
            DirectoryObjectType newDirectoryObjectType;
            Enum.TryParse<DirectoryObjectType>(DdlNewDirectoryObjectType.SelectedValue, out newDirectoryObjectType);
            bool useMainClaimTypeOfDirectoryObject = false;

            if (RdbNewItemClassicClaimType.Checked)
            {
                if (String.IsNullOrEmpty(TxtNewClaimType.Text))
                {
                    this.LabelErrorMessage.Text = TextErrorFieldsMissing;
                    ShowNewItemForm = true;
                    BuildAttributesListTable(false);
                    return;
                }
            }
            else if (RdbNewItemPermissionMetadata.Checked)
            {
                if (String.IsNullOrEmpty(DdlNewEntityMetadata.SelectedValue))
                {
                    this.LabelErrorMessage.Text = TextErrorFieldsMissing;
                    ShowNewItemForm = true;
                    BuildAttributesListTable(false);
                    return;
                }
                newClaimType = String.Empty;
            }
            else
            {
                useMainClaimTypeOfDirectoryObject = true;
                newClaimType = String.Empty;
            }

            ClaimTypeConfig newCTConfig = new ClaimTypeConfig();
            newCTConfig.ClaimType = newClaimType;
            newCTConfig.DirectoryObjectProperty = newDirectoryObjectProp;
            newCTConfig.EntityType = newDirectoryObjectType;
            newCTConfig.UseMainClaimTypeOfDirectoryObject = useMainClaimTypeOfDirectoryObject;
            newCTConfig.EntityDataKey = DdlNewEntityMetadata.SelectedValue;
            bool convertSuccess = Enum.TryParse<AzureADObjectProperty>(DdlNewGraphPropertyToDisplay.SelectedValue, out newDirectoryObjectProp);
            newCTConfig.DirectoryObjectPropertyToShowAsDisplayText = convertSuccess ? newDirectoryObjectProp : AzureADObjectProperty.NotSet;

            try
            {
                // ClaimTypeConfigCollection.Add() may thrown an exception if new ClaimTypeConfig is not valid for any reason
                PersistedObject.ClaimTypes.Add(newCTConfig);
            }
            catch (Exception ex)
            {
                this.LabelErrorMessage.Text = ex.Message;
                ShowNewItemForm = true;
                BuildAttributesListTable(false);
                return;
            }

            // Update configuration and rebuild table with new configuration
            CommitChanges();
            BuildAttributesListTable(false);
        }
    }
}
