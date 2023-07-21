using azurecp;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System;
using System.Linq;
using System.Web.UI;
using Yvand.ClaimsProviders;
using Yvand.ClaimsProviders.Configuration;
using Yvand.ClaimsProviders.Configuration.AzureAD;
using static Yvand.ClaimsProviders.ClaimsProviderLogging;

namespace azurecp.ControlTemplates
{
    public abstract class AzureCPUserControl : UserControl
    {
        /// <summary>
        /// This member is used in the markup code and cannot be made as a property
        /// </summary>
        public string ClaimsProviderName;

        /// <summary>
        /// This member is used in the markup code and cannot be made as a property
        /// </summary>
        public string PersistedObjectName;

        private Guid _PersistedObjectID;
        public string PersistedObjectID
        {
            get
            {
                return (this._PersistedObjectID == null || this._PersistedObjectID == Guid.Empty) ? String.Empty : this._PersistedObjectID.ToString();
            }
            set
            {
                this._PersistedObjectID = new Guid(value);
            }
        }

        private AzureADEntityProviderConfiguration _PersistedObject;
        protected AzureADEntityProviderConfiguration PersistedObject
        {
            get
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    if (_PersistedObject == null)
                    {
                        _PersistedObject = AzureADEntityProviderConfiguration.GetConfiguration(PersistedObjectName);
                    }
                    if (_PersistedObject == null)
                    {
                        SPContext.Current.Web.AllowUnsafeUpdates = true;
                        _PersistedObject = AzureADEntityProviderConfiguration.CreateConfiguration(this.PersistedObjectID, this.PersistedObjectName, this.ClaimsProviderName);
                        SPContext.Current.Web.AllowUnsafeUpdates = false;
                    }
                });
                return _PersistedObject as AzureADEntityProviderConfiguration;
            }
            //set { _PersistedObject = value; }
        }

        protected SPTrustedLoginProvider CurrentTrustedLoginProvider;
        protected IdentityClaimTypeConfig IdentityCTConfig;
        protected ConfigStatus Status;

        protected long PersistedObjectVersion
        {
            get
            {
                if (ViewState[ViewStatePersistedObjectVersionKey] == null)
                    ViewState.Add(ViewStatePersistedObjectVersionKey, PersistedObject.Version);
                return (long)ViewState[ViewStatePersistedObjectVersionKey];
            }
            set { ViewState[ViewStatePersistedObjectVersionKey] = value; }
        }

        protected string MostImportantError
        {
            get
            {
                if (Status == ConfigStatus.AllGood)
                {
                    return String.Empty;
                }

                if ((Status & ConfigStatus.NoSPTrustAssociation) == ConfigStatus.NoSPTrustAssociation)
                {
                    return String.Format(TextErrorNoSPTrustAssociation, SPEncode.HtmlEncode(ClaimsProviderName));
                }

                if ((Status & ConfigStatus.PersistedObjectNotFound) == ConfigStatus.PersistedObjectNotFound)
                {
                    return TextErrorPersistedObjectNotFound;
                }

                if ((Status & ConfigStatus.NoIdentityClaimType) == ConfigStatus.NoIdentityClaimType)
                {
                    return String.Format(TextErrorNoIdentityClaimType, CurrentTrustedLoginProvider.DisplayName, CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType);
                }

                if ((Status & ConfigStatus.PersistedObjectStale) == ConfigStatus.PersistedObjectStale)
                {
                    return TextErrorPersistedObjectStale;
                }

                if ((Status & ConfigStatus.ClaimsProviderNamePropNotSet) == ConfigStatus.ClaimsProviderNamePropNotSet)
                {
                    return TextErrorClaimsProviderNameNotSet;
                }

                if ((Status & ConfigStatus.PersistedObjectNamePropNotSet) == ConfigStatus.PersistedObjectNamePropNotSet)
                {
                    return TextErrorPersistedObjectNameNotSet;
                }

                if ((Status & ConfigStatus.PersistedObjectIDPropNotSet) == ConfigStatus.PersistedObjectIDPropNotSet)
                {
                    return TextErrorPersistedObjectIDNotSet;
                }

                return String.Empty;
            }
        }

        protected static readonly string ViewStatePersistedObjectVersionKey = "PersistedObjectVersion";
        protected static readonly string TextErrorPersistedObjectNotFound = "PersistedObject cannot be found.";
        protected static readonly string TextErrorPersistedObjectStale = "Modifications were not applied because the persisted object was modified after this page was loaded. Please refresh the page and try again.";
        protected static readonly string TextErrorNoSPTrustAssociation = "{0} is currently not associated with any TrustedLoginProvider, which is required to create entities.<br/>Visit <a href=\"" + ClaimsProviderConstants.PUBLICSITEURL + "\" target=\"_blank\">AzureCP site</a> for more information.<br/>Refresh this page once '{0}' is associated with a TrustedLoginProvider.";
        protected static readonly string TextErrorNoIdentityClaimType = "The TrustedLoginProvider {0} is set with identity claim type '{1}', but is not set in claim types configuration list.<br/>Please visit claim types configuration page to add it.";
        protected static readonly string TextErrorClaimsProviderNameNotSet = "The attribute 'ClaimsProviderName' must be set in the user control.";
        protected static readonly string TextErrorPersistedObjectNameNotSet = "The attribute 'PersistedObjectName' must be set in the user control.";
        protected static readonly string TextErrorPersistedObjectIDNotSet = "The attribute 'PersistedObjectID' must be set in the user control.";

        /// <summary>
        /// Ensures configuration is valid to proceed
        /// </summary>
        /// <returns></returns>
        public virtual ConfigStatus ValidatePrerequisite()
        {
            if (!this.IsPostBack)
            {
                // DataBind() must be called to bind attributes that are set as "<%# #>"in .aspx
                // But only during initial page load, otherwise it would reset bindings in other controls like SPGridView
                DataBind();
                ViewState.Add("ClaimsProviderName", ClaimsProviderName);
                ViewState.Add("PersistedObjectName", PersistedObjectName);
                ViewState.Add("PersistedObjectID", PersistedObjectID);
            }
            else
            {
                ClaimsProviderName = ViewState["ClaimsProviderName"].ToString();
                PersistedObjectName = ViewState["PersistedObjectName"].ToString();
                PersistedObjectID = ViewState["PersistedObjectID"].ToString();
            }

            Status = ConfigStatus.AllGood;
            if (String.IsNullOrEmpty(ClaimsProviderName)) { Status |= ConfigStatus.ClaimsProviderNamePropNotSet; }
            if (String.IsNullOrEmpty(PersistedObjectName)) { Status |= ConfigStatus.PersistedObjectNamePropNotSet; }
            if (String.IsNullOrEmpty(PersistedObjectID)) { Status |= ConfigStatus.PersistedObjectIDPropNotSet; }
            if (Status != ConfigStatus.AllGood)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] {MostImportantError}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                // Should not go further if those requirements are not met
                return Status;
            }

            if (CurrentTrustedLoginProvider == null)
            {
                CurrentTrustedLoginProvider = Utils.GetSPTrustAssociatedWithClaimsProvider(this.ClaimsProviderName);
                if (CurrentTrustedLoginProvider == null)
                {
                    Status |= ConfigStatus.NoSPTrustAssociation;
                    return Status;
                }
            }

            if (PersistedObject == null)
            {
                Status |= ConfigStatus.PersistedObjectNotFound;
            }

            if (Status != ConfigStatus.AllGood)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] {MostImportantError}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                // Should not go further if those requirements are not met
                return Status;
            }

            // AzureADEntityProviderConfiguration.GetConfiguration will call method AzureADEntityProviderConfiguration.CheckAndCleanConfiguration();
            //PersistedObject.CheckAndCleanConfiguration(CurrentTrustedLoginProvider.Name);
            PersistedObject.ClaimTypes.SPTrust = CurrentTrustedLoginProvider;
            if (IdentityCTConfig == null && Status == ConfigStatus.AllGood)
            {
                IdentityCTConfig = PersistedObject.ClaimTypes.FirstOrDefault(x => String.Equals(CurrentTrustedLoginProvider.IdentityClaimTypeInformation.MappedClaimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase) && !x.UseMainClaimTypeOfDirectoryObject) as IdentityClaimTypeConfig;
                if (IdentityCTConfig == null)
                {
                    Status |= ConfigStatus.NoIdentityClaimType;
                }
            }
            if (PersistedObjectVersion != PersistedObject.Version)
            {
                Status |= ConfigStatus.PersistedObjectStale;
            }

            if (Status != ConfigStatus.AllGood)
            {
                ClaimsProviderLogging.Log($"[{ClaimsProviderName}] {MostImportantError}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
            }
            return Status;
        }

        public virtual void CommitChanges()
        {
            PersistedObject.Update();
            PersistedObjectVersion = PersistedObject.Version;
        }
    }

    [Flags]
    public enum ConfigStatus
    {
        AllGood = 0x0,
        PersistedObjectNotFound = 0x1,
        NoSPTrustAssociation = 0x2,
        NoIdentityClaimType = 0x4,
        PersistedObjectStale = 0x8,
        ClaimsProviderNamePropNotSet = 0x10,
        PersistedObjectNamePropNotSet = 0x20,
        PersistedObjectIDPropNotSet = 0x40
    };
}
