using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Utilities;
using System;
using System.Web.UI;
using Yvand.ClaimsProviders.Config;

namespace Yvand.ClaimsProviders.Administration
{
    // Sadly, using a generic class with a UserControl seems not possible: https://stackoverflow.com/questions/74733106/asp-net-webforms-usercontrol-with-generic-type-parameter
    //public abstract class AzureCPUserControl<TConfiguration> : UserControl where TConfiguration : EntityProviderConfiguration
    public abstract class AzureCPUserControl : UserControl
    {
        /// <summary>
        /// This member is used in the markup code and cannot be made as a property
        /// </summary>
        public string ClaimsProviderName;

        /// <summary>
        /// This member is used in the markup code and cannot be made as a property
        /// </summary>
        public string ConfigurationName;

        private Guid _ConfigurationID;
        public string ConfigurationID
        {
            get
            {
                return (this._ConfigurationID == null || this._ConfigurationID == Guid.Empty) ? String.Empty : this._ConfigurationID.ToString();
            }
            set
            {
                this._ConfigurationID = new Guid(value);
            }
        }

        private AADEntityProviderConfig<IAADSettings> _Configuration;
        protected AADEntityProviderConfig<IAADSettings> Configuration
        {
            get
            {
                SPSecurity.RunWithElevatedPrivileges(delegate ()
                {
                    if (_Configuration == null)
                    {
                        _Configuration = (AADEntityProviderConfig<IAADSettings>)AADEntityProviderConfig<IAADSettings>.GetGlobalConfiguration(new Guid(this.ConfigurationID), true);
                    }
                    if (_Configuration == null)
                    {
                        SPContext.Current.Web.AllowUnsafeUpdates = true;
                        _Configuration = (AADEntityProviderConfig<IAADSettings>)AADEntityProviderConfig<IAADSettings>.CreateGlobalConfiguration(new Guid(this.ConfigurationID), this.ConfigurationName, this.ClaimsProviderName, typeof(AADEntityProviderConfig<IAADSettings>));
                        SPContext.Current.Web.AllowUnsafeUpdates = false;
                    }
                    _Configuration.RefreshLocalConfigurationIfNeeded();
                });
                return _Configuration;
            }
        }

        protected ConfigStatus Status;

        protected long ConfigurationVersion
        {
            get
            {
                if (ViewState[ViewStatePersistedObjectVersionKey] == null)
                    ViewState.Add(ViewStatePersistedObjectVersionKey, Configuration.Version);
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
                    return String.Format(TextErrorNoIdentityClaimType, Configuration.SPTrust.DisplayName, Configuration.SPTrust.IdentityClaimTypeInformation.MappedClaimType);
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
                ViewState.Add("PersistedObjectName", ConfigurationName);
                ViewState.Add("PersistedObjectID", ConfigurationID);
            }
            else
            {
                ClaimsProviderName = ViewState["ClaimsProviderName"].ToString();
                ConfigurationName = ViewState["PersistedObjectName"].ToString();
                ConfigurationID = ViewState["PersistedObjectID"].ToString();
            }

            Status = ConfigStatus.AllGood;
            if (String.IsNullOrEmpty(ClaimsProviderName)) { Status |= ConfigStatus.ClaimsProviderNamePropNotSet; }
            if (String.IsNullOrEmpty(ConfigurationName)) { Status |= ConfigStatus.PersistedObjectNamePropNotSet; }
            if (String.IsNullOrEmpty(ConfigurationID)) { Status |= ConfigStatus.PersistedObjectIDPropNotSet; }
            if (Status != ConfigStatus.AllGood)
            {
                Logger.Log($"[{ClaimsProviderName}] {MostImportantError}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                // Should not go further if those requirements are not met
                return Status;
            }

            if (Configuration == null)
            {
                Status |= ConfigStatus.PersistedObjectNotFound;
                return Status;
            }

            if (Configuration.SPTrust == null)
            {
                Status |= ConfigStatus.NoSPTrustAssociation;
                return Status;
            }

            if (Status != ConfigStatus.AllGood)
            {
                Logger.Log($"[{ClaimsProviderName}] {MostImportantError}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
                // Should not go further if those requirements are not met
                return Status;
            }

            if (Configuration.LocalConfiguration.IdentityClaimTypeConfig == null)
            {
                Status |= ConfigStatus.NoIdentityClaimType;
            }
            if (ConfigurationVersion != Configuration.Version)
            {
                Status |= ConfigStatus.PersistedObjectStale;
            }

            if (Status != ConfigStatus.AllGood)
            {
                Logger.Log($"[{ClaimsProviderName}] {MostImportantError}", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Configuration);
            }
            return Status;
        }

        public virtual void CommitChanges()
        {
            Configuration.Update();
            ConfigurationVersion = Configuration.Version;
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
