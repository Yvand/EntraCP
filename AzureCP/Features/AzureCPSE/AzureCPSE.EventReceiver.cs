using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.Runtime.InteropServices;
using Yvand.ClaimsProviders.Configuration;

namespace Yvand.ClaimsProviders
{
    /// <summary>
    /// This class handles events raised during feature activation, deactivation, installation, uninstallation, and upgrade.
    /// </summary>
    /// <remarks>
    /// The GUID attached to this class may be used during packaging and should not be modified.
    /// </remarks>
    [Guid("39c10d12-2c7f-4148-bd81-2283a5ce4a27")]
    public class AzureCPEventReceiver : SPClaimProviderFeatureReceiver
    {
        public override string ClaimProviderAssembly => typeof(AzureCPSE).Assembly.FullName;

        public override string ClaimProviderDescription => AzureCPSE.ClaimsProviderName;

        public override string ClaimProviderDisplayName => AzureCPSE.ClaimsProviderName;

        public override string ClaimProviderType => typeof(AzureCPSE).FullName;

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            ExecBaseFeatureActivated(properties);
        }

        private void ExecBaseFeatureActivated(Microsoft.SharePoint.SPFeatureReceiverProperties properties)
        {
            // Wrapper function for base FeatureActivated. 
            // Used because base keywork can lead to unverifiable code inside lambda expression
            base.FeatureActivated(properties);
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                try
                {
                    ClaimsProviderLogging svc = ClaimsProviderLogging.Local;
                    ClaimsProviderLogging.Log($"[{AzureCPSE.ClaimsProviderName}] Activating farm-scoped feature for claims provider \"{AzureCPSE.ClaimsProviderName}\"", TraceSeverity.High, EventSeverity.Information, ClaimsProviderLogging.TraceCategory.Configuration);
                    //AzureCPConfig existingConfig = AzureCPConfig.GetConfiguration(ClaimsProviderConstants.CONFIG_NAME);
                    //if (existingConfig == null)
                    //{
                    //    AzureCPConfig.CreateDefaultConfiguration();
                    //}
                    //else
                    //{
                    //    ClaimsProviderLogging.Log($"[{AzureCP._ProviderInternalName}] Use configuration \"{ClaimsProviderConstants.CONFIG_NAME}\" found in the configuration database", TraceSeverity.High, EventSeverity.Information, ClaimsProviderLogging.TraceCategory.Configuration);
                    //}
                }
                catch (Exception ex)
                {
                    ClaimsProviderLogging.LogException(AzureCPSE.ClaimsProviderName, $"activating farm-scoped feature for claims provider \"{AzureCPSE.ClaimsProviderName}\"", ClaimsProviderLogging.TraceCategory.Configuration, ex);
                }
            });
        }

        public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                try
                {
                    ClaimsProviderLogging.Log($"[{AzureCPSE.ClaimsProviderName}] Uninstalling farm-scoped feature for claims provider \"{AzureCPSE.ClaimsProviderName}\": Deleting configuration from the farm", TraceSeverity.High, EventSeverity.Information, ClaimsProviderLogging.TraceCategory.Configuration);
                    //AzureCPConfig.DeleteConfiguration(ClaimsProviderConstants.CONFIG_NAME);
                    ClaimsProviderLogging.Unregister();
                }
                catch (Exception ex)
                {
                    ClaimsProviderLogging.LogException(AzureCPSE.ClaimsProviderName, $"deactivating farm-scoped feature for claims provider \"{AzureCPSE.ClaimsProviderName}\"", ClaimsProviderLogging.TraceCategory.Configuration, ex);
                }
            });
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                try
                {
                    ClaimsProviderLogging.Log($"[{AzureCPSE.ClaimsProviderName}] Deactivating farm-scoped feature for claims provider \"{AzureCPSE.ClaimsProviderName}\": Removing claims provider from the farm (but not its configuration)", TraceSeverity.High, EventSeverity.Information, ClaimsProviderLogging.TraceCategory.Configuration);
                    base.RemoveClaimProvider(AzureCPSE.ClaimsProviderName);
                }
                catch (Exception ex)
                {
                    ClaimsProviderLogging.LogException(AzureCPSE.ClaimsProviderName, $"deactivating farm-scoped feature for claims provider \"{AzureCPSE.ClaimsProviderName}\"", ClaimsProviderLogging.TraceCategory.Configuration, ex);
                }
            });
        }

        /// <summary>
        /// Upgrade must be explicitely triggered as documented in https://www.sharepointnutsandbolts.com/2010/06/feature-upgrade-part-1-fundamentals.html
        /// In PowerShell: 
        /// $feature = [Microsoft.SharePoint.Administration.SPWebService]::AdministrationService.Features["d1817470-ca9f-4b0c-83c5-ea61f9b0660d"]
        /// $feature.Upgrade($false)
        /// Since it's not automatic, this mechanism won't be used at all
        /// </summary>
        /// <param name="properties"></param>
        /// <param name="upgradeActionName"></param>
        /// <param name="parameters"></param>
        //public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, IDictionary<string, string> parameters)
        //{
        //}
    }
}
