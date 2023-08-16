using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.Runtime.InteropServices;
using Yvand.ClaimsProviders.Config;

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
        public override string ClaimProviderAssembly => typeof(AzureCP).Assembly.FullName;

        public override string ClaimProviderDescription => AzureCP.ClaimsProviderName;

        public override string ClaimProviderDisplayName => AzureCP.ClaimsProviderName;

        public override string ClaimProviderType => typeof(AzureCP).FullName;

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            ExecBaseFeatureActivated(properties);
        }

        private void ExecBaseFeatureActivated(Microsoft.SharePoint.SPFeatureReceiverProperties properties)
        {
            // Wrapper function for base FeatureActivated. 
            // Used because base keywork can lead to unverifiable code inside lambda expression
            base.FeatureActivated(properties);
            SPSecurity.RunWithElevatedPrivileges((SPSecurity.CodeToRunElevated)delegate ()
            {
                try
                {
                    Logger svc = Logger.Local;
                    Logger.Log($"[{AzureCP.ClaimsProviderName}] Activating farm-scoped feature for claims provider \"{AzureCP.ClaimsProviderName}\"", TraceSeverity.High, EventSeverity.Information, TraceCategory.Configuration);
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
                    Logger.LogException((string)AzureCP.ClaimsProviderName, $"activating farm-scoped feature for claims provider \"{AzureCP.ClaimsProviderName}\"", TraceCategory.Configuration, ex);
                }
            });
        }

        public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges((SPSecurity.CodeToRunElevated)delegate ()
            {
                try
                {
                    Logger.Log($"[{AzureCP.ClaimsProviderName}] Uninstalling farm-scoped feature for claims provider \"{AzureCP.ClaimsProviderName}\": Deleting configuration from the farm", TraceSeverity.High, EventSeverity.Information, TraceCategory.Configuration);
                    //AzureCPConfig.DeleteConfiguration(ClaimsProviderConstants.CONFIG_NAME);
                    Logger.Unregister();
                }
                catch (Exception ex)
                {
                    Logger.LogException((string)AzureCP.ClaimsProviderName, $"deactivating farm-scoped feature for claims provider \"{AzureCP.ClaimsProviderName}\"", TraceCategory.Configuration, ex);
                }
            });
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges((SPSecurity.CodeToRunElevated)delegate ()
            {
                try
                {
                    Logger.Log($"[{AzureCP.ClaimsProviderName}] Deactivating farm-scoped feature for claims provider \"{AzureCP.ClaimsProviderName}\": Removing claims provider from the farm (but not its configuration)", TraceSeverity.High, EventSeverity.Information, TraceCategory.Configuration);
                    base.RemoveClaimProvider((string)AzureCP.ClaimsProviderName);
                }
                catch (Exception ex)
                {
                    Logger.LogException((string)AzureCP.ClaimsProviderName, $"deactivating farm-scoped feature for claims provider \"{AzureCP.ClaimsProviderName}\"", TraceCategory.Configuration, ex);
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
