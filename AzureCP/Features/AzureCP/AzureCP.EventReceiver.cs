using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration.Claims;
using System.Collections.Generic;

namespace azurecp
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

        public override string ClaimProviderDescription => AzureCP._ProviderInternalName;

        public override string ClaimProviderDisplayName => AzureCP._ProviderInternalName;

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
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                ClaimsProviderLogging svc = ClaimsProviderLogging.Local;
            });
        }

        public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                ClaimsProviderLogging.Unregister();
            });
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                base.RemoveClaimProvider(AzureCP._ProviderInternalName);
                AzureCPConfig.DeleteConfiguration(ClaimsProviderConstants.AZURECPCONFIG_NAME);
            });
        }

        public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, IDictionary<string, string> parameters)
        {
            // Upgrade must be explicitely triggered as documented in https://www.sharepointnutsandbolts.com/2010/06/feature-upgrade-part-1-fundamentals.html
            // In PowerShell: 
            // $feature = [Microsoft.SharePoint.Administration.SPWebService]::AdministrationService.Features["d1817470-ca9f-4b0c-83c5-ea61f9b0660d"]
            // $feature.Upgrade($false)
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                ClaimsProviderLogging svc = ClaimsProviderLogging.Local;
                var spTrust = AzureCP.GetSPTrustAssociatedWithCP(AzureCP._ProviderInternalName);
                string spTrustName = spTrust == null ? String.Empty : spTrust.Name;
                // AzureCPConfig.GetConfiguration will call method AzureCPConfig.CheckAndCleanConfiguration();
                AzureCPConfig config = AzureCPConfig.GetConfiguration(ClaimsProviderConstants.AZURECPCONFIG_NAME, spTrustName);
                //if (config != null)
                //{
                //    config.CheckAndCleanConfiguration(spTrustName);
                //}
            });
        }
    }
}
