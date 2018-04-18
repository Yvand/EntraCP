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
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                AzureCPLogging svc = AzureCPLogging.Local;
            });
        }        

        public override void FeatureUninstalling(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                AzureCPLogging.Unregister();
            });
        }

        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                base.RemoveClaimProvider(AzureCP._ProviderInternalName);
                AzureCPConfig.DeleteAzureCPConfig(ClaimsProviderConstants.AZURECPCONFIG_NAME);
            });
        }

        public override void FeatureUpgrading(SPFeatureReceiverProperties properties, string upgradeActionName, IDictionary<string, string> parameters)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                AzureCPLogging svc = AzureCPLogging.Local;
            });
        }
    }
}
