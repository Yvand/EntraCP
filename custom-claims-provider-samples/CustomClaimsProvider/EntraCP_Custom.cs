using System.Collections.Generic;
using Yvand.EntraClaimsProvider;
using Yvand.EntraClaimsProvider.Configuration;

namespace CustomClaimsProvider
{
    public class EntraCP_Custom : EntraCP
    {
        /// <summary>
        /// Sets the name of the claims provider, also set in (Get-SPTrustedIdentityTokenIssuer).ClaimProviderName property
        /// </summary>
        public new const string ClaimsProviderName = "EntraCP_Custom";
        
        /// <summary>
        /// Do not remove or change this property
        /// </summary>
        public override string Name => ClaimsProviderName;

        public EntraCP_Custom(string displayName) : base(displayName)
        {
        }

        public override IEntraIDProviderSettings GetSettings()
        {
            ClaimsProviderSettings settings = ClaimsProviderSettings.GetDefaultSettings(ClaimsProviderName);
            EntraIDTenant tenant = new EntraIDTenant
            {
                AzureCloud = AzureCloudName.AzureGlobal,
                Name = "TENANTNAME.onmicrosoft.com",
                ClientId = "CLIENTID",
                ClientSecret = "CLIENTSECRET",
            };
            settings.EntraIDTenants = new List<EntraIDTenant>() { tenant };
            return settings;
        }
    }
}
