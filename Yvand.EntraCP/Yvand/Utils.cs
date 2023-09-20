using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Yvand.Config;

namespace Yvand
{
    public static class Utils
    {
        /// <summary>
        /// Gets the first SharePoint TrustedLoginProvider that has its property ClaimProviderName equals to <paramref name="claimProviderName"/>
        /// LIMITATION: The same claims provider (uniquely identified by its name) cannot be associated to multiple TrustedLoginProvider because at runtime there is no way to determine what TrustedLoginProvider is currently calling
        /// </summary>
        /// <param name="claimProviderName"></param>
        /// <returns></returns>
        public static SPTrustedLoginProvider GetSPTrustAssociatedWithClaimsProvider(string claimProviderName)
        {
            if (String.IsNullOrWhiteSpace(claimProviderName))
            {
                return null;
            }

            var lp = SPSecurityTokenServiceManager.Local.TrustedLoginProviders.Where(x => String.Equals(x.ClaimProviderName, claimProviderName, StringComparison.OrdinalIgnoreCase));

            if (lp != null && lp.Count() == 1)
            {
                return lp.First();
            }

            if (lp != null && lp.Count() > 1)
            {
                Logger.Log($"[{claimProviderName}] Cannot continue because '{claimProviderName}' is set with multiple SPTrustedIdentityTokenIssuer", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
            }
            Logger.Log($"[{claimProviderName}] Cannot continue because '{claimProviderName}' is not set with any SPTrustedIdentityTokenIssuer.\r\nVisit {ClaimsProviderConstants.PUBLICSITEURL} for more information.", TraceSeverity.High, EventSeverity.Warning, TraceCategory.Core);
            return null;
        }

        /// <summary>
        /// Checks if the claims provider <paramref name="claimsProviderName"/> should run in the specified <paramref name="context"/>
        /// </summary>
        /// <param name="context">The URI of the current site, or null</param>
        /// <param name="claimsProviderName">The name of the claims provider</param>
        /// <returns></returns>
        public static bool IsClaimsProviderUsedInCurrentContext(Uri context, string claimsProviderName)
        {
            if (context == null) { return true; }
            var webApp = SPWebApplication.Lookup(context);
            if (webApp == null) { return false; }
            if (webApp.IsAdministrationWebApplication) { return true; }

            // Not central admin web app, enable EntraCP only if current web app uses it
            // It is not possible to exclude zones where EntraCP is not used because:
            // Consider following scenario: default zone is WinClaims, intranet zone is Federated:
            // In intranet zone, when creating permission, EntraCP will be called 2 times. The 2nd time (in FillResolve (SPClaim)), the context will always be the URL of the default zone
            foreach (var zone in Enum.GetValues(typeof(SPUrlZone)))
            {
                SPIisSettings iisSettings = webApp.GetIisSettingsWithFallback((SPUrlZone)zone);
                if (!iisSettings.UseTrustedClaimsAuthenticationProvider)
                {
                    continue;
                }

                // Get the list of authentication providers associated with the zone
                foreach (SPAuthenticationProvider prov in iisSettings.ClaimsAuthenticationProviders)
                {
                    if (prov.GetType() == typeof(SPTrustedAuthenticationProvider))
                    {
                        // Check if the current SPTrustedAuthenticationProvider is associated with the claim provider
                        if (String.Equals(prov.ClaimProviderName, claimsProviderName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        public static IdentityClaimTypeConfig IdentifyIdentityClaimTypeConfigFromClaimTypeConfigCollection(ClaimTypeConfigCollection claimTypeConfigCollection, string identityClaimType)
        {
            ClaimTypeConfig claimTypeConfig = claimTypeConfigCollection.FirstOrDefault(x =>
                String.Equals(x.ClaimType, identityClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                !x.UseMainClaimTypeOfDirectoryObject &&
                x.EntityProperty != DirectoryObjectProperty.NotSet);

            if (claimTypeConfig != null) 
            {
                claimTypeConfig = IdentityClaimTypeConfig.ConvertClaimTypeConfig(claimTypeConfig);
            }
            return (IdentityClaimTypeConfig)claimTypeConfig;
        }

        /// <summary>
        /// Copy in target all the fields of source which have the decoration [Persisted] set on the specified type (fields inherited from parent types are ignored)
        /// </summary>
        /// <param name="T"></param>
        /// <param name="source"></param>
        /// <param name="target"></param>
        /// <returns>The target object with fields decorated with [Persisted] set from the source object</returns>
        public static object CopyPersistedFields(Type T, object source, object target)
        {
            List<FieldInfo> persistedFields = T
            .GetRuntimeFields()
            .Where(field => field.GetCustomAttributes(typeof(PersistedAttribute), inherit: false).Any())
            .ToList();

            foreach(FieldInfo field in persistedFields) 
            {
                field.SetValue(target, field.GetValue(source));
            }
            return target;
        }
    }
}
