using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Yvand.ClaimsProviders.Config;
using static Yvand.ClaimsProviders.Logger;

namespace Yvand.ClaimsProviders
{
    public static class Utils
    {
        /// <summary>
        /// Get the first TrustedLoginProvider associated with current claim provider
        /// LIMITATION: The same claims provider (uniquely identified by its name) cannot be associated to multiple TrustedLoginProvider because at runtime there is no way to determine what TrustedLoginProvider is currently calling
        /// </summary>
        /// <param name="claimProviderName"></param>
        /// <returns></returns>
        public static SPTrustedLoginProvider GetSPTrustAssociatedWithClaimsProvider(string claimProviderName)
        {
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
        /// Check if AzureCP should process input (and show results) based on current URL (context)
        /// </summary>
        /// <param name="context">The context, as a URI</param>
        /// <returns></returns>
        public static bool ShouldRun(Uri context, string claimProviderName)
        {
            if (context == null) { return true; }
            var webApp = SPWebApplication.Lookup(context);
            if (webApp == null) { return false; }
            if (webApp.IsAdministrationWebApplication) { return true; }

            // Not central admin web app, enable AzureCP only if current web app uses it
            // It is not possible to exclude zones where AzureCP is not used because:
            // Consider following scenario: default zone is WinClaims, intranet zone is Federated:
            // In intranet zone, when creating permission, AzureCP will be called 2 times. The 2nd time (in FillResolve (SPClaim)), the context will always be the URL of the default zone
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
                        if (String.Equals(prov.ClaimProviderName, claimProviderName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
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
