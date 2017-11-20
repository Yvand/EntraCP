using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.ActiveDirectory.GraphClient.Extensions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using WIF = System.Security.Claims;

/*
 * DO NOT directly edit AzureCP class. It is designed to be inherited to customize it as desired.
 * Please download "AzureCP for Developers.zip" on http://azurecp.codeplex.com to find examples and guidance.
 * */

namespace azurecp
{
    /// <summary>
    /// Provides search and resolution against Azure Active Directory
    /// Visit http://azurecp.codeplex.com/ for documentation and updates.
    /// Please report any bug to http://azurecp.codeplex.com/.
    /// Author: Yvan Duhamel
    /// Copyright (c) 2014, Yvan Duhamel. All rights reserved.
    /// </summary>
    public class AzureCP : SPClaimProvider
    {
        public const string _ProviderInternalName = "AzureCP";
        public virtual string ProviderInternalName { get { return "AzureCP"; } }

        private object Sync_Init = new object();
        private ReaderWriterLockSlim Lock_Config = new ReaderWriterLockSlim();
        private long AzureCPConfigVersion = 0;

        /// <summary>
        /// Contains configuration currently used by claims provider
        /// </summary>
        public IAzureCPConfiguration CurrentConfiguration;

        /// <summary>
        /// SPTrust associated with the claims provider
        /// </summary>
        protected SPTrustedLoginProvider SPTrust;

        /// <summary>
        /// object mapped to the identity claim in the SPTrustedIdentityTokenIssuer
        /// </summary>
        AzureADObject IdentityAzureObject;

        /// <summary>
        /// Processed list to use. It is guarranted to never contain an empty ClaimType
        /// </summary>
        public List<AzureADObject> ProcessedAzureObjects;
        public List<AzureADObject> ProcessedAzureObjectsMetadata;
        protected virtual string PickerEntityDisplayText { get { return "({0}) {1}"; } }
        protected virtual string PickerEntityOnMouseOver { get { return "{0}={1}"; } }

        protected string IssuerName
        {
            get
            {
                // The advantage of using the SPTrustedLoginProvider name for the issuer name is that it makes possible and easy to replace current claims provider with another one.
                // The other claims provider would simply have to use SPTrustedLoginProvider name too
                return SPOriginalIssuers.Format(SPOriginalIssuerType.TrustedProvider, SPTrust.Name);
            }
        }

        public AzureCP(string displayName)
            : base(displayName)
        {
            AzureCPLogging.Log(String.Format("[{0}] Constructor called", ProviderInternalName), TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);
        }

        /// <summary>
        /// Initializes claim provider. This method is reserved for internal use and is not intended to be called from external code or changed
        /// </summary>
        public bool Initialize(Uri context, string[] entityTypes)
        {
            // Ensures thread safety to initialize class variables
            lock (Sync_Init)
            {
                // 1ST PART: GET CONFIGURATION OBJECT
                AzureCPConfig globalConfiguration = null;
                bool refreshConfig = false;
                bool success = true;
                bool initializeFromPersistedObject = true;
                try
                {
                    if (SPTrust == null)
                    {
                        SPTrust = GetSPTrustAssociatedWithCP(ProviderInternalName);
                        if (SPTrust == null) return false;
                    }
                    if (!CheckIfShouldProcessInput(context)) return false;

                    // Should not try to get PersistedObject if not OOB AzureCP since with current design it works correctly only for OOB AzureCP
                    if (String.Equals(ProviderInternalName, AzureCP._ProviderInternalName, StringComparison.InvariantCultureIgnoreCase))
                    {
                        globalConfiguration = AzureCPConfig.GetFromConfigDB();
                        if (globalConfiguration == null)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject not found. Visit AzureCP admin pages in central administration to create it.", ProviderInternalName),
                                TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            // Cannot continue since it's not inherited and no persisted object exists
                            success = false;
                        }
                        else if (globalConfiguration.AzureADObjects == null || globalConfiguration.AzureADObjects.Count == 0)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject was found but there are no AzureADObject set. Visit AzureCP admin pages in central administration to create it.", ProviderInternalName),
                                TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            // Cannot continue 
                            success = false;
                        }
                        else if (globalConfiguration.AzureTenants == null || globalConfiguration.AzureTenants.Count == 0)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject was found but there are no Azure tenant set. Visit AzureCP admin pages in central administration to add one.", ProviderInternalName),
                                TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            // Cannot continue 
                            success = false;
                        }
                        else
                        {
                            // Persisted object is found and seems valid
                            AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject found, version: {1}, previous version: {2}", ProviderInternalName, globalConfiguration.Version.ToString(), this.AzureCPConfigVersion.ToString()),
                                TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);
                            if (this.AzureCPConfigVersion != globalConfiguration.Version)
                            {
                                refreshConfig = true;
                                this.AzureCPConfigVersion = globalConfiguration.Version;
                                AzureCPLogging.Log(String.Format("[{0}] AzureCPConfig PersistedObject changed, refreshing configuration", ProviderInternalName),
                                    TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Core);
                            }
                        }
                    }
                    else
                    {
                        // AzureCP class inherited, refresh config
                        // Configuration will be retrieved in SetCustomSettings method
                        initializeFromPersistedObject = false;
                        refreshConfig = true;
                        AzureCPLogging.Log(String.Format("[{0}] AzureCP class inherited", ProviderInternalName),
                            TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Core);
                    }
                }
                catch (Exception ex)
                {
                    success = false;
                    AzureCPLogging.LogException(ProviderInternalName, "in Initialize", AzureCPLogging.Categories.Core, ex);
                }
                finally
                { }

                if (!success) return success;
                if (!refreshConfig) return success;

                // 2ND PART: APPLY CONFIGURATION
                // Configuration needs to be refreshed, lock current thread in write mode
                Lock_Config.EnterWriteLock();
                try
                {
                    AzureCPLogging.Log(String.Format("[{0}] Refreshing configuration", ProviderInternalName),
                        TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Core);

                    // Create local persisted object that will never be saved in config DB, it's just a local copy
                    this.CurrentConfiguration = new AzureCPConfig();
                    if (initializeFromPersistedObject)
                    {
                        // All settings come from persisted object
                        this.CurrentConfiguration.AlwaysResolveUserInput = globalConfiguration.AlwaysResolveUserInput;
                        this.CurrentConfiguration.FilterExactMatchOnly = globalConfiguration.FilterExactMatchOnly;
                        this.CurrentConfiguration.AugmentAADRoles = globalConfiguration.AugmentAADRoles;

                        // Retrieve AzureADObjects
                        // A copy of collection AzureADObjects must be created because SetActualAADObjectCollection() may change it and it should be made in a copy totally independant from the persisted object
                        this.CurrentConfiguration.AzureADObjects = new List<AzureADObject>();
                        foreach (AzureADObject currentObject in globalConfiguration.AzureADObjects)
                        {
                            // Create a new AzureADObject
                            this.CurrentConfiguration.AzureADObjects.Add(currentObject.CopyPersistedProperties());
                        }

                        // Retrieve AzureTenants
                        // Create a copy of the collection to work in an copy separated from persisted object
                        this.CurrentConfiguration.AzureTenants = new List<AzureTenant>();
                        foreach (AzureTenant currentObject in globalConfiguration.AzureTenants)
                        {
                            // Create a copy from persisted object
                            this.CurrentConfiguration.AzureTenants.Add(currentObject.CopyPersistedProperties());
                        }
                    }
                    else
                    {
                        // All settings come from overriden SetCustomConfiguration method
                        SetCustomConfiguration(context, entityTypes);

                        // Ensure we get what we expect
                        if (this.CurrentConfiguration.AzureADObjects == null || this.CurrentConfiguration.AzureADObjects.Count == 0)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureADObjects was not set. Override method SetCustomConfiguration to set it.", ProviderInternalName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            return false;
                        }

                        if (this.CurrentConfiguration.AzureTenants == null || this.CurrentConfiguration.AzureTenants.Count == 0)
                        {
                            AzureCPLogging.Log(String.Format("[{0}] AzureTenants was not set. Override method SetCustomConfiguration to set it.", ProviderInternalName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);
                            return false;
                        }
                    }
                    success = this.ProcessAzureADObjectCollection(this.CurrentConfiguration.AzureADObjects);
                }
                catch (Exception ex)
                {
                    success = false;
                    AzureCPLogging.LogException(ProviderInternalName, "in Initialize, while refreshing configuration", AzureCPLogging.Categories.Core, ex);
                }
                finally
                {
                    Lock_Config.ExitWriteLock();
                }
                return success;
            }
        }

        /// <summary>
        /// Initializes claim provider. This method is reserved for internal use and is not intended to be called from external code or changed
        /// </summary>
        /// <param name="AzureADObjects"></param>
        /// <returns></returns>
        private bool ProcessAzureADObjectCollection(List<AzureADObject> AzureADObjectCollection)
        {
            bool success = true;
            try
            {
                bool identityClaimTypeFound = false;
                // Get attributes defined in trust based on their claim type (unique way to map them)
                List<AzureADObject> claimTypesSetInTrust = new List<AzureADObject>();
                // There is a bug in the SharePoint API: SPTrustedLoginProvider.ClaimTypes should retrieve SPTrustedClaimTypeInformation.MappedClaimType, but it returns SPTrustedClaimTypeInformation.InputClaimType instead, so we cannot rely on it
                //foreach (var attr in _AttributesDefinitionList.Where(x => AssociatedSPTrustedLoginProvider.ClaimTypes.Contains(x.claimType)))
                //{
                //    attributesDefinedInTrust.Add(attr);
                //}
                foreach (SPTrustedClaimTypeInformation ClaimTypeInformation in SPTrust.ClaimTypeInformation)
                {
                    // Search if current claim type in trust exists in AzureADObjects
                    // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                    List<AzureADObject> azureObjectColl = AzureADObjectCollection.FindAll(x =>
                        String.Equals(x.ClaimType, ClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        !x.CreateAsIdentityClaim &&
                        x.GraphProperty != GraphProperty.None);
                    AzureADObject azureObject;
                    if (azureObjectColl.Count == 1)
                    {
                        azureObject = azureObjectColl.First();
                        claimTypesSetInTrust.Add(azureObject);

                        if (String.Equals(SPTrust.IdentityClaimTypeInformation.MappedClaimType, azureObject.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                        {
                            // Identity claim type found, set IdentityAzureADObject property
                            identityClaimTypeFound = true;
                            IdentityAzureObject = azureObject;
                        }
                    }
                }

                // Check if identity claim is there. Should always check property SPTrustedClaimTypeInformation.MappedClaimType: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.administration.claims.sptrustedclaimtypeinformation.mappedclaimtype.aspx
                if (!identityClaimTypeFound)
                {
                    AzureCPLogging.Log(String.Format("[{0}] Impossible to continue because identity claim type \"{1}\" set in the SPTrustedIdentityTokenIssuer \"{2}\" is missing in AzureADObjects.", ProviderInternalName, SPTrust.IdentityClaimTypeInformation.MappedClaimType, SPTrust.Name), TraceSeverity.Unexpected, EventSeverity.ErrorCritical, AzureCPLogging.Categories.Core);
                    return false;
                }

                // This check is to find if there is a duplicate of the identity claim type that uses the same GraphProperty
                //AzureADObject objectToDelete = claimTypesSetInTrust.Find(x =>
                //    !String.Equals(x.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                //    !x.CreateAsIdentityClaim &&
                //    x.GraphProperty == GraphProperty.UserPrincipalName);
                //if (objectToDelete != null) claimTypesSetInTrust.Remove(objectToDelete);

                // Check if there are objects that should be always queried (CreateAsIdentityClaim) to add in the list
                List<AzureADObject> additionalObjects = new List<AzureADObject>();
                foreach (AzureADObject attr in AzureADObjectCollection.Where(x => x.CreateAsIdentityClaim))// && !claimTypesSetInTrust.Contains(x, new LDAPPropertiesComparer())))
                {
                    // Check if identity claim type is already using same GraphProperty, and ignore current object if so
                    if (IdentityAzureObject.GraphProperty == attr.GraphProperty) continue;

                    // Normally ClaimType should be null if CreateAsIdentityClaim is set to true, but we check here it and handle this scenario
                    if (!String.IsNullOrEmpty(attr.ClaimType))
                    {
                        if (String.Equals(SPTrust.IdentityClaimTypeInformation.MappedClaimType, attr.ClaimType))
                        {
                            // Not a big deal since it's set with identity claim type, so no inconsistent behavior to expect, just record an information
                            AzureCPLogging.Log(String.Format("[{0}] Object with GraphProperty {1} is set with CreateAsIdentityClaim to true and ClaimType {2}. Remove ClaimType property as it is useless.", ProviderInternalName, attr.GraphProperty, attr.ClaimType), TraceSeverity.Monitorable, EventSeverity.Information, AzureCPLogging.Categories.Core);
                        }
                        else if (claimTypesSetInTrust.Count(x => String.Equals(x.ClaimType, attr.ClaimType)) > 0)
                        {
                            // Same claim type already exists with CreateAsIdentityClaim == false. 
                            // Current object is a bad one and shouldn't be added. Don't add it but continue to build objects list
                            AzureCPLogging.Log(String.Format("[{0}] Claim type {1} is defined twice with CreateAsIdentityClaim set to true and false, which is invalid. Remove entry with CreateAsIdentityClaim set to true.", ProviderInternalName, attr.ClaimType), TraceSeverity.Monitorable, EventSeverity.Information, AzureCPLogging.Categories.Core);
                            continue;
                        }
                    }

                    attr.ClaimType = SPTrust.IdentityClaimTypeInformation.MappedClaimType;    // Give those objects the identity claim type
                    attr.ClaimEntityType = SPClaimEntityTypes.User;
                    attr.GraphPropertyToDisplay = IdentityAzureObject.GraphPropertyToDisplay; // Must be set otherwise display text of permissions will be inconsistent
                    additionalObjects.Add(attr);
                }

                ProcessedAzureObjects = new List<AzureADObject>(claimTypesSetInTrust.Count + additionalObjects.Count);
                ProcessedAzureObjects.AddRange(claimTypesSetInTrust);
                ProcessedAzureObjects.AddRange(additionalObjects);

                // Parse objects to configure some settings
                // An object can have ClaimType set to null if only used to populate metadata of permission created
                foreach (var attr in ProcessedAzureObjects.Where(x => x.ClaimType != null))
                {
                    var trustedClaim = SPTrust.GetClaimTypeInformationFromMappedClaimType(attr.ClaimType);
                    // It should never be null
                    if (trustedClaim == null) continue;
                    attr.ClaimTypeMappingName = trustedClaim.DisplayName;
                }

                // Any metadata for a user with GraphProperty actually set is valid
                this.ProcessedAzureObjectsMetadata = AzureADObjectCollection.FindAll(x =>
                    !String.IsNullOrEmpty(x.EntityDataKey) &&
                    x.GraphProperty != GraphProperty.None &&
                    x.ClaimEntityType == SPClaimEntityTypes.User);
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "while processing AzureADObjects", AzureCPLogging.Categories.Core, ex);
                success = false;
            }
            return success;
        }

        /// <summary>
        /// Override this method to customize configuration of AzureCP
        /// </summary> 
        /// <param name="context">The context, as a URI</param>
        /// <param name="entityTypes">The EntityType entity types set to scope the search to</param>
        protected virtual void SetCustomConfiguration(Uri context, string[] entityTypes)
        {
        }

        /// <summary>
        /// Check if AzureCP should process input (and show results) based on current URL (context)
        /// </summary>
        /// <param name="context">The context, as a URI</param>
        /// <returns></returns>
        protected virtual bool CheckIfShouldProcessInput(Uri context)
        {
            if (context == null) return true;
            var webApp = SPWebApplication.Lookup(context);
            if (webApp == null) return false;
            if (webApp.IsAdministrationWebApplication) return true;

            // Not central admin web app, enable AzureCP only if current web app uses it
            // It is not possible to exclude zones where AzureCP is not used because:
            // Consider following scenario: default zone is NTLM, intranet zone is claims
            // In intranet zone, when creating permission, AzureCP will be called 2 times, but the 2nd time (from FillResolve (SPClaim)) the context will always be the URL of default zone
            foreach (var zone in Enum.GetValues(typeof(SPUrlZone)))
            {
                SPIisSettings iisSettings = webApp.GetIisSettingsWithFallback((SPUrlZone)zone);
                if (!iisSettings.UseTrustedClaimsAuthenticationProvider)
                    continue;

                // Get the list of authentication providers associated with the zone
                foreach (SPAuthenticationProvider prov in iisSettings.ClaimsAuthenticationProviders)
                {
                    if (prov.GetType() == typeof(Microsoft.SharePoint.Administration.SPTrustedAuthenticationProvider))
                    {
                        // Check if the current SPTrustedAuthenticationProvider is associated with the claim provider
                        if (String.Equals(prov.ClaimProviderName, ProviderInternalName, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Get the first TrustedLoginProvider associated with current claim provider
        /// LIMITATION: The same claims provider (uniquely identified by its name) cannot be associated to multiple TrustedLoginProvider because at runtime there is no way to determine what TrustedLoginProvider is currently calling
        /// </summary>
        /// <param name="providerInternalName"></param>
        /// <returns></returns>
        public static SPTrustedLoginProvider GetSPTrustAssociatedWithCP(string providerInternalName)
        {
            var lp = SPSecurityTokenServiceManager.Local.TrustedLoginProviders.Where(x => String.Equals(x.ClaimProviderName, providerInternalName, StringComparison.OrdinalIgnoreCase));

            if (lp != null && lp.Count() == 1)
                return lp.First();

            if (lp != null && lp.Count() > 1)
                AzureCPLogging.Log(String.Format("[{0}] Claims provider {0} is associated to multiple SPTrustedIdentityTokenIssuer, which is not supported because at runtime there is no way to determine what TrustedLoginProvider is currently calling", providerInternalName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Core);

            AzureCPLogging.Log(String.Format("[{0}] Claims provider {0} is not associated with any SPTrustedIdentityTokenIssuer so it cannot create permissions.\r\nVisit http://ldapcp.codeplex.com for installation procedure or set property ClaimProviderName with PowerShell cmdlet Get-SPTrustedIdentityTokenIssuer to create association.", providerInternalName), TraceSeverity.High, EventSeverity.Warning, AzureCPLogging.Categories.Core);
            return null;
        }

        public void BuildFilterAndProcessResultsAsync(string input, List<AzureADObject> azureObjectsToQuery, bool exactSearch, Uri context, string[] entityTypes, ref List<AzurecpResult> results)
        {
            // Create named delegate for users and groups
            Expression<Func<IUser, bool>> userDelegate = null;
            Expression<Func<IGroup, bool>> groupDelegate = null;

            Expression<Func<IUser, bool>> userQuery = null;
            Expression<Func<IGroup, bool>> groupQuery = null;

            foreach (AzureADObject adObject in azureObjectsToQuery)
            {
                if (adObject.ClaimEntityType == SPClaimEntityTypes.User)
                {
                    // Ensure property is of type System.String
                    PropertyInfo pi = typeof(User).GetProperty(adObject.GraphProperty.ToString());
                    if (pi == null) continue;
                    if (pi.PropertyType != typeof(System.String)) continue;

                    if (exactSearch)
                    {
                        if (adObject.GraphProperty == GraphProperty.City) userDelegate = u => u.City.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.Country) userDelegate = u => u.Country.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.Department) userDelegate = u => u.Department.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.DisplayName) userDelegate = u => u.DisplayName.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.GivenName) userDelegate = u => u.GivenName.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.JobTitle) userDelegate = u => u.JobTitle.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.Mail) userDelegate = u => u.Mail.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.MailNickname) userDelegate = u => u.MailNickname.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.Mobile) userDelegate = u => u.Mobile.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.PhysicalDeliveryOfficeName) userDelegate = u => u.PhysicalDeliveryOfficeName.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.PostalCode) userDelegate = u => u.PostalCode.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.PreferredLanguage) userDelegate = u => u.PreferredLanguage.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.State) userDelegate = u => u.State.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.StreetAddress) userDelegate = u => u.StreetAddress.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.Surname) userDelegate = u => u.Surname.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.TelephoneNumber) userDelegate = u => u.TelephoneNumber.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.UserPrincipalName) userDelegate = u => u.UserPrincipalName.Equals(input);
                    }
                    else
                    {
                        if (adObject.GraphProperty == GraphProperty.City) userDelegate = u => u.City.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.Country) userDelegate = u => u.Country.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.Department) userDelegate = u => u.Department.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.DisplayName) userDelegate = u => u.DisplayName.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.GivenName) userDelegate = u => u.GivenName.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.JobTitle) userDelegate = u => u.JobTitle.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.Mail) userDelegate = u => u.Mail.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.MailNickname) userDelegate = u => u.MailNickname.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.Mobile) userDelegate = u => u.Mobile.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.PhysicalDeliveryOfficeName) userDelegate = u => u.PhysicalDeliveryOfficeName.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.PostalCode) userDelegate = u => u.PostalCode.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.PreferredLanguage) userDelegate = u => u.PreferredLanguage.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.State) userDelegate = u => u.State.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.StreetAddress) userDelegate = u => u.StreetAddress.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.Surname) userDelegate = u => u.Surname.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.TelephoneNumber) userDelegate = u => u.TelephoneNumber.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.UserPrincipalName) userDelegate = u => u.UserPrincipalName.StartsWith(input);
                    }
                    userQuery = userQuery.Or(userDelegate);
                }
                else if (adObject.ClaimEntityType == SPClaimEntityTypes.FormsRole)
                {
                    // Ensure property is of type System.String
                    PropertyInfo pi = typeof(Group).GetProperty(adObject.GraphProperty.ToString());
                    if (pi == null) continue;
                    if (pi.PropertyType != typeof(System.String)) continue;

                    if (exactSearch)
                    {
                        if (adObject.GraphProperty == GraphProperty.Description) groupDelegate = g => g.Description.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.DisplayName) groupDelegate = g => g.DisplayName.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.Mail) groupDelegate = g => g.Mail.Equals(input);
                        if (adObject.GraphProperty == GraphProperty.MailNickname) groupDelegate = g => g.MailNickname.Equals(input);
                    }
                    else
                    {
                        if (adObject.GraphProperty == GraphProperty.Description) groupDelegate = g => g.Description.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.DisplayName) groupDelegate = g => g.DisplayName.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.Mail) groupDelegate = g => g.Mail.StartsWith(input);
                        if (adObject.GraphProperty == GraphProperty.MailNickname) groupDelegate = g => g.MailNickname.StartsWith(input);
                    }
                    groupQuery = groupQuery.Or(groupDelegate);
                }
            }

            Task<List<AzurecpResult>> searchResultsTask = this.QueryAzureADCollectionAsync(input, userQuery, groupQuery);
            searchResultsTask.Wait();
            List<AzurecpResult> searchResults = searchResultsTask.Result;
            if (searchResults == null || searchResults.Count == 0) return;

            // If exactSearch is true, we don't care about attributes with CreateAsIdentityClaim = true
            List<AzureADObject> azureObjects;
            if (exactSearch) azureObjects = azureObjectsToQuery.FindAll(x => !x.CreateAsIdentityClaim);
            else azureObjects = azureObjectsToQuery;

            foreach (AzurecpResult searchResult in searchResults)
            {
                Type currentObjectType = null;
                DirectoryObject currentObject = null;
                string claimEntityType = null;
                if (searchResult.DirectoryObjectResult is User)
                {
                    currentObjectType = typeof(User);
                    currentObject = searchResult.DirectoryObjectResult;
                    claimEntityType = SPClaimEntityTypes.User;
                }
                else
                {
                    currentObjectType = typeof(Group);
                    currentObject = searchResult.DirectoryObjectResult;
                    claimEntityType = SPClaimEntityTypes.FormsRole;
                }

                foreach (AzureADObject azureObject in azureObjects.Where(x => x.ClaimEntityType == claimEntityType))
                {
                    // Get value with of current GraphProperty
                    string graphPropertyValue = GetGraphPropertyValue(currentObject, azureObject.GraphProperty.ToString());

                    // Check if property exists (no null) and has a value (not String.Empty)
                    if (String.IsNullOrEmpty(graphPropertyValue)) continue;

                    // Check if current value mathes input, otherwise go to next GraphProperty to check
                    if (exactSearch)
                    {
                        if (!String.Equals(graphPropertyValue, input, StringComparison.InvariantCultureIgnoreCase)) continue;
                    }
                    else
                    {
                        if (!graphPropertyValue.StartsWith(input, StringComparison.InvariantCultureIgnoreCase)) continue;
                    }

                    // Current GraphProperty value matches user input. Add current object in search results if it passes following checks
                    string queryMatchValue = graphPropertyValue;
                    string valueToCheck = queryMatchValue;
                    // Check if current object is not already in the collection
                    AzureADObject objCompare;
                    if (azureObject.CreateAsIdentityClaim)
                    {
                        objCompare = IdentityAzureObject;
                        // Get the value of the GraphProperty linked to IdentityAzureObject
                        valueToCheck = GetGraphPropertyValue(currentObject, IdentityAzureObject.GraphProperty.ToString());
                        if (String.IsNullOrEmpty(valueToCheck)) continue;
                    }
                    else
                    {
                        objCompare = azureObject;
                    }

                    // if claim type, GraphProperty and value are identical, then result is already in collection
                    int numberResultFound = results.FindAll(x =>
                        String.Equals(x.AzureObject.ClaimType, objCompare.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        //x.AzureObject.GraphProperty == objCompare.GraphProperty &&
                        String.Equals(x.PermissionValue, valueToCheck, StringComparison.InvariantCultureIgnoreCase)).Count;
                    if (numberResultFound > 0) continue;

                    // Passed the checks, add it to the searchResults list
                    results.Add(
                        new AzurecpResult
                        {
                            AzureObject = azureObject,
                            //GraphPropertyValue = graphPropertyValue,
                            PermissionValue = valueToCheck,
                            QueryMatchValue = queryMatchValue,
                            DirectoryObjectResult = currentObject,
                            TenantId = searchResult.TenantId,
                        });
                }
            }

            AzureCPLogging.Log(String.Format("[{0}] {1} permission(s) to create after filtering", ProviderInternalName, results.Count), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
            foreach (AzurecpResult result in results)
            {
                PickerEntity pe = CreatePickerEntityHelper(result);
                result.PickerEntity = pe;
            }
        }

        /// <summary>
        /// Returns the graph property value of a GraphObject (User, Group, Role)
        /// </summary>
        /// <param name="src"></param>
        /// <param name="propName"></param>
        /// <returns>Null if property doesn't exist. String.Empty if property exists but has no value. Actual value otherwise</returns>
        public static string GetGraphPropertyValue(object src, string propName)
        {
            System.Reflection.PropertyInfo pi = src.GetType().GetProperty(propName);
            if (pi == null) return null;    // Property doesn't exist
            object propertyValue = pi.GetValue(src, null);
            return propertyValue == null ? String.Empty : propertyValue.ToString();
        }

        /// <summary>
        /// Query Azure tenants
        /// </summary>
        /// <param name="userQuery"></param>
        /// <param name="groupQuery"></param>
        /// <param name="input"></param>
        /// <returns></returns>
        private async Task<List<AzurecpResult>> QueryAzureADCollectionAsync(string input, Expression<Func<IUser, bool>> userQuery, Expression<Func<IGroup, bool>> groupQuery)
        {
            if (userQuery == null && groupQuery == null) return null;
            List<AzurecpResult> allSearchResults = new List<AzurecpResult>();
            var lockResults = new object();

            foreach (AzureTenant coco in this.CurrentConfiguration.AzureTenants)
            //Parallel.ForEach(this.CurrentConfiguration.AzureTenants, async coco =>
            //var queryTenantTasks = this.CurrentConfiguration.AzureTenants.Select (async coco =>
            {
                Stopwatch timer = new Stopwatch();
                List<AzurecpResult> searchResult = null;
                try
                {
                    timer.Start();
                    searchResult = await QueryAzureADAsync(coco, userQuery, groupQuery, true).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    AzureCPLogging.LogException(ProviderInternalName, String.Format("in QueryAzureADCollectionAsync while querying tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, ex);
                }
                finally
                {
                    timer.Stop();
                }

                if (searchResult != null && searchResult.Count > 0)
                {
                    lock (lockResults)
                    {
                        allSearchResults.AddRange(searchResult);
                        //AzureCPLogging.Log(String.Format("[{0}] Search on {1} took {2}ms and found {3} result(s) for '{4}'", ProviderInternalName, coco.TenantName, timer.ElapsedMilliseconds.ToString(), searchResult.Count.ToString(), input), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                        AzureCPLogging.Log(String.Format("[{0}] Got {1} result(s) in {2}ms from \"{3}\" with input '{4}'", ProviderInternalName, searchResult.Count.ToString(), timer.ElapsedMilliseconds.ToString(), coco.TenantName, input),
                            TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                    }
                }
                else AzureCPLogging.Log(String.Format("[{0}] Got no result in {1}ms from \"{2}\" with input '{3}'", ProviderInternalName, timer.ElapsedMilliseconds.ToString(), coco.TenantName, input), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                //});
            }
            return allSearchResults;
        }

        /// <summary>
        /// Query a specific Azure tenant
        /// </summary>
        /// <param name="userFilter"></param>
        /// <param name="groupFilter"></param>
        /// <param name="coco"></param>
        /// <returns></returns>
        private async Task<List<AzurecpResult>> QueryAzureADAsync(AzureTenant coco, Expression<Func<IUser, bool>> userQuery, Expression<Func<IGroup, bool>> groupQuery, bool firstAttempt)
        {
            List<AzurecpResult> allAADResults = new List<AzurecpResult>();
            try
            {
                object lockAddResultToCollection = new object();
                using (new SPMonitoredScope(String.Format("[{0}] Connecting to Azure AD {1}", ProviderInternalName, coco.TenantName), 1000))
                {
                    if (coco.ADClient == null)
                    {
                        ActiveDirectoryClient activeDirectoryClient;
                        try
                        {
                            activeDirectoryClient = AuthenticationHelper.GetActiveDirectoryClientAsApplication(coco.TenantName, coco.TenantId, coco.ClientId, coco.ClientSecret);
                        }
                        catch (AuthenticationException ex)
                        {
                            //You should implement retry and back-off logic per the guidance given here:http://msdn.microsoft.com/en-us/library/dn168916.aspx
                            //InnerException Message will contain the HTTP error status codes mentioned in the link above
                            AzureCPLogging.LogException(ProviderInternalName, String.Format("while acquiring token for tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, ex);
                            return null;
                        }
                        coco.ADClient = activeDirectoryClient;
                        AzureCPLogging.Log(String.Format("[{0}] Got new access token for tenant '{1}'", ProviderInternalName, coco.TenantName), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                    }

                    Task userQueryTask = Task.Run(async () =>
                    {
                        try
                        {
                            if (userQuery == null) return;
                            IUserCollection userCollection = coco.ADClient.Users;
                            IPagedCollection<IUser> userSearchResults = await userCollection.Where(userQuery).ExecuteAsync().ConfigureAwait(false);
                            if (userSearchResults == null) return;
                            do
                            {
                                List<IUser> searchResultsList = userSearchResults.CurrentPage.ToList();
                                foreach (IDirectoryObject objectResult in searchResultsList)
                                {
                                    AzurecpResult azurecpResult = new AzurecpResult();
                                    azurecpResult.DirectoryObjectResult = objectResult as DirectoryObject;
                                    azurecpResult.TenantId = coco.TenantId;
                                    lock (lockAddResultToCollection)
                                    {
                                        allAADResults.Add(azurecpResult);
                                    }
                                }
                                userSearchResults = await userSearchResults.GetNextPageAsync().ConfigureAwait(false);
                            } while (userSearchResults != null && userSearchResults.MorePagesAvailable);
                        }
                        catch (Exception ex)
                        {
                            AzureCPLogging.LogException(ProviderInternalName, String.Format("while getting users in tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, ex);
                            throw ex;
                        }
                    });

                    Task groupQueryTask = Task.Run(async () =>
                    {
                        try
                        {
                            if (groupQuery == null) return;
                            IGroupCollection groupCollection = coco.ADClient.Groups;
                            IPagedCollection<IGroup> groupSearchResults = await groupCollection.Where(groupQuery).ExecuteAsync().ConfigureAwait(false);
                            if (groupSearchResults == null) return;
                            do
                            {
                                List<IGroup> searchResultsList = groupSearchResults.CurrentPage.ToList();
                                foreach (IDirectoryObject objectResult in searchResultsList)
                                {
                                    AzurecpResult azurecpResult = new AzurecpResult();
                                    azurecpResult.DirectoryObjectResult = objectResult as DirectoryObject;
                                    azurecpResult.TenantId = coco.TenantId;
                                    lock (lockAddResultToCollection)
                                    {
                                        allAADResults.Add(azurecpResult);
                                    }
                                }
                                groupSearchResults = await groupSearchResults.GetNextPageAsync().ConfigureAwait(false);
                            } while (groupSearchResults != null && groupSearchResults.MorePagesAvailable);
                        }
                        catch (Exception ex)
                        {
                            AzureCPLogging.LogException(ProviderInternalName, String.Format("while getting groups in tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, ex);
                            throw ex;
                        }
                    });

                    await Task.WhenAll(userQueryTask, groupQueryTask);
                    #region tests
                    //Task<IPagedCollection<IUser>> userQueryTask = Task.Run(async () =>
                    //{
                    //    if (userQuery == null) return null;
                    //    IUserCollection userCollection = coco.ADClient.Users;
                    //    return await userCollection.Where(userQuery).ExecuteAsync();
                    //});

                    //Task<IPagedCollection<IGroup>> groupQueryTask = Task.Run(async () =>
                    //{
                    //    if (groupQuery == null) return null;
                    //    IGroupCollection groupCollection = coco.ADClient.Groups;
                    //    return await groupCollection.Where(groupQuery).ExecuteAsync();
                    //});

                    ////Task<IPagedCollection<IGroup>> groupQueryTask = new Task<IPagedCollection<IGroup>>(() =>
                    ////{
                    ////    if (groupQuery == null) return null;
                    ////    IGroupCollection groupCollection = coco.ADClient.Groups;
                    ////    Task<IPagedCollection<IGroup>> result = groupCollection.Where(groupQuery).ExecuteAsync();
                    ////    return result.Result;
                    ////});

                    //Task userQueryContinuationTask = userQueryTask.ContinueWith((t) =>
                    //{
                    //    if (t.IsFaulted)
                    //    {
                    //        // Previous task running the AAD query got an error
                    //        AzureCPLogging.LogException(ProviderInternalName, String.Format("while querying users in tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, t.Exception);
                    //        return;
                    //    }
                    //    try
                    //    {
                    //        if (t.Status == TaskStatus.Canceled) return;
                    //        if (t.Result == null) return;
                    //        IPagedCollection<IUser> userSearchResults = t.Result;
                    //        do
                    //        {
                    //            List<IUser> searchResultsList = userSearchResults.CurrentPage.ToList();
                    //            foreach (IDirectoryObject objectResult in searchResultsList)
                    //            {
                    //                AzurecpResult azurecpResult = new AzurecpResult();
                    //                azurecpResult.DirectoryObjectResult = objectResult as DirectoryObject;
                    //                azurecpResult.TenantId = coco.TenantId;
                    //                lock (lockAddResultToCollection)
                    //                {
                    //                    allAADResults.Add(azurecpResult);
                    //                }
                    //            }
                    //            //userSearchResults = await userSearchResults.GetNextPageAsync().ConfigureAwait(false);
                    //            Task<IPagedCollection<IUser>> userSearchResultsTask = userSearchResults.GetNextPageAsync();
                    //            userSearchResults = userSearchResultsTask.Result;
                    //        } while (userSearchResults != null && userSearchResults.MorePagesAvailable);
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        AzureCPLogging.LogException(ProviderInternalName, String.Format("while getting user results in tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, ex);
                    //    }
                    //});

                    //Task groupQueryContinuationTask = groupQueryTask.ContinueWith((t) =>
                    //{
                    //    if (t.IsFaulted)
                    //    {
                    //        // Previous task running the AAD query got an error
                    //        AzureCPLogging.LogException(ProviderInternalName, String.Format("while querying groups in tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, t.Exception);
                    //        return;
                    //    }
                    //    try
                    //    {
                    //        if (t.Status == TaskStatus.Canceled) return;
                    //        if (t.Result == null) return;
                    //        IPagedCollection<IGroup> groupSearchResults = t.Result;
                    //        do
                    //        {
                    //            List<IGroup> searchResultsList = groupSearchResults.CurrentPage.ToList();
                    //            foreach (IDirectoryObject objectResult in searchResultsList)
                    //            {
                    //                AzurecpResult azurecpResult = new AzurecpResult();
                    //                azurecpResult.DirectoryObjectResult = objectResult as DirectoryObject;
                    //                azurecpResult.TenantId = coco.TenantId;
                    //                lock (lockAddResultToCollection)
                    //                {
                    //                    allAADResults.Add(azurecpResult);
                    //                }
                    //            }
                    //            //userSearchResults = await userSearchResults.GetNextPageAsync().ConfigureAwait(false);
                    //            Task<IPagedCollection<IGroup>> userSearchResultsTask = groupSearchResults.GetNextPageAsync();
                    //            groupSearchResults = userSearchResultsTask.Result;
                    //        } while (groupSearchResults != null && groupSearchResults.MorePagesAvailable);
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        AzureCPLogging.LogException(ProviderInternalName, String.Format("while getting group results in tenant {0}", coco.TenantName), AzureCPLogging.Categories.Lookup, ex);
                    //    }
                    //});
                    #endregion
                }
            }
            catch (Exception ex)
            {
                bool tryAgain = false;
                // Handle exceptions documented in http://blogs.msdn.com/b/aadgraphteam/archive/2014/06/02/azure-active-directory-graph-client-library-1-0-api-reference-publish.aspx
                if (ex.InnerException is ExpiredTokenException)
                {
                    // AccessToken provided as a part of GraphConnection has expired. Reset it and try to renew it
                    coco.ADClient = null;
                    tryAgain = true;
                    AzureCPLogging.Log(String.Format("[{0}] Access token of Azure AD tenant '{1}' expired. Renew it and try again: ExpiredTokenException: {2}", ProviderInternalName, coco.TenantName, ex.InnerException.Message),
                        TraceSeverity.High, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                }
                else if (ex.InnerException is AuthorizationException)
                {
                    // Insufficient privileges to complete the operation
                    AzureCPLogging.Log(String.Format("[{0}] Insufficient privileges to access tenant '{3}'. Check permissions of AzureCP application in Azure AD: AuthorizationException: {1}, Callstack: {2}.", ProviderInternalName, ex.InnerException.Message, ex.InnerException.StackTrace, coco.TenantName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                }
                else if (ex.InnerException is UnsupportedQueryException)
                {
                    // userFilter provided is not supported by the server
                    AzureCPLogging.Log(String.Format("[{0}] Invalid search filter while querying tenant '{3}', which indicates invalid object in AzureADObjects: UnsupportedQueryException: {1}, Callstack: {2}.", ProviderInternalName, ex.InnerException.Message, ex.InnerException.StackTrace, coco.TenantName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                }
                else if (ex.InnerException is ArgumentNullException)
                {
                    // objectType is null
                    AzureCPLogging.Log(String.Format("[{0}] objectType is null while querying tenant '{3}', which indicates a null or invalid ClaimEntityType in an object in AzureADObjects: ArgumentNullException: {1}, Callstack: {2}.", ProviderInternalName, ex.InnerException.Message, ex.InnerException.StackTrace, coco.TenantName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                }
                else if (ex.InnerException is AuthenticationException)
                {
                    // accessToken provided as a part of GraphConnection is not valid
                    coco.ADClient = null;
                    tryAgain = true;
                    AzureCPLogging.Log(String.Format("[{0}] accessToken provided as a part of GraphConnection is not valid while querying tenant '{3}': AuthenticationException: {1}, Callstack: {2}.", ProviderInternalName, ex.InnerException.Message, ex.InnerException.StackTrace, coco.TenantName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                }
                else if (ex.InnerException is RequestThrottledException)
                {
                    // Number of calls has exceeded the throttle limit set by the server
                    AzureCPLogging.Log(String.Format("[{0}] Number of calls exceeded the throttle limit set by the server while querying tenant '{3}': RequestThrottledException: {1}, Callstack: {2}.", ProviderInternalName, ex.InnerException.Message, ex.InnerException.StackTrace, coco.TenantName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                }
                else if (ex.InnerException is PageNotAvailableException)
                {
                    // pageToken has expired (which is not used here)
                    tryAgain = true;
                    AzureCPLogging.Log(String.Format("[{0}] pageToken expired while querying tenant {3}: PageNotAvailableException: {1}, Callstack: '{2}'.", ProviderInternalName, ex.InnerException.Message, ex.InnerException.StackTrace, coco.TenantName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                }
                else if (ex.InnerException is GraphException)
                {
                    // Non specific GraphException that must be last checked as it's base exception of all exceptions types above
                    // (documentation is wrong to say that this is a network error, it may be true but it just can't assume that)
                    tryAgain = true;
                    AzureCPLogging.Log(String.Format("[{0}] GraphException occurred while querying tenant '{3}': {1}, Callstack: {2}.", ProviderInternalName, ex.InnerException.Message, ex.InnerException.StackTrace, coco.TenantName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                }
                else if (ex.InnerException is Microsoft.Data.OData.ODataErrorException)
                {
                    // Typically occurs when app doesn't have enough privileges
                    AzureCPLogging.Log(String.Format("[{0}] ODataErrorException occurred while querying tenant '{3}': {1}, Callstack: {2}.", ProviderInternalName, ex.InnerException.Message, ex.InnerException.StackTrace, coco.TenantName), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Lookup);
                }
                else
                {
                    // Unknown exception
                    tryAgain = true;
                    AzureCPLogging.LogException(ProviderInternalName, String.Format("while querying tenant '{0}'", coco.TenantName), AzureCPLogging.Categories.Lookup, ex);
                }

                if (firstAttempt && tryAgain)
                {
                    AzureCPLogging.Log(String.Format("[{0}] Trying query one more time on tenant '{1}'...", ProviderInternalName),
                        TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Lookup);
                    allAADResults = await QueryAzureADAsync(coco, userQuery, groupQuery, false).ConfigureAwait(false);
                }
            }
            return allAADResults;
        }

        /// <summary>
        /// Gets the group membership of a user
        /// </summary>
        /// <param name="userToAugment"></param>
        /// <param name="coco"></param>
        /// <returns></returns>
        private List<AzurecpResult> GetUserMembership(User userToAugment, AzureTenant coco)
        {
            object lockAddResultToCollection = new object();
            List<AzurecpResult> searchResults = new List<AzurecpResult>();
            IUserFetcher retrievedUserFetcher = userToAugment;
            IPagedCollection<IDirectoryObject> pagedCollection = retrievedUserFetcher.MemberOf.ExecuteAsync().Result;
            //IPagedCollection<IDirectoryObject> pagedCollection = await retrievedUserFetcher.MemberOf.ExecuteAsync();
            //IPagedCollection<IDirectoryObject> pagedCollection = pagedCollectionTask.Result;
            do
            {
                List<IDirectoryObject> directoryObjects = pagedCollection.CurrentPage.ToList();
                foreach (IDirectoryObject directoryObject in directoryObjects)
                {
                    if (directoryObject is Group)
                    {
                        AzurecpResult result = new AzurecpResult();
                        Group group = directoryObject as Group;
                        result.DirectoryObjectResult = group;
                        result.TenantId = coco.TenantId;
                        lock (lockAddResultToCollection)
                        {
                            searchResults.Add(result);
                        }
                    }
                    //if (directoryObject is DirectoryRole)
                    //{
                    //    DirectoryRole role = directoryObject as DirectoryRole;
                    //}
                }
                pagedCollection = pagedCollection.GetNextPageAsync().Result;
            } while (pagedCollection != null && pagedCollection.MorePagesAvailable);
            return searchResults;
        }

        /// <summary>
        /// Create the SPClaim with proper issuer name
        /// </summary>
        /// <param name="type">Claim type</param>
        /// <param name="value">Claim value</param>
        /// <param name="valueType">Claim valueType</param>
        /// <param name="inputHasKeyword">Did the original input contain a keyword?</param>
        /// <returns></returns>
        protected virtual new SPClaim CreateClaim(string type, string value, string valueType)
        {
            string claimValue = String.Empty;
            var obj = ProcessedAzureObjects.FirstOrDefault(x => String.Equals(x.ClaimType, type, StringComparison.InvariantCultureIgnoreCase));
            claimValue = value;
            // SPClaimProvider.CreateClaim issues with SPOriginalIssuerType.ClaimProvider
            //return CreateClaim(type, claimValue, valueType);
            return new SPClaim(type, claimValue, valueType, IssuerName);
        }

        protected virtual PickerEntity CreatePickerEntityHelper(AzurecpResult result)
        {
            PickerEntity pe = CreatePickerEntity();
            SPClaim claim;
            string permissionValue = result.PermissionValue;
            string permissionClaimType = result.AzureObject.ClaimType;
            bool isIdentityClaimType = false;

            if (String.Equals(result.AzureObject.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase)
                || result.AzureObject.CreateAsIdentityClaim)
            {
                isIdentityClaimType = true;
            }

            if (result.AzureObject.CreateAsIdentityClaim)
            {
                // This azureObject is not directly linked to a claim type, so permission is created with identity claim type
                permissionClaimType = IdentityAzureObject.ClaimType;
                permissionValue = FormatPermissionValue(permissionClaimType, permissionValue, isIdentityClaimType, result);
                claim = CreateClaim(
                    permissionClaimType,
                    permissionValue,
                    IdentityAzureObject.ClaimValueType);
                pe.EntityType = IdentityAzureObject.ClaimEntityType;
            }
            else
            {
                permissionValue = FormatPermissionValue(permissionClaimType, permissionValue, isIdentityClaimType, result);
                claim = CreateClaim(
                    permissionClaimType,
                    permissionValue,
                    result.AzureObject.ClaimValueType);
                pe.EntityType = result.AzureObject.ClaimEntityType;
            }

            pe.DisplayText = FormatPermissionDisplayText(permissionClaimType, permissionValue, isIdentityClaimType, result);
            pe.Description = String.Format(
                PickerEntityOnMouseOver,
                result.AzureObject.GraphProperty.ToString(),
                result.QueryMatchValue);
            pe.Claim = claim;
            pe.IsResolved = true;
            //pe.EntityGroupName = "";

            int nbMetadata = 0;
            // Populate metadata attributes of permission created
            foreach (var entityAttrib in ProcessedAzureObjectsMetadata)
            {
                // if there is actally a value in the GraphObject, then it can be set
                string entityAttribValue = GetGraphPropertyValue(result.DirectoryObjectResult, entityAttrib.GraphProperty.ToString());
                if (!String.IsNullOrEmpty(entityAttribValue))
                {
                    pe.EntityData[entityAttrib.EntityDataKey] = entityAttribValue;
                    nbMetadata++;
                    AzureCPLogging.Log(String.Format("[{0}] Added metadata \"{1}\" with value \"{2}\" to permission", ProviderInternalName, entityAttrib.EntityDataKey, entityAttribValue), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                }
            }

            AzureCPLogging.Log(String.Format("[{0}] Created permission: display text: \"{1}\", value: \"{2}\", claim type: \"{3}\", and filled with {4} metadata.", ProviderInternalName, pe.DisplayText, pe.Claim.Value, pe.Claim.ClaimType, nbMetadata.ToString()), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
            return pe;
        }

        /// <summary>
        /// Override this method to customize value of permission created
        /// </summary>
        /// <param name="claimType"></param>
        /// <param name="claimValue"></param>
        /// <param name="netBiosName"></param>
        /// <returns></returns>
        protected virtual string FormatPermissionValue(string claimType, string claimValue, bool isIdentityClaimType, AzurecpResult result)
        {
            return claimValue;
        }

        /// <summary>
        /// Override this method to customize display text of permission created
        /// </summary>
        /// <param name="displayText"></param>
        /// <param name="claimType"></param>
        /// <param name="claimValue"></param>
        /// <param name="isIdentityClaim"></param>
        /// <param name="result"></param>
        /// <returns></returns>
        protected virtual string FormatPermissionDisplayText(string claimType, string claimValue, bool isIdentityClaimType, AzurecpResult result)
        {
            string permissionDisplayText = String.Empty;
            string valueDisplayedInPermission = String.Empty;

            if (result.AzureObject.GraphPropertyToDisplay != GraphProperty.None)
            {
                if (!isIdentityClaimType) permissionDisplayText = "(" + result.AzureObject.ClaimTypeMappingName + ") ";

                string graphPropertyToDisplayValue = GetGraphPropertyValue(result.DirectoryObjectResult, result.AzureObject.GraphPropertyToDisplay.ToString());
                if (!String.IsNullOrEmpty(graphPropertyToDisplayValue)) permissionDisplayText += graphPropertyToDisplayValue;
                else permissionDisplayText += result.PermissionValue;

            }
            else
            {
                if (isIdentityClaimType)
                {
                    permissionDisplayText = result.QueryMatchValue;
                }
                else
                {
                    permissionDisplayText = String.Format(
                        PickerEntityDisplayText,
                        result.AzureObject.ClaimTypeMappingName,
                        result.PermissionValue);
                }
            }

            return permissionDisplayText;
        }

        protected virtual PickerEntity CreatePickerEntityForSpecificClaimType(string input, AzureADObject claimTypesToResolve, bool inputHasKeyword)
        {
            List<PickerEntity> entities = CreatePickerEntityForSpecificClaimTypes(
                input,
                new List<AzureADObject>()
                    {
                        claimTypesToResolve,
                    },
                inputHasKeyword);
            return entities == null ? null : entities.First();
        }

        protected virtual List<PickerEntity> CreatePickerEntityForSpecificClaimTypes(string input, List<AzureADObject> claimTypesToResolve, bool inputHasKeyword)
        {
            List<PickerEntity> entities = new List<PickerEntity>();
            foreach (var claimTypeToResolve in claimTypesToResolve)
            {
                PickerEntity pe = CreatePickerEntity();
                SPClaim claim = CreateClaim(claimTypeToResolve.ClaimType, input, claimTypeToResolve.ClaimValueType);

                if (String.Equals(claim.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    pe.DisplayText = input;
                }
                else
                {
                    pe.DisplayText = String.Format(
                        PickerEntityDisplayText,
                        claimTypeToResolve.ClaimTypeMappingName,
                        input);
                }

                pe.EntityType = claimTypeToResolve.ClaimEntityType;
                pe.Description = String.Format(
                    PickerEntityOnMouseOver,
                    claimTypeToResolve.GraphProperty.ToString(),
                    input);

                pe.Claim = claim;
                pe.IsResolved = true;
                //pe.EntityGroupName = "";

                if (claimTypeToResolve.ClaimEntityType == SPClaimEntityTypes.User && !String.IsNullOrEmpty(claimTypeToResolve.EntityDataKey))
                {
                    pe.EntityData[claimTypeToResolve.EntityDataKey] = pe.Claim.Value;
                    AzureCPLogging.Log(String.Format("[{0}] Added metadata \"{1}\" with value \"{2}\" to permission", ProviderInternalName, claimTypeToResolve.EntityDataKey, pe.EntityData[claimTypeToResolve.EntityDataKey]), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                }
                entities.Add(pe);
                AzureCPLogging.Log(String.Format("[{0}] Created permission: display text: \"{1}\", value: \"{2}\", claim type: \"{3}\".", ProviderInternalName, pe.DisplayText, pe.Claim.Value, pe.Claim.ClaimType), TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
            }
            return entities.Count > 0 ? entities : null;
        }

        /// <summary>
        /// Called when claims provider is added to the farm. At this point the persisted object is not created yet so we can't pass actual claim type list
        /// If assemblyBinding for Newtonsoft.Json was not correctly added on the server, this method will generate an assembly load exception during feature activation
        /// Also called every 1st query in people picker
        /// </summary>
        /// <param name="claimTypes"></param>
        protected override void FillClaimTypes(List<string> claimTypes)
        {
            AzureCPLogging.Log(String.Format("[{0}] FillClaimTypes called.", ProviderInternalName), TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);
            if (claimTypes == null) return;
            try
            {
                this.Lock_Config.EnterReadLock();
                if (ProcessedAzureObjects == null) return;
                foreach (var azureObject in ProcessedAzureObjects)
                {
                    claimTypes.Add(azureObject.ClaimType);
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillClaimTypes", AzureCPLogging.Categories.Core, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
        }

        protected override void FillClaimValueTypes(List<string> claimValueTypes)
        {
            claimValueTypes.Add(WIF.ClaimValueTypes.String);
        }

        protected override void FillClaimsForEntity(Uri context, SPClaim entity, SPClaimProviderContext claimProviderContext, List<SPClaim> claims)
        {
            Augment(context, entity, claimProviderContext, claims);
        }

        protected override void FillClaimsForEntity(Uri context, SPClaim entity, List<SPClaim> claims)
        {
            Augment(context, entity, null, claims);
        }

        /// <summary>
        /// Perform augmentation of entity supplied
        /// </summary>
        /// <param name="context"></param>
        /// <param name="entity">entity to augment</param>
        /// <param name="claimProviderContext">Can be null</param>
        /// <param name="claims"></param>
        protected virtual void Augment(Uri context, SPClaim entity, SPClaimProviderContext claimProviderContext, List<SPClaim> claims)
        {
            try
            {
                Stopwatch timer = new Stopwatch();
                timer.Start();
                SPClaim decodedEntity;
                if (SPClaimProviderManager.IsUserIdentifierClaim(entity))
                    decodedEntity = SPClaimProviderManager.DecodeUserIdentifierClaim(entity);
                else
                {
                    if (SPClaimProviderManager.IsEncodedClaim(entity.Value))
                        decodedEntity = SPClaimProviderManager.Local.DecodeClaim(entity.Value);
                    else
                        decodedEntity = entity;
                }

                SPOriginalIssuerType loginType = SPOriginalIssuers.GetIssuerType(decodedEntity.OriginalIssuer);
                if (loginType != SPOriginalIssuerType.TrustedProvider && loginType != SPOriginalIssuerType.ClaimProvider)
                {
                    AzureCPLogging.Log(String.Format("[{0}] Not trying to augment '{1}' because OriginalIssuer is '{2}'.", ProviderInternalName, decodedEntity.Value, decodedEntity.OriginalIssuer),
                        TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Claims_Augmentation);
                    return;
                }

                AzureCPLogging.Log(String.Format("[{0}] Starting augmentation for user '{1}'.", ProviderInternalName, decodedEntity.Value),
                    TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Augmentation);

                if (!Initialize(context, null))
                    return;

                this.Lock_Config.EnterReadLock();
                try
                {
                    if (!this.CurrentConfiguration.AugmentAADRoles)
                        return;

                    // Check if there are groups to add in SAML token
                    var groups = this.ProcessedAzureObjects.FindAll(x => x.ClaimEntityType == SPClaimEntityTypes.FormsRole);
                    if (groups.Count == 0)
                    {
                        AzureCPLogging.Log(String.Format("[{0}] No object with ClaimEntityType = SPClaimEntityTypes.FormsRole found.", ProviderInternalName),
                            TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Augmentation);
                        return;
                    }
                    if (groups.Count != 1)
                    {
                        AzureCPLogging.Log(String.Format("[{0}] Found \"{1}\" objects configured with ClaimEntityType = SPClaimEntityTypes.FormsRole, instead of 1 expected.", ProviderInternalName),
                            TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Augmentation);
                        return;
                    }
                    AzureADObject groupObject = groups.First();

                    string input = decodedEntity.Value;

                    // Get user in AAD from UPN claim type
                    List<AzureADObject> identityObjects = ProcessedAzureObjects.FindAll(x =>
                        String.Equals(x.ClaimType, IdentityAzureObject.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                        !x.CreateAsIdentityClaim);
                    if (identityObjects.Count != 1)
                    {
                        // Expect only 1 object with claim type UPN
                        AzureCPLogging.Log(String.Format("[{0}] Found \"{1}\" objects configured with identity claim type {2} and CreateAsIdentityClaim set to false, instead of 1 expected.", ProviderInternalName, identityObjects.Count, IdentityAzureObject.ClaimType),
                            TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Augmentation);
                        return;
                    }
                    AzureADObject identityObject = identityObjects.First();

                    List<AzurecpResult> results = new List<AzurecpResult>();
                    BuildFilterAndProcessResultsAsync(input, identityObjects, true, context, null, ref results);

                    if (results.Count == 0)
                    {
                        // User not found
                        AzureCPLogging.Log(String.Format("[{0}] User with {1}='{2}' was not found in Azure tenant(s).", ProviderInternalName, identityObject.GraphProperty.ToString(), input),
                            TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Augmentation);
                        return;
                    }
                    else if (results.Count != 1)
                    {
                        // Expect only 1 user
                        AzureCPLogging.Log(String.Format("[{0}] Found \"{1}\" users with {2}='{3}' instead of 1 expected, aborting augmentation.", ProviderInternalName, results.Count, identityObject.GraphProperty.ToString(), input),
                            TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Augmentation);
                        return;
                    }
                    AzurecpResult result = results.First();

                    // Get groups this user is member of from his Azure tenant
                    AzureTenant userTenant = this.CurrentConfiguration.AzureTenants.First(x => String.Equals(x.TenantId, result.TenantId, StringComparison.InvariantCultureIgnoreCase));
                    AzureCPLogging.Log(String.Format("[{0}] Starting augmentation for user \"{1}\" on tenant {2}", ProviderInternalName, input, userTenant.TenantName),
                        TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Augmentation);

                    List<AzurecpResult> userMembership = GetUserMembership(result.DirectoryObjectResult as User, userTenant);
                    foreach (AzurecpResult groupResult in userMembership)
                    {
                        Group group = groupResult.DirectoryObjectResult as Group;
                        SPClaim claim = CreateClaim(groupObject.ClaimType, group.DisplayName, groupObject.ClaimValueType);
                        claims.Add(claim);
                        AzureCPLogging.Log(String.Format("[{0}] User {1} augmented with Azure AD group \"{2}\" (claim type {3}).", ProviderInternalName, input, group.DisplayName, groupObject.ClaimType),
                            TraceSeverity.Verbose, EventSeverity.Information, AzureCPLogging.Categories.Claims_Augmentation);
                    }
                    timer.Stop();
                    AzureCPLogging.Log(String.Format("[{0}] Augmentation of user '{1}' completed in {2}ms with {3} AAD group(s) added from '{4}'",
                        ProviderInternalName, input, timer.ElapsedMilliseconds.ToString(), userMembership.Count, userTenant.TenantName),
                        TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Augmentation);
                }
                catch (Exception ex)
                {
                    AzureCPLogging.LogException(ProviderInternalName, "in FillClaimsForEntity", AzureCPLogging.Categories.Claims_Augmentation, ex);
                }
                finally
                {
                    this.Lock_Config.ExitReadLock();
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillClaimsForEntity (parent catch)", AzureCPLogging.Categories.Claims_Augmentation, ex);
            }
        }

        protected override void FillEntityTypes(List<string> entityTypes)
        {
            entityTypes.Add(SPClaimEntityTypes.User);
            entityTypes.Add(SPClaimEntityTypes.FormsRole);
        }

        protected override void FillHierarchy(Uri context, string[] entityTypes, string hierarchyNodeID, int numberOfLevels, Microsoft.SharePoint.WebControls.SPProviderHierarchyTree hierarchy)
        {
            AzureCPLogging.Log(String.Format("[{0}] FillHierarchy called", ProviderInternalName),
                TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);

            //SPSecurity.RunWithElevatedPrivileges(delegate ()
            //{
            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                if (hierarchyNodeID == null)
                {
                    // Root level
                    //foreach (var azureObject in FinalAttributeList.Where(x => !String.IsNullOrEmpty(x.peoplePickerAttributeHierarchyNodeId) && !x.CreateAsIdentityClaim && entityTypes.Contains(x.ClaimEntityType)))
                    foreach (var azureObject in this.ProcessedAzureObjects.FindAll(x => !x.CreateAsIdentityClaim && entityTypes.Contains(x.ClaimEntityType)))
                    {
                        hierarchy.AddChild(
                            new Microsoft.SharePoint.WebControls.SPProviderHierarchyNode(
                                _ProviderInternalName,
                                azureObject.ClaimTypeMappingName,
                                azureObject.ClaimType,
                                true));
                    }
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillHierarchy", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
            //});
        }

        protected override void FillResolve(Uri context, string[] entityTypes, SPClaim resolveInput, List<Microsoft.SharePoint.WebControls.PickerEntity> resolved)
        {
            AzureCPLogging.Log(String.Format("[{0}] FillResolve(SPClaim) called, incoming claim value: \"{1}\", claim type: \"{2}\", claim issuer: \"{3}\"", ProviderInternalName, resolveInput.Value, resolveInput.ClaimType, resolveInput.OriginalIssuer),
                            TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);

            //SPSecurity.RunWithElevatedPrivileges(delegate ()
            //{
            if (!Initialize(context, entityTypes))
                return;

            // Ensure incoming claim should be validated by AzureCP
            // Must be made after call to Initialize because SPTrustedLoginProvider name must be known
            if (!String.Equals(resolveInput.OriginalIssuer, IssuerName, StringComparison.InvariantCultureIgnoreCase))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                string input = resolveInput.Value;
                // Resolve value only against the incoming claim type
                List<AzureADObject> attributes = this.ProcessedAzureObjects.FindAll(x =>
                    String.Equals(x.ClaimType, resolveInput.ClaimType, StringComparison.InvariantCultureIgnoreCase)
                    && !x.CreateAsIdentityClaim);
                if (attributes.Count != 1)
                {
                    // Should always find only 1 object at this stage
                    AzureCPLogging.Log(String.Format("[{0}] Found {1} objects that match the claim type \"{2}\", but exactly 1 is expected. Verify that there is no duplicate claim type. Aborting operation.", ProviderInternalName, attributes.Count().ToString(), resolveInput.ClaimType), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Picking);
                    return;
                }
                AzureADObject attribute = attributes.First();

                if (this.CurrentConfiguration.AlwaysResolveUserInput)
                {
                    PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                        input,
                        attribute,
                        false);
                    if (entity != null)
                    {
                        resolved.Add(entity);
                        AzureCPLogging.Log(String.Format("[{0}] Validated permission without AAD lookup because AzureCP configured to always resolve input. Claim value: \"{1}\", Claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                    }
                    return;
                }

                // Claims provider is called by static methods in SPClaimProviderOperations class. As a consequence, results must be declared in the method (and not in the class) to ensure that each thread has it own unique collection
                List<AzurecpResult> results = new List<AzurecpResult>();
                BuildFilterAndProcessResultsAsync(input, attributes, true, context, entityTypes, ref results);
                if (results != null && results.Count == 1)
                {
                    resolved.Add(results[0].PickerEntity);
                    AzureCPLogging.Log(String.Format("[{0}] Validated permission with AAD lookup. Claim value: \"{1}\", Claim type: \"{2}\"", ProviderInternalName, results[0].PickerEntity.Claim.Value, results[0].PickerEntity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                    return;
                }
                else if (!String.IsNullOrEmpty(attribute.PrefixToBypassLookup))
                {
                    // At this stage, it is impossible to know if input was originally created with the keyword that bypasses AAD lookup
                    // But it should be validated anyway since keyword is set for this claim type
                    PickerEntity entity = CreatePickerEntityForSpecificClaimType(input, attribute, false);
                    if (entity != null)
                    {
                        resolved.Add(entity);
                        AzureCPLogging.Log(String.Format("[{0}] Validated permission without LDAP lookup because corresponding claim type has a keyword associated. Claim value: \"{1}\", Claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                        return;
                    }
                }
                else if (results != null && results.Count != 1)
                {
                    AzureCPLogging.Log(String.Format("[{0}] Validation with AAD lookup created {1} permissions instead of 1 expected. Aborting operation", ProviderInternalName, results.Count.ToString()), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Picking);
                    return;
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillResolve(SPClaim)", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
            //});
        }

        protected override void FillResolve(Uri context, string[] entityTypes, string resolveInput, List<Microsoft.SharePoint.WebControls.PickerEntity> resolved)
        {
            AzureCPLogging.Log(String.Format("[{0}] FillResolve(string) called, incoming input \"{1}\"", ProviderInternalName, resolveInput),
                TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);

            //SPSecurity.RunWithElevatedPrivileges(delegate ()
            //{
            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                string input = resolveInput;
                // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                List<AzureADObject> azureObjects = this.ProcessedAzureObjects.FindAll(x => entityTypes.Contains(x.ClaimEntityType));

                if (this.CurrentConfiguration.AlwaysResolveUserInput)
                {
                    List<PickerEntity> entities = CreatePickerEntityForSpecificClaimTypes(
                        input,
                        azureObjects.FindAll(x => !x.CreateAsIdentityClaim),
                        false);
                    if (entities != null)
                    {
                        foreach (var entity in entities)
                        {
                            resolved.Add(entity);
                            AzureCPLogging.Log(String.Format("[{0}] Added permission created without AAD lookup because AzureCP configured to always resolve input: claim value: {1}, claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                        }
                    }
                    return;
                }

                // Check if input starts with an AzureADObject.PrefixToBypassLookup
                List<AzureADObject> objectsMatchingInputPrefix = azureObjects.FindAll(x =>
                    !String.IsNullOrEmpty(x.PrefixToBypassLookup) &&
                    input.StartsWith(x.PrefixToBypassLookup, StringComparison.InvariantCultureIgnoreCase));
                if (objectsMatchingInputPrefix.Count > 0)
                {
                    // Input has a prefix, so it should be validated with no lookup
                    AzureADObject objectMatchingInputPrefix = objectsMatchingInputPrefix.First();
                    if (objectsMatchingInputPrefix.Count > 1)
                    {
                        // Multiple objects have same prefix, which is bad
                        AzureCPLogging.Log(String.Format("[{0}] Multiple objects have same prefix '{1}', which is bad.", ProviderInternalName, objectMatchingInputPrefix.PrefixToBypassLookup), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Picking);
                        return;
                    }

                    // Get PickerEntity from the current objectMatchingInputPrefix
                    PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                        input.Substring(objectMatchingInputPrefix.PrefixToBypassLookup.Length),
                        objectMatchingInputPrefix,
                        true);
                    if (entity != null)
                    {
                        resolved.Add(entity);
                        AzureCPLogging.Log(String.Format("[{0}] Added permission created without AAD lookup because input matches a keyword: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                    }
                }
                else
                {
                    // Perform AAD lookup
                    // Claims provider is called by static methods in SPClaimProviderOperations class. As a consequence, results must be declared in the method (and not in the class) to ensure that each thread has it own unique collection
                    List<AzurecpResult> results = new List<AzurecpResult>();
                    BuildFilterAndProcessResultsAsync(
                        input,
                        azureObjects,
                        this.CurrentConfiguration.FilterExactMatchOnly,
                        context,
                        entityTypes,
                        ref results);

                    if (results != null && results.Count > 0)
                    {
                        foreach (var result in results)
                        {
                            resolved.Add(result.PickerEntity);
                            AzureCPLogging.Log(String.Format("[{0}] Added permission created with AAD lookup: claim value: \"{1}\", claim type: \"{2}\"", ProviderInternalName, result.PickerEntity.Claim.Value, result.PickerEntity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillResolve(string)", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
            //});
        }

        protected override void FillSchema(Microsoft.SharePoint.WebControls.SPProviderSchema schema)
        {
            //add the schema element we need at a minimum in our picker node
            schema.AddSchemaElement(new SPSchemaElement(PeopleEditorEntityDataKeys.DisplayName, "Display Name", SPSchemaElementType.Both));
        }

        protected override void FillSearch(Uri context, string[] entityTypes, string searchPattern, string hierarchyNodeID, int maxCount, Microsoft.SharePoint.WebControls.SPProviderHierarchyTree searchTree)
        {
            AzureCPLogging.Log(String.Format("[{0}] FillSearch called, incoming input: \"{1}\"", ProviderInternalName, searchPattern),
                TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);

            //SPSecurity.RunWithElevatedPrivileges(delegate ()
            //{
            if (!Initialize(context, entityTypes))
                return;

            this.Lock_Config.EnterReadLock();
            try
            {
                string input = searchPattern;
                SPProviderHierarchyNode matchNode = null;
                // List<T>.FindAll returns an empty list if no result found: http://msdn.microsoft.com/en-us/library/fh1w7y8z(v=vs.110).aspx
                List<AzureADObject> azureObjects;
                if (!String.IsNullOrEmpty(hierarchyNodeID))
                {
                    // Restrict search to objects currently selected in the hierarchy (may return multiple results if identity claim type)
                    azureObjects = this.ProcessedAzureObjects.FindAll(x =>
                        String.Equals(x.ClaimType, hierarchyNodeID, StringComparison.InvariantCultureIgnoreCase) &&
                        entityTypes.Contains(x.ClaimEntityType));
                }
                else
                {
                    azureObjects = this.ProcessedAzureObjects.FindAll(x => entityTypes.Contains(x.ClaimEntityType));
                }

                if (this.CurrentConfiguration.AlwaysResolveUserInput)
                {
                    List<PickerEntity> entities = CreatePickerEntityForSpecificClaimTypes(
                        input,
                        azureObjects.FindAll(x => !x.CreateAsIdentityClaim),
                        false);
                    if (entities != null)
                    {
                        foreach (var entity in entities)
                        {
                            // Add current PickerEntity to the corresponding attribute in the hierarchy
                            // Use Claim type has key
                            string entityClaimType = entity.Claim.ClaimType;
                            // ClaimTypeMappingName cannot be null as it is value of SPClaimTypeMapping.IncomingClaimTypeDisplayName, which is mandatory
                            string ClaimTypeMappingName = azureObjects
                                .First(x =>
                                    !x.CreateAsIdentityClaim &&
                                    String.Equals(x.ClaimType, entityClaimType, StringComparison.InvariantCultureIgnoreCase))
                                .ClaimTypeMappingName;

                            if (searchTree.HasChild(entityClaimType))
                            {
                                matchNode = searchTree.Children.First(x => String.Equals(x.HierarchyNodeID, entityClaimType, StringComparison.InvariantCultureIgnoreCase));
                            }
                            else
                            {
                                matchNode = new SPProviderHierarchyNode(_ProviderInternalName, ClaimTypeMappingName, entityClaimType, true);
                                searchTree.AddChild(matchNode);
                            }
                            matchNode.AddEntity(entity);
                            AzureCPLogging.Log(String.Format("[{0}] Added permission created without AAD lookup because AzureCP configured to always resolve input: claim value: \"{1}\", claim type: \"{2}\" to the list of results.", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                        }
                    }
                    return;
                }

                // Check if input starts with PrefixToBypassLookup in a AzureADObject
                List<AzureADObject> objectsMatchingInputPrefix = azureObjects.FindAll(x =>
                    !String.IsNullOrEmpty(x.PrefixToBypassLookup) &&
                    input.StartsWith(x.PrefixToBypassLookup, StringComparison.InvariantCultureIgnoreCase));
                if (objectsMatchingInputPrefix.Count > 0)
                {
                    // Input has a prefix, so it should be validated with no lookup
                    AzureADObject objectMatchingInputPrefix = objectsMatchingInputPrefix.First();
                    if (objectsMatchingInputPrefix.Count > 1)
                    {
                        // Multiple objects have same prefix, which is bad
                        AzureCPLogging.Log(String.Format("[{0}] Multiple objects have same prefix {1}, which is bad.", ProviderInternalName, objectMatchingInputPrefix.PrefixToBypassLookup), TraceSeverity.Unexpected, EventSeverity.Error, AzureCPLogging.Categories.Claims_Picking);
                        return;
                    }

                    PickerEntity entity = CreatePickerEntityForSpecificClaimType(
                        input.Substring(objectMatchingInputPrefix.PrefixToBypassLookup.Length),
                        objectMatchingInputPrefix,
                        true);

                    if (searchTree.HasChild(objectMatchingInputPrefix.ClaimType))
                    {
                        matchNode = searchTree.Children.First(x => String.Equals(x.HierarchyNodeID, objectMatchingInputPrefix.ClaimType, StringComparison.InvariantCultureIgnoreCase));
                    }
                    else
                    {
                        matchNode = new SPProviderHierarchyNode(_ProviderInternalName, objectMatchingInputPrefix.ClaimTypeMappingName, objectMatchingInputPrefix.ClaimType, true);
                        searchTree.AddChild(matchNode);
                    }
                    matchNode.AddEntity(entity);
                    AzureCPLogging.Log(String.Format("[{0}] Added permission created without AAD lookup because input matches a keyword: claim value: \"{1}\", claim type: \"{2}\" to the list of results.", ProviderInternalName, entity.Claim.Value, entity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                }
                else
                {
                    // Perform AAD lookup
                    // Claims provider is called by static methods in SPClaimProviderOperations class. As a consequence, results must be declared in the method (and not in the class) to ensure that each thread has it own unique collection
                    List<AzurecpResult> results = new List<AzurecpResult>();
                    BuildFilterAndProcessResultsAsync(
                        input,
                        azureObjects,
                        this.CurrentConfiguration.FilterExactMatchOnly,
                        context,
                        entityTypes,
                        ref results);

                    if (results != null && results.Count > 0)
                    {
                        foreach (var result in results)
                        {
                            // Add current PickerEntity to the corresponding attribute in the hierarchy
                            if (searchTree.HasChild(result.AzureObject.ClaimType))
                            {
                                matchNode = searchTree.Children.First(x => x.HierarchyNodeID == result.AzureObject.ClaimType);
                            }
                            else
                            {
                                matchNode = new SPProviderHierarchyNode(_ProviderInternalName, result.AzureObject.ClaimTypeMappingName, result.AzureObject.ClaimType, true);
                                searchTree.AddChild(matchNode);
                            }
                            matchNode.AddEntity(result.PickerEntity);
                            AzureCPLogging.Log(String.Format("[{0}] Added permission created with AAD lookup: claim value: \"{1}\", claim type: \"{2}\" to the list of results.", ProviderInternalName, result.PickerEntity.Claim.Value, result.PickerEntity.Claim.ClaimType), TraceSeverity.Medium, EventSeverity.Information, AzureCPLogging.Categories.Claims_Picking);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in FillSearch", AzureCPLogging.Categories.Claims_Picking, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
            //});
        }

        public override string Name { get { return ProviderInternalName; } }
        public override bool SupportsEntityInformation { get { return true; } }
        public override bool SupportsHierarchy { get { return true; } }
        public override bool SupportsResolve { get { return true; } }
        public override bool SupportsSearch { get { return true; } }
        public override bool SupportsUserKey { get { return true; } }

        /// <summary>
        /// Return the identity claim type
        /// </summary>
        /// <returns></returns>
        public override string GetClaimTypeForUserKey()
        {
            AzureCPLogging.Log(String.Format("[{0}] GetClaimTypeForUserKey called", ProviderInternalName),
                TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Core);

            if (!Initialize(null, null))
                return null;

            this.Lock_Config.EnterReadLock();
            try
            {
                return IdentityAzureObject.ClaimType;
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in GetClaimTypeForUserKey", AzureCPLogging.Categories.Rehydration, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
            return null;
        }

        /// <summary>
        /// Return the user key (SPClaim with identity claim type) from the incoming entity
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        protected override SPClaim GetUserKeyForEntity(SPClaim entity)
        {
            if (!Initialize(null, null))
                return null;

            // There are 2 scenarios:
            // 1: OriginalIssuer is "SecurityTokenService": Value looks like "05.t|yvanhost|yvand@yvanhost.local", claim type is "http://schemas.microsoft.com/sharepoint/2009/08/claims/userid" and it must be decoded properly
            // 2: OriginalIssuer is AzureCP: in this case incoming entity is valid and returned as is
            if (String.Equals(entity.OriginalIssuer, IssuerName, StringComparison.InvariantCultureIgnoreCase))
                return entity;

            SPClaimProviderManager cpm = SPClaimProviderManager.Local;
            SPClaim curUser = SPClaimProviderManager.DecodeUserIdentifierClaim(entity);

            this.Lock_Config.EnterReadLock();
            try
            {
                AzureCPLogging.Log(String.Format("[{0}] Return user key for user \"{1}\"", ProviderInternalName, entity.Value),
                    TraceSeverity.VerboseEx, EventSeverity.Information, AzureCPLogging.Categories.Rehydration);
                return CreateClaim(IdentityAzureObject.ClaimType, curUser.Value, curUser.ValueType);
            }
            catch (Exception ex)
            {
                AzureCPLogging.LogException(ProviderInternalName, "in GetUserKeyForEntity", AzureCPLogging.Categories.Rehydration, ex);
            }
            finally
            {
                this.Lock_Config.ExitReadLock();
            }
            return null;
        }
    }

    public class AzurecpResult
    {
        public DirectoryObject DirectoryObjectResult;
        public AzureADObject AzureObject;
        public PickerEntity PickerEntity;
        //public string GraphPropertyName;// Available in azureObject.GraphProperty
        public string PermissionValue;
        public string QueryMatchValue;
        public string TenantId;
    }
}
