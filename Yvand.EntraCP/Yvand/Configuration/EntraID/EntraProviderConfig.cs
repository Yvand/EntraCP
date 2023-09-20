using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using Microsoft.SharePoint.WebControls;
using System;
using System.Collections.Generic;

namespace Yvand.Config
{
    //public interface IEntraSettings : IEntraIDProviderSettings
    //{
    //    /// <summary>
    //    /// Gets the list of Azure tenants to use to get entities
    //    /// </summary>
    //    List<EntraIDTenant> EntraIDTenantList { get; }

    //    /// <summary>
    //    /// Gets the proxy address used by AzureCP to connect to Azure AD
    //    /// </summary>
    //    string ProxyAddress { get; }

    //    /// <summary>
    //    /// Gets if only security-enabled groups should be returned
    //    /// </summary>
    //    bool FilterSecurityEnabledGroupsOnly { get; }
    //}

    //public class EntraProviderSettings : EntraIDProviderSettings, IEntraSettings
    //{
    //    public List<EntraIDTenant> EntraIDTenantList { get; set; } = new List<EntraIDTenant>();
    //    public string ProxyAddress { get; set; }
    //    public bool FilterSecurityEnabledGroupsOnly { get; set; } = false;
    //    public EntraProviderSettings() : base() { }

    //    /// <summary>
    //    /// Returns the default claim types configuration list, based on the identity claim type set in the TrustedLoginProvider associated with <paramref name="claimProviderName"/>
    //    /// </summary>
    //    /// <returns></returns>
    //    public static ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig(string claimsProviderName)
    //    {
    //        if (String.IsNullOrWhiteSpace(claimsProviderName))
    //        {
    //            throw new ArgumentNullException(nameof(claimsProviderName));
    //        }

    //        SPTrustedLoginProvider spTrust = Utils.GetSPTrustAssociatedWithClaimsProvider(claimsProviderName);
    //        if (spTrust == null)
    //        {
    //            Logger.Log($"No SPTrustedLoginProvider associated with claims provider '{claimsProviderName}' was found.", TraceSeverity.Unexpected, EventSeverity.Error, TraceCategory.Core);
    //            return null;
    //        }

    //        ClaimTypeConfigCollection newCTConfigCollection = new ClaimTypeConfigCollection(spTrust)
    //        {
    //            // Identity claim type. "Name" (http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name) is a reserved claim type in SharePoint that cannot be used in the SPTrust.
    //            //new ClaimTypeConfig{EntityType = DirectoryObjectType.User, DirectoryObjectProperty = AzureADObjectProperty.UserPrincipalName, ClaimType = WIF4_5.ClaimTypes.Upn},
    //            new IdentityClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.UserPrincipalName, ClaimType = spTrust.IdentityClaimTypeInformation.MappedClaimType},

    //            // Additional properties to find user and create entity with the identity claim type (UseMainClaimTypeOfDirectoryObject=true)
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.DisplayName, UseMainClaimTypeOfDirectoryObject = true, EntityDataKey = PeopleEditorEntityDataKeys.DisplayName},
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.GivenName, UseMainClaimTypeOfDirectoryObject = true}, //Yvan
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.Surname, UseMainClaimTypeOfDirectoryObject = true},   //Duhamel
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.Mail, EntityDataKey = PeopleEditorEntityDataKeys.Email, UseMainClaimTypeOfDirectoryObject = true},

    //            // Additional properties to populate metadata of entity created: no claim type set, EntityDataKey is set and UseMainClaimTypeOfDirectoryObject = false (default value)
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.MobilePhone, EntityDataKey = PeopleEditorEntityDataKeys.MobilePhone},
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.JobTitle, EntityDataKey = PeopleEditorEntityDataKeys.JobTitle},
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.Department, EntityDataKey = PeopleEditorEntityDataKeys.Department},
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.User, EntityProperty = DirectoryObjectProperty.OfficeLocation, EntityDataKey = PeopleEditorEntityDataKeys.Location},

    //            // Group
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, EntityProperty = DirectoryObjectProperty.Id, ClaimType = ClaimsProviderConstants.DefaultMainGroupClaimType, EntityPropertyToUseAsDisplayText = DirectoryObjectProperty.DisplayName},
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, EntityProperty = DirectoryObjectProperty.DisplayName, UseMainClaimTypeOfDirectoryObject = true, EntityDataKey = PeopleEditorEntityDataKeys.DisplayName},
    //            new ClaimTypeConfig{EntityType = DirectoryObjectType.Group, EntityProperty = DirectoryObjectProperty.Mail, EntityDataKey = PeopleEditorEntityDataKeys.Email},
    //        };
    //        return newCTConfigCollection;
    //    }
    //}

    //public class EntraProviderConfig<TSettings> : EntraIDProviderConfiguration<TSettings>
    //    where TSettings : IEntraSettings
    //{
    //    protected List<EntraIDTenant> EntraIDTenantList
    //    {
    //        get => _EntraIDTenantList;
    //        set => _EntraIDTenantList = value;
    //    }
    //    [Persisted]
    //    private List<EntraIDTenant> _EntraIDTenantList = new List<EntraIDTenant>();

    //    protected string ProxyAddress
    //    {
    //        get => _ProxyAddress;
    //        set => _ProxyAddress = value;
    //    }
    //    [Persisted]
    //    private string _ProxyAddress;

    //    /// <summary>
    //    /// Set if only AAD groups with securityEnabled = true should be returned - https://docs.microsoft.com/en-us/graph/api/resources/groups-overview?view=graph-rest-1.0
    //    /// </summary>
    //    protected bool FilterSecurityEnabledGroupsOnly
    //    {
    //        get => _FilterSecurityEnabledGroupsOnly;
    //        set => _FilterSecurityEnabledGroupsOnly = value;
    //    }
    //    [Persisted]
    //    private bool _FilterSecurityEnabledGroupsOnly = false;

    //    public EntraProviderConfig() : base() { }
    //    public EntraProviderConfig(string configurationName, SPPersistedObject parent, string claimsProviderName) : base(configurationName, parent, claimsProviderName)
    //    {
    //    }

    //    public override bool InitializeDefaultSettings()
    //    {
    //        return base.InitializeDefaultSettings();
    //    }

    //    //protected override bool InitializeInternalRuntimeSettings()
    //    //{
    //    //    bool success = base.InitializeInternalRuntimeSettings();
    //    //    if (this.EntraIDTenantList != null)
    //    //    {
    //    //        foreach (var tenant in this.EntraIDTenantList)
    //    //        {
    //    //            tenant.InitializeAuthentication(this.Timeout, this.ProxyAddress);
    //    //        }
    //    //    }
    //    //    return success;
    //    //}

    //    //protected override TSettings GenerateSettingsFromCurrentConfiguration()
    //    //{
    //    //    IEntraSettings entityProviderSettings = new EntraProviderSettings()
    //    //    {
    //    //        AlwaysResolveUserInput = this.AlwaysResolveUserInput,
    //    //        ClaimTypes = this.ClaimTypes,
    //    //        CustomData = this.CustomData,
    //    //        EnableAugmentation = this.EnableAugmentation,
    //    //        EntityDisplayTextPrefix = this.EntityDisplayTextPrefix,
    //    //        FilterExactMatchOnly = this.FilterExactMatchOnly,
    //    //        Timeout = this.Timeout,
    //    //        Version = this.Version,

    //    //        // Properties specific to type IEntraSettings
    //    //        EntraIDTenantList = this.EntraIDTenantList,
    //    //        ProxyAddress = this.ProxyAddress,
    //    //        FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly,
    //    //    };
    //    //    return (TSettings)entityProviderSettings;

    //    //    //TSettings baseEntityProviderSettings = base.GenerateSettingsFromConfiguration();
    //    //    //EntraProviderSettings entityProviderSettings = baseEntityProviderSettings as EntraProviderSettings;
    //    //    //entityProviderSettings.EntraIDTenantList = this.EntraIDTenantList;
    //    //    //entityProviderSettings.ProxyAddress = this.ProxyAddress;
    //    //    //entityProviderSettings.FilterSecurityEnabledGroupsOnly = this.FilterSecurityEnabledGroupsOnly;
    //    //    //return (TSettings)(IEntraSettings)entityProviderSettings;
    //    //}

    //    //public override void ApplySettings(TSettings settings, bool commitIfValid)
    //    //{
    //    //    if (settings == null)
    //    //    {
    //    //        return;
    //    //    }
    //    //    // Properties specific to type IEntraSettings
    //    //    this.EntraIDTenantList = settings.EntraIDTenantList;
    //    //    this.FilterSecurityEnabledGroupsOnly = settings.FilterSecurityEnabledGroupsOnly;
    //    //    this.ProxyAddress = settings.ProxyAddress;

    //    //    base.ApplySettings(settings, commitIfValid);
    //    //}

    //    //public override TSettings GetDefaultSettings()
    //    //{
    //    //    IEntraSettings entityProviderSettings = new EntraProviderSettings
    //    //    {
    //    //        ClaimTypes = EntraProviderSettings.ReturnDefaultClaimTypesConfig(this.ClaimsProviderName),
    //    //    };
    //    //    return (TSettings)entityProviderSettings;
    //    //}

    //    ///// <summary>
    //    ///// Generate and return default configuration
    //    ///// </summary>
    //    ///// <returns></returns>
    //    //public static EntraProviderConfig<TSettings> ReturnDefaultConfiguration(string claimsProviderName)
    //    //{
    //    //    EntraProviderConfig<TSettings> defaultConfig = new EntraProviderConfig<TSettings>();
    //    //    defaultConfig.ClaimsProviderName = claimsProviderName;
    //    //    defaultConfig.ClaimTypes = EntraProviderSettings.ReturnDefaultClaimTypesConfig(claimsProviderName);
    //    //    return defaultConfig;
    //    //}

    //    //public override ClaimTypeConfigCollection ReturnDefaultClaimTypesConfig()
    //    //{
    //    //    return EntraProviderSettings.ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
    //    //}

    //    //public void ResetClaimTypesList()
    //    //{
    //    //    ClaimTypes.Clear();
    //    //    ClaimTypes = EntraProviderSettings.ReturnDefaultClaimTypesConfig(this.ClaimsProviderName);
    //    //    Logger.Log($"Claim types list of configuration '{Name}' was successfully reset to default configuration",
    //    //        TraceSeverity.High, EventSeverity.Information, TraceCategory.Core);
    //    //}

    //    //public override void ValidateConfiguration()
    //    //{
    //    //    foreach (EntraIDTenant tenant in this.EntraIDTenantList)
    //    //    {
    //    //        if (tenant == null)
    //    //        {
    //    //            throw new InvalidOperationException("Configuration is not valid because a tenant is null in EntraIDTenantList");
    //    //        }

    //    //        if (String.IsNullOrWhiteSpace(tenant.Name))
    //    //        {
    //    //            throw new InvalidOperationException("Configuration is not valid because a tenant has its Name property empty");
    //    //        }

    //    //        if (String.IsNullOrWhiteSpace(tenant.ClientId))
    //    //        {
    //    //            throw new InvalidOperationException("Configuration is not valid because a tenant has its ClientId property empty");
    //    //        }
    //    //    }
    //    //    base.ValidateConfiguration();
    //    //}
    //}
}
