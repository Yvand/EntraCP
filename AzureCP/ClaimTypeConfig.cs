﻿using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using WIF = System.Security.Claims;

namespace azurecp
{
    public class IdentityClaimTypeConfig : ClaimTypeConfig
    {
        public AzureADObjectProperty DirectoryObjectPropertyForGuestUsers
        {
            get { return (AzureADObjectProperty)Enum.ToObject(typeof(AzureADObjectProperty), _DirectoryObjectPropertyForGuestUsers); }
            set { _DirectoryObjectPropertyForGuestUsers = (int)value; }
        }
        [Persisted]
        private int _DirectoryObjectPropertyForGuestUsers = (int)AzureADObjectProperty.Mail;

        public IdentityClaimTypeConfig()
        {
        }

        public IdentityClaimTypeConfig(ClaimTypeConfig ctConfig)
        {
            this.DirectoryObjectPropertyForGuestUsers = ((IdentityClaimTypeConfig)ctConfig).DirectoryObjectPropertyForGuestUsers;
        }

        public static IdentityClaimTypeConfig ConvertClaimTypeConfig(ClaimTypeConfig ctConfig)
        {
            IdentityClaimTypeConfig identityCTConfig = new IdentityClaimTypeConfig(ctConfig);
            identityCTConfig.ClaimType = ctConfig.ClaimType;
            identityCTConfig.ClaimTypeDisplayName = ctConfig.ClaimTypeDisplayName;
            identityCTConfig.ClaimValueType = ctConfig.ClaimValueType;
            identityCTConfig.DirectoryObjectProperty = ctConfig.DirectoryObjectProperty;
            identityCTConfig.DirectoryObjectPropertyToShowAsDisplayText = ctConfig.DirectoryObjectPropertyToShowAsDisplayText;
            identityCTConfig.EntityDataKey = ctConfig.EntityDataKey;
            identityCTConfig.EntityType = ctConfig.EntityType;
            identityCTConfig.FilterExactMatchOnly = ctConfig.FilterExactMatchOnly;
            identityCTConfig.PrefixToBypassLookup = ctConfig.PrefixToBypassLookup;
            identityCTConfig.UseMainClaimTypeOfDirectoryObject = ctConfig.UseMainClaimTypeOfDirectoryObject;
            return identityCTConfig;
        }
    }

    /// <summary>
    /// Stores configuration associated to a claim type, and its mapping with the Azure AD attribute (GraphProperty)
    /// </summary>
    public class ClaimTypeConfig : SPAutoSerializingObject, IEquatable<ClaimTypeConfig>
    {
        /// <summary>
        /// Azure AD attribute mapped to the claim type
        /// </summary>
        public AzureADObjectProperty DirectoryObjectProperty
        {
            get { return (AzureADObjectProperty)Enum.ToObject(typeof(AzureADObjectProperty), _DirectoryObjectProperty); }
            set { _DirectoryObjectProperty = (int)value; }
        }
        [Persisted]
        private int _DirectoryObjectProperty;

        public DirectoryObjectType EntityType
        {
            get { return (DirectoryObjectType)Enum.ToObject(typeof(DirectoryObjectType), _DirectoryObjectType); }
            set { _DirectoryObjectType = (int)value; }
        }
        [Persisted]
        private int _DirectoryObjectType;

        public string ClaimType
        {
            get { return _ClaimType; }
            set { _ClaimType = value; }
        }
        [Persisted]
        private string _ClaimType;

        internal bool SupportsWildcard
        {
            get
            {
                if (this.DirectoryObjectProperty == AzureADObjectProperty.Id)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
        }

        /// <summary>
        /// If set to true, property ClaimType should not be set
        /// </summary>
        public bool UseMainClaimTypeOfDirectoryObject
        {
            get { return _CreateAsIdentityClaim; }
            set { _CreateAsIdentityClaim = value; }
        }
        [Persisted]
        private bool _CreateAsIdentityClaim = false;

        /// <summary>
        /// Can contain a member of class PeopleEditorEntityDataKey http://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.webcontrols.peopleeditorentitydatakeys_members(v=office.15).aspx
        /// to populate additional metadata in permission created
        /// </summary>
        public string EntityDataKey
        {
            get { return _EntityDataKey; }
            set { _EntityDataKey = value; }
        }
        [Persisted]
        private string _EntityDataKey;

        /// <summary>
        /// Stores property SPTrustedClaimTypeInformation.DisplayName of current claim type.
        /// </summary>
        public string ClaimTypeDisplayName
        {
            get { return _ClaimTypeDisplayName; }
            set { _ClaimTypeDisplayName = value; }
        }
        [Persisted]
        private string _ClaimTypeDisplayName;

        /// <summary>
        /// Every claim value type is String by default
        /// </summary>
        public string ClaimValueType
        {
            get { return _ClaimValueType; }
            set { _ClaimValueType = value; }
        }
        [Persisted]
        private string _ClaimValueType = WIF.ClaimValueTypes.String;



        /// <summary>
        /// If set, its value can be used as a prefix in the people picker to create a permission without actually quyerying Azure AD
        /// </summary>
        public string PrefixToBypassLookup
        {
            get { return _PrefixToBypassLookup; }
            set { _PrefixToBypassLookup = value; }
        }
        [Persisted]
        private string _PrefixToBypassLookup;

        public AzureADObjectProperty DirectoryObjectPropertyToShowAsDisplayText
        {
            get { return (AzureADObjectProperty)Enum.ToObject(typeof(AzureADObjectProperty), _DirectoryObjectPropertyToShowAsDisplayText); }
            set { _DirectoryObjectPropertyToShowAsDisplayText = (int)value; }
        }
        [Persisted]
        private int _DirectoryObjectPropertyToShowAsDisplayText;

        /// <summary>
        /// Set to only return values that exactly match the input
        /// </summary>
        public bool FilterExactMatchOnly
        {
            get { return _FilterExactMatchOnly; }
            set { _FilterExactMatchOnly = value; }
        }
        [Persisted]
        private bool _FilterExactMatchOnly = false;

        /// <summary>
        /// Returns a copy of the current object. This copy does not have any member of the base SharePoint base class set
        /// </summary>
        /// <returns></returns>
        public ClaimTypeConfig CopyConfiguration()
        {
            ClaimTypeConfig copy;
            if (this is IdentityClaimTypeConfig)
            {
                copy = new IdentityClaimTypeConfig();
                FieldInfo[] fieldsToCopyFromInheritedClass = typeof(IdentityClaimTypeConfig).GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly);
                foreach (FieldInfo field in fieldsToCopyFromInheritedClass)
                {
                    field.SetValue(copy, field.GetValue(this));
                }
            }
            else
            {
                copy = new ClaimTypeConfig();
            }

            // Copy non-inherited private fields
            FieldInfo[] fieldsToCopy = typeof(ClaimTypeConfig).GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly);
            foreach (FieldInfo field in fieldsToCopy)
            {
                field.SetValue(copy, field.GetValue(this));
            }
            return copy;
        }

        /// <summary>
        /// Apply configuration in parameter to current object. It does not copy SharePoint base class members
        /// </summary>
        /// <param name="configToApply"></param>
        internal void ApplyConfiguration(ClaimTypeConfig configToApply)
        {
            // Copy non-inherited private fields
            FieldInfo[] fieldsToCopy = this.GetType().GetFields(BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly);
            foreach (FieldInfo field in fieldsToCopy)
            {
                field.SetValue(this, field.GetValue(configToApply));
            }
        }

        public bool Equals(ClaimTypeConfig other)
        {
            if (new ClaimTypeConfigSameConfig().Equals(this, other))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

    /// <summary>
    /// Implements ICollection<ClaimTypeConfig> to add validation when collection is changed
    /// </summary>
    public class ClaimTypeConfigCollection : ICollection<ClaimTypeConfig>
    {   // Follows article https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.icollection-1?view=netframework-4.7.1

        /// <summary>
        /// Internal collection serialized in persisted object
        /// </summary>
        internal Collection<ClaimTypeConfig> innerCol = new Collection<ClaimTypeConfig>();

        public int Count
        {
            get
            {
                // If innerCol is null, it means that collection is not correctly set in the persisted object, very likely because it was migrated from a previons version of AzureCP
                // But this may not be the right place to handle this here: this should be handled in upper layer to clean the persisted object
                //if (innerCol == null)
                //{                    
                //    ClaimTypeConfigCollection newConfig = AzureCPConfig.ReturnDefaultClaimTypesConfig();
                //    this.innerCol = newConfig.innerCol;
                //}
                return innerCol.Count;
            }
        }

        public bool IsReadOnly => false;

        /// <summary>
        /// If set, more checks can be done when collection is changed
        /// </summary>
        public SPTrustedLoginProvider SPTrust
        {
            get => _SPTrust;
            set => _SPTrust = value;
        }
        private SPTrustedLoginProvider _SPTrust;

        public ClaimTypeConfigCollection()
        {
        }

        internal ClaimTypeConfigCollection(ref Collection<ClaimTypeConfig> innerCol)
        {
            this.innerCol = innerCol;
        }

        public ClaimTypeConfig this[int index]
        {
            get { return (ClaimTypeConfig)innerCol[index]; }
            set { innerCol[index] = value; }
        }

        public void Add(ClaimTypeConfig item)
        {
            Add(item, true);
        }

        internal void Add(ClaimTypeConfig item, bool strictChecks)
        {
            if (item.DirectoryObjectProperty == AzureADObjectProperty.NotSet)
            {
                throw new InvalidOperationException($"Property DirectoryObjectProperty is required");
            }

            if (item.UseMainClaimTypeOfDirectoryObject && !String.IsNullOrEmpty(item.ClaimType))
            {
                throw new InvalidOperationException($"No claim type should be set if UseMainClaimTypeOfDirectoryObject is set to true");
            }

            if (!item.UseMainClaimTypeOfDirectoryObject && String.IsNullOrEmpty(item.ClaimType) && String.IsNullOrEmpty(item.EntityDataKey))
            {
                throw new InvalidOperationException($"EntityDataKey is required if ClaimType is not set and UseMainClaimTypeOfDirectoryObject is set to false");
            }

            if (Contains(item, new ClaimTypeConfigSamePermissionMetadata()))
            {
                throw new InvalidOperationException($"Entity metadata '{item.EntityDataKey}' already exists in the collection for the directory object {item.EntityType}");
            }

            if (Contains(item, new ClaimTypeConfigSameClaimType()))
            {
                throw new InvalidOperationException($"Claim type '{item.ClaimType}' already exists in the collection");
            }

            if (Contains(item, new ClaimTypeConfigEnsureUniquePrefixToBypassLookup()))
            {
                throw new InvalidOperationException($"Prefix '{item.PrefixToBypassLookup}' is already set with another claim type and must be unique");
            }

            if (Contains(item, new ClaimTypeConfigSameDirectoryConfiguration()))
            {
                throw new InvalidOperationException($"An item with property '{item.DirectoryObjectProperty}' already exists for the object type '{item.EntityType}'");
            }

            if (Contains(item))
            {
                if (String.IsNullOrEmpty(item.ClaimType))
                {
                    throw new InvalidOperationException($"This configuration with DirectoryObjectProperty '{item.DirectoryObjectProperty}' and EntityType '{item.EntityType}' already exists in the collection");
                }
                else
                {
                    throw new InvalidOperationException($"This configuration with claim type '{item.ClaimType}' already exists in the collection");
                }
            }

            if (ClaimsProviderConstants.EnforceOnly1ClaimTypeForGroup && item.EntityType == DirectoryObjectType.Group)
            {
                if (Contains(item, new ClaimTypeConfigEnforeOnly1ClaimTypePerObjectType()))
                {
                    throw new InvalidOperationException($"A claim type for EntityType '{DirectoryObjectType.Group.ToString()}' already exists in the collection");
                }
            }

            if (strictChecks)
            {
                // If current item has UseMainClaimTypeOfDirectoryObject = true: check if another item with same EntityType AND a claim type defined
                // Another valid item may be added later, and even if not, code should handle that scenario
                if (item.UseMainClaimTypeOfDirectoryObject && innerCol.FirstOrDefault(x => x.EntityType == item.EntityType && !String.IsNullOrEmpty(x.ClaimType)) == null)
                {
                    throw new InvalidOperationException($"Cannot add this item (with UseMainClaimTypeOfDirectoryObject set to true) because collecton does not contain an item with same EntityType '{item.EntityType.ToString()}' AND a ClaimType set");
                }
            }

            // If SPTrustedLoginProvider is set, additional checks can be done
            if (SPTrust != null)
            {
                // If current claim type is identity claim type: EntityType must be User
                if (String.Equals(item.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (item.EntityType != DirectoryObjectType.User)
                    {
                        throw new InvalidOperationException($"Identity claim type must be configured with EntityType 'User'");
                    }
                }
            }

            innerCol.Add(item);
        }

        /// <summary>
        /// Only ClaimTypeConfig with property ClaimType already set can be updated
        /// </summary>
        /// <param name="oldClaimType">Claim type of ClaimTypeConfig object to update</param>
        /// <param name="newItem">New version of ClaimTypeConfig object</param>
        public void Update(string oldClaimType, ClaimTypeConfig newItem)
        {
            if (String.IsNullOrEmpty(oldClaimType)) { throw new ArgumentNullException("oldClaimType"); }
            if (newItem == null) { throw new ArgumentNullException("newItem"); }

            // If SPTrustedLoginProvider is set, additional checks can be done
            if (SPTrust != null)
            {
                // Specific checks if current claim type is identity claim type
                if (String.Equals(oldClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    // We don't allow to change claim type
                    if (!String.Equals(newItem.ClaimType, oldClaimType, StringComparison.InvariantCultureIgnoreCase))
                    {
                        throw new InvalidOperationException($"Claim type cannot be changed because current item is the configuration of the identity claim type");
                    }

                    // EntityType must be User
                    if (newItem.EntityType != DirectoryObjectType.User)
                    {
                        throw new InvalidOperationException($"Identity claim type must be configured with EntityType 'User'");
                    }
                }
            }

            // Create a temp collection that is a copy of current collection
            ClaimTypeConfigCollection testUpdateCollection = new ClaimTypeConfigCollection();
            foreach (ClaimTypeConfig curCTConfig in innerCol)
            {
                testUpdateCollection.Add(curCTConfig.CopyConfiguration(), false);
            }

            // Update ClaimTypeConfig in testUpdateCollection
            ClaimTypeConfig ctConfigToUpdate = testUpdateCollection.First(x => String.Equals(x.ClaimType, oldClaimType, StringComparison.InvariantCultureIgnoreCase));
            ctConfigToUpdate.ApplyConfiguration(newItem);

            // Test change in testUpdateCollection by adding all items in a new temp collection
            ClaimTypeConfigCollection testNewItemCollection = new ClaimTypeConfigCollection();
            foreach (ClaimTypeConfig curCTConfig in testUpdateCollection)
            {
                // ClaimTypeConfigCollection.Add() may thrown an exception if newItem is not valid for any reason
                testNewItemCollection.Add(curCTConfig, false);
            }

            // No error, current collection can safely be updated
            innerCol.First(x => String.Equals(x.ClaimType, oldClaimType, StringComparison.InvariantCultureIgnoreCase)).ApplyConfiguration(newItem);
        }

        /// <summary>
        /// Update the DirectoryObjectProperty of the identity ClaimTypeConfig. If new value duplicates an existing item, it will be removed from the collection
        /// </summary>
        /// <param name="newIdentifier">new DirectoryObjectProperty</param>
        /// <returns>True if the identity ClaimTypeConfig was successfully updated</returns>
        public bool UpdateUserIdentifier(AzureADObjectProperty newIdentifier)
        {
            if (newIdentifier == AzureADObjectProperty.NotSet) { throw new ArgumentNullException("newIdentifier"); }

            bool identifierUpdated = false;
            IdentityClaimTypeConfig identityClaimType = innerCol.FirstOrDefault(x => x is IdentityClaimTypeConfig) as IdentityClaimTypeConfig;
            if (identityClaimType == null)
            {
                return identifierUpdated;
            }

            if (identityClaimType.DirectoryObjectProperty == newIdentifier)
            {
                return identifierUpdated;
            }

            // Check if the new DirectoryObjectProperty duplicates an existing item, and delete it if so
            for (int i = 0; i < innerCol.Count; i++)
            {
                ClaimTypeConfig curCT = (ClaimTypeConfig)innerCol[i];
                if (curCT.EntityType == DirectoryObjectType.User &&
                    curCT.DirectoryObjectProperty == newIdentifier)
                {
                    innerCol.RemoveAt(i);
                    break;  // There can be only 1 potential duplicate
                }
            }

            identityClaimType.DirectoryObjectProperty = newIdentifier;
            identifierUpdated = true;
            return identifierUpdated;
        }

        /// <summary>
        /// Update the DirectoryObjectPropertyForGuestUsers of the identity ClaimTypeConfig.
        /// </summary>
        /// <param name="newIdentifier">new DirectoryObjectPropertyForGuestUsers</param>
        /// <returns></returns>
        public bool UpdateIdentifierForGuestUsers(AzureADObjectProperty newIdentifier)
        {
            if (newIdentifier == AzureADObjectProperty.NotSet) { throw new ArgumentNullException("newIdentifier"); }

            bool identifierUpdated = false;
            IdentityClaimTypeConfig identityClaimType = innerCol.FirstOrDefault(x => x is IdentityClaimTypeConfig) as IdentityClaimTypeConfig;
            if (identityClaimType == null)
            {
                return identifierUpdated;
            }

            if (identityClaimType.DirectoryObjectPropertyForGuestUsers == newIdentifier)
            {
                return identifierUpdated;
            }

            identityClaimType.DirectoryObjectPropertyForGuestUsers = newIdentifier;
            identifierUpdated = true;
            return identifierUpdated;
        }

        public void Clear()
        {
            innerCol.Clear();
        }

        /// <summary>
        /// Test equality based on ClaimTypeConfigSameConfig (default implementation of IEquitable<T> in ClaimTypeConfig)
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public bool Contains(ClaimTypeConfig item)
        {
            bool found = false;
            foreach (ClaimTypeConfig ct in innerCol)
            {
                if (ct.Equals(item))
                {
                    found = true;
                }
            }
            return found;
        }

        public bool Contains(ClaimTypeConfig item, EqualityComparer<ClaimTypeConfig> comp)
        {
            bool found = false;
            foreach (ClaimTypeConfig ct in innerCol)
            {
                if (comp.Equals(ct, item))
                {
                    found = true;
                }
            }
            return found;
        }

        public void CopyTo(ClaimTypeConfig[] array, int arrayIndex)
        {
            if (array == null) { throw new ArgumentNullException("The array cannot be null."); }
            if (arrayIndex < 0) { throw new ArgumentOutOfRangeException("The starting array index cannot be negative."); }
            if (Count > array.Length - arrayIndex + 1) { throw new ArgumentException("The destination array has fewer elements than the collection."); }

            for (int i = 0; i < innerCol.Count; i++)
            {
                array[i + arrayIndex] = innerCol[i];
            }
        }

        public bool Remove(ClaimTypeConfig item)
        {
            if (SPTrust != null && String.Equals(item.ClaimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase))
            {
                throw new InvalidOperationException($"Cannot delete claim type \"{item.ClaimType}\" because it is the identity claim type of \"{SPTrust.Name}\"");
            }

            bool result = false;
            for (int i = 0; i < innerCol.Count; i++)
            {
                ClaimTypeConfig curCT = (ClaimTypeConfig)innerCol[i];
                if (new ClaimTypeConfigSameConfig().Equals(curCT, item))
                {
                    innerCol.RemoveAt(i);
                    result = true;
                    break;
                }
            }
            return result;
        }

        public bool Remove(string claimType)
        {
            if (String.IsNullOrEmpty(claimType))
            {
                throw new ArgumentNullException("claimType");
            }
            if (SPTrust != null && String.Equals(claimType, SPTrust.IdentityClaimTypeInformation.MappedClaimType, StringComparison.InvariantCultureIgnoreCase))
            {
                throw new InvalidOperationException($"Cannot delete claim type \"{claimType}\" because it is the identity claim type of \"{SPTrust.Name}\"");
            }

            bool result = false;
            for (int i = 0; i < innerCol.Count; i++)
            {
                ClaimTypeConfig curCT = (ClaimTypeConfig)innerCol[i];
                if (String.Equals(claimType, curCT.ClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    innerCol.RemoveAt(i);
                    result = true;
                    break;
                }
            }
            return result;
        }

        public IEnumerator<ClaimTypeConfig> GetEnumerator()
        {
            return new ClaimTypeConfigEnumerator(this);
        }
        IEnumerator IEnumerable.GetEnumerator()
        {
            return new ClaimTypeConfigEnumerator(this);
        }

        public ClaimTypeConfig GetByClaimType(string claimType)
        {
            if (String.IsNullOrEmpty(claimType)) { throw new ArgumentNullException("claimType"); }
            ClaimTypeConfig ctConfig = innerCol.FirstOrDefault(x => String.Equals(claimType, x.ClaimType, StringComparison.InvariantCultureIgnoreCase));
            return ctConfig;
        }
    }

    public class ClaimTypeConfigEnumerator : IEnumerator<ClaimTypeConfig>
    {
        private ClaimTypeConfigCollection _collection;
        private int curIndex;
        private ClaimTypeConfig curBox;


        public ClaimTypeConfigEnumerator(ClaimTypeConfigCollection collection)
        {
            _collection = collection;
            curIndex = -1;
            curBox = default(ClaimTypeConfig);

        }

        public bool MoveNext()
        {
            //Avoids going beyond the end of the collection.
            if (++curIndex >= _collection.Count)
            {
                return false;
            }
            else
            {
                // Set current box to next item in collection.
                curBox = _collection[curIndex];
            }
            return true;
        }

        public void Reset() { curIndex = -1; }

        void IDisposable.Dispose()
        {
            // Not implemented
        }

        public ClaimTypeConfig Current
        {
            get { return curBox; }
        }

        object IEnumerator.Current
        {
            get { return Current; }
        }
    }

    /// <summary>
    /// Ensure that properties ClaimType, DirectoryObjectProperty and EntityType are unique
    /// </summary>
    public class ClaimTypeConfigSameConfig : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if (String.Equals(existingCTConfig.ClaimType, newCTConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                existingCTConfig.DirectoryObjectProperty == newCTConfig.DirectoryObjectProperty &&
                existingCTConfig.EntityType == newCTConfig.EntityType)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode(ClaimTypeConfig ct)
        {
            string hCode = ct.ClaimType + ct.DirectoryObjectProperty + ct.EntityType;
            return hCode.GetHashCode();
        }
    }

    /// <summary>
    /// Ensure that property ClaimType is unique
    /// </summary>
    public class ClaimTypeConfigSameClaimType : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if (String.Equals(existingCTConfig.ClaimType, newCTConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                !String.IsNullOrEmpty(newCTConfig.ClaimType))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode(ClaimTypeConfig ct)
        {
            string hCode = ct.ClaimType + ct.DirectoryObjectProperty + ct.EntityType;
            return hCode.GetHashCode();
        }
    }

    /// <summary>
    /// Ensure that property EntityDataKey is unique for the EntityType
    /// </summary>
    public class ClaimTypeConfigSamePermissionMetadata : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if (!String.IsNullOrEmpty(newCTConfig.EntityDataKey) &&
                String.Equals(existingCTConfig.EntityDataKey, newCTConfig.EntityDataKey, StringComparison.InvariantCultureIgnoreCase) &&
                existingCTConfig.EntityType == newCTConfig.EntityType)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode(ClaimTypeConfig ct)
        {
            string hCode = ct.ClaimType + ct.DirectoryObjectProperty + ct.EntityType;
            return hCode.GetHashCode();
        }
    }

    /// <summary>
    /// Ensure that there is no duplicate of "PrefixToBypassLookup" property
    /// </summary>
    internal class ClaimTypeConfigEnsureUniquePrefixToBypassLookup : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if (!String.IsNullOrEmpty(newCTConfig.PrefixToBypassLookup) &&
                String.Equals(newCTConfig.PrefixToBypassLookup, existingCTConfig.PrefixToBypassLookup, StringComparison.InvariantCultureIgnoreCase))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode(ClaimTypeConfig ct)
        {
            string hCode = ct.PrefixToBypassLookup;
            return hCode.GetHashCode();
        }
    }

    /// <summary>
    /// Should be used only to ensure that only 1 claim type is set per EntityType
    /// </summary>
    internal class ClaimTypeConfigEnforeOnly1ClaimTypePerObjectType : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if ((!String.IsNullOrEmpty(newCTConfig.ClaimType) && !String.IsNullOrEmpty(existingCTConfig.ClaimType)) &&
                existingCTConfig.EntityType == newCTConfig.EntityType &&
                existingCTConfig.UseMainClaimTypeOfDirectoryObject == newCTConfig.UseMainClaimTypeOfDirectoryObject &&
                newCTConfig.UseMainClaimTypeOfDirectoryObject == false)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode(ClaimTypeConfig ct)
        {
            string hCode = ct.ClaimType + ct.EntityType + ct.UseMainClaimTypeOfDirectoryObject.ToString();
            return hCode.GetHashCode();
        }
    }

    /// <summary>
    /// Check if a given object type (user or group) has 2 ClaimTypeConfig with the same LDAPAttribute and LDAPClass
    /// </summary>
    public class ClaimTypeConfigSameDirectoryConfiguration : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if (existingCTConfig.DirectoryObjectProperty == newCTConfig.DirectoryObjectProperty &&
                existingCTConfig.EntityType == newCTConfig.EntityType)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode(ClaimTypeConfig ct)
        {
            string hCode = ct.DirectoryObjectProperty.ToString() + ct.EntityType;
            return hCode.GetHashCode();
        }
    }
}
