using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Administration.Claims;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WIF = System.Security.Claims;
using static azurecp.ClaimsProviderLogging;
using System.Collections.ObjectModel;
using System.Collections;

namespace azurecp
{
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

        public AzureADObjectType DirectoryObjectType
        {
            get { return (AzureADObjectType)Enum.ToObject(typeof(AzureADObjectType), _DirectoryObjectType); }
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

        internal ClaimTypeConfig CopyCurrentObject()
        {
            ClaimTypeConfig copy = new ClaimTypeConfig()
            {
                _ClaimType = this._ClaimType,
                _DirectoryObjectProperty = this._DirectoryObjectProperty,
                _DirectoryObjectType = this._DirectoryObjectType,
                _EntityDataKey = this._EntityDataKey,
                _ClaimValueType = this._ClaimValueType,
                _CreateAsIdentityClaim = this._CreateAsIdentityClaim,
                _PrefixToBypassLookup = this._PrefixToBypassLookup,
                _DirectoryObjectPropertyToShowAsDisplayText = this._DirectoryObjectPropertyToShowAsDisplayText,
                _FilterExactMatchOnly = this._FilterExactMatchOnly,
                _ClaimTypeDisplayName = this._ClaimTypeDisplayName,
            };
            return copy;
        }

        internal void SetFromObject(ClaimTypeConfig objectToCopy)
        {
            _ClaimType = objectToCopy._ClaimType;
            _DirectoryObjectProperty = objectToCopy._DirectoryObjectProperty;
            _DirectoryObjectType = objectToCopy._DirectoryObjectType;
            _EntityDataKey = objectToCopy._EntityDataKey;
            _ClaimValueType = objectToCopy._ClaimValueType;
            _CreateAsIdentityClaim = objectToCopy._CreateAsIdentityClaim;
            _PrefixToBypassLookup = objectToCopy._PrefixToBypassLookup;
            _DirectoryObjectPropertyToShowAsDisplayText = objectToCopy._DirectoryObjectPropertyToShowAsDisplayText;
            _FilterExactMatchOnly = objectToCopy._FilterExactMatchOnly;
            _ClaimTypeDisplayName = objectToCopy._ClaimTypeDisplayName;
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
    /// Implements ICollection<ClaimTypeConfig> to add validation
    /// </summary>
    public class ClaimTypeConfigCollection : ICollection<ClaimTypeConfig>
    {   // Follows article https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.icollection-1?view=netframework-4.7.1

        /// <summary>
        /// Internal collection serialized in persisted object
        /// </summary>
        internal Collection<ClaimTypeConfig> innerCol = new Collection<ClaimTypeConfig>();

        public int Count => innerCol.Count;

        public bool IsReadOnly => false;

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
                throw new InvalidOperationException($"Permission metadata '{item.EntityDataKey}' already exists in the collection for the directory object {item.DirectoryObjectType}");
            }

            if (Contains(item, new ClaimTypeConfigSameClaimType()))
            {
                throw new InvalidOperationException($"Claim type '{item.ClaimType}' already exists in the collection");
            }

            if (Contains(item, new ClaimTypeConfigEnsureUniquePrefixToBypassLookup()))
            {
                throw new InvalidOperationException($"Prefix '{item.PrefixToBypassLookup}' is already set with another claim type and must be unique");
            }

            if (Contains(item))
            {
                if (String.IsNullOrEmpty(item.ClaimType))
                    throw new InvalidOperationException($"This configuration with DirectoryObjectProperty '{item.DirectoryObjectProperty}' and DirectoryObjectType '{item.DirectoryObjectType}' already exists in the collection");
                else
                    throw new InvalidOperationException($"This configuration with claim type '{item.ClaimType}' already exists in the collection");
            }

            if (ClaimsProviderConstants.EnforceOnly1ClaimTypeForGroup && item.DirectoryObjectType == AzureADObjectType.Group)
            {
                if (Contains(item, new ClaimTypeConfigEnforeOnly1ClaimTypePerObjectType()))
                {
                    throw new InvalidOperationException($"A claim type for DirectoryObjectType '{AzureADObjectType.Group.ToString()}' already exists in the collection");
                }
            }

            // Cannot add item with UseMainClaimTypeOfDirectoryObject true if collection does not contain an item with same directory object type AND a claim type defined
            if (item.UseMainClaimTypeOfDirectoryObject && innerCol.FirstOrDefault (x => x.DirectoryObjectType == item.DirectoryObjectType && !String.IsNullOrEmpty(x.ClaimType)) == null)
            {
                throw new InvalidOperationException($"Cannot add this item (with UseMainClaimTypeOfDirectoryObject set to true) because collecton does not contain an item with same DirectoryObjectType '{item.DirectoryObjectType.ToString()}' AND a ClaimType set");
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
            if (String.IsNullOrEmpty(oldClaimType)) throw new ArgumentNullException("oldClaimType");
            if (newItem == null) throw new ArgumentNullException("newItem");

            // Create a temporary copy of the collection without the old item, to test if new item can be added
            ClaimTypeConfigCollection temporaryCollection = new ClaimTypeConfigCollection();
            foreach (ClaimTypeConfig curCTConfig in innerCol.Where(x => !String.Equals(x.ClaimType, oldClaimType, StringComparison.InvariantCultureIgnoreCase)))
            {
                temporaryCollection.Add(curCTConfig);
            }

            // ClaimTypeConfigCollection.Add() may thrown an exception if newItem is not valid for any reason
            temporaryCollection.Add(newItem);

            // ClaimTypeConfigCollection.Add() did not thrown an exception, current item can be safely updated
            for (int i = 0; i < innerCol.Count; i++)
            {
                ClaimTypeConfig curCT = (ClaimTypeConfig)innerCol[i];
                if (String.Equals(curCT.ClaimType, oldClaimType, StringComparison.InvariantCultureIgnoreCase))
                {
                    innerCol.ElementAt(i).SetFromObject(newItem);
                    break;
                }
            }
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
            if (array == null)
                throw new ArgumentNullException("The array cannot be null.");
            if (arrayIndex < 0)
                throw new ArgumentOutOfRangeException("The starting array index cannot be negative.");
            if (Count > array.Length - arrayIndex + 1)
                throw new ArgumentException("The destination array has fewer elements than the collection.");

            for (int i = 0; i < innerCol.Count; i++)
            {
                array[i + arrayIndex] = innerCol[i];
            }
        }

        public bool Remove(ClaimTypeConfig item)
        {
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
            if (String.IsNullOrEmpty(claimType)) throw new ArgumentNullException("claimType");
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

        void IDisposable.Dispose() { }

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
    /// Ensure that properties ClaimType, DirectoryObjectProperty and DirectoryObjectType are unique
    /// </summary>
    public class ClaimTypeConfigSameConfig : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if (String.Equals(existingCTConfig.ClaimType, newCTConfig.ClaimType, StringComparison.InvariantCultureIgnoreCase) &&
                existingCTConfig.DirectoryObjectProperty == newCTConfig.DirectoryObjectProperty &&
                existingCTConfig.DirectoryObjectType == newCTConfig.DirectoryObjectType)
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
            string hCode = ct.ClaimType + ct.DirectoryObjectProperty + ct.DirectoryObjectType;
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
            string hCode = ct.ClaimType + ct.DirectoryObjectProperty + ct.DirectoryObjectType;
            return hCode.GetHashCode();
        }
    }

    /// <summary>
    /// Ensure that property EntityDataKey is unique for the DirectoryObjectType
    /// </summary>
    public class ClaimTypeConfigSamePermissionMetadata : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if (!String.IsNullOrEmpty(newCTConfig.EntityDataKey) &&
                String.Equals(existingCTConfig.EntityDataKey, newCTConfig.EntityDataKey, StringComparison.InvariantCultureIgnoreCase) &&
                existingCTConfig.DirectoryObjectType == newCTConfig.DirectoryObjectType)
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
            string hCode = ct.ClaimType + ct.DirectoryObjectProperty + ct.DirectoryObjectType;
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
    /// Should be used only to ensure that only 1 claim type is set per DirectoryObjectType
    /// </summary>
    internal class ClaimTypeConfigEnforeOnly1ClaimTypePerObjectType : EqualityComparer<ClaimTypeConfig>
    {
        public override bool Equals(ClaimTypeConfig existingCTConfig, ClaimTypeConfig newCTConfig)
        {
            if ((!String.IsNullOrEmpty(newCTConfig.ClaimType) && !String.IsNullOrEmpty(existingCTConfig.ClaimType)) &&
                existingCTConfig.DirectoryObjectType == newCTConfig.DirectoryObjectType &&
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
            string hCode = ct.ClaimType + ct.DirectoryObjectType + ct.UseMainClaimTypeOfDirectoryObject.ToString();
            return hCode.GetHashCode();
        }
    }
}
