/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
    {
    using System.Collections.Generic;

    /// <summary>
    /// Represents a simple property bag.
    /// </summary>
    /// <typeparam name="TKey">The type of the key.</typeparam>
    internal class SimplePropertyBag<TKey> : IEnumerable<KeyValuePair<TKey, object>>
        {
        private Dictionary<TKey, object> items = new();
        private List<TKey> removedItems = new();
        private List<TKey> addedItems = new();
        private List<TKey> modifiedItems = new();

        /// <summary>
        /// Add item to change list.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="changeList">The change list.</param>
        private static void InternalAddItemToChangeList(TKey key, List<TKey> changeList)
            {
            if (!changeList.Contains(key))
                {
                changeList.Add(key);
                }
            }

        /// <summary>
        /// Triggers dispatch of the change event.
        /// </summary>
        private void Changed()
            {
            if (OnChange != null)
                {
                OnChange();
                }
            }

        /// <summary>
        /// Remove item.
        /// </summary>
        /// <param name="key">The key.</param>
        private void InternalRemoveItem(TKey key)
            {
            object value;

            if (TryGetValue(key, out value))
                {
                items.Remove(key);
                removedItems.Add(key);
                Changed();
                }
            }

        /// <summary>
        /// Gets the added items.
        /// </summary>
        /// <value>The added items.</value>
        internal IEnumerable<TKey> AddedItems
            {
            get { return addedItems; }
            }

        /// <summary>
        /// Gets the removed items.
        /// </summary>
        /// <value>The removed items.</value>
        internal IEnumerable<TKey> RemovedItems
            {
            get { return removedItems; }
            }

        /// <summary>
        /// Gets the modified items.
        /// </summary>
        /// <value>The modified items.</value>
        internal IEnumerable<TKey> ModifiedItems
            {
            get { return modifiedItems; }
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="SimplePropertyBag&lt;TKey&gt;"/> class.
        /// </summary>
        public SimplePropertyBag()
            {
            }

        /// <summary>
        /// Clears the change log.
        /// </summary>
        public void ClearChangeLog()
            {
            removedItems.Clear();
            addedItems.Clear();
            modifiedItems.Clear();
            }

        /// <summary>
        /// Determines whether the specified key is in the property bag.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns>
        ///     <c>true</c> if the specified key exists; otherwise, <c>false</c>.
        /// </returns>
        public bool ContainsKey(TKey key)
            {
            return items.ContainsKey(key);
            }

        /// <summary>
        /// Tries to get value.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        /// <returns>True if value exists in property bag.</returns>
        public bool TryGetValue(TKey key, out object value)
            {
            return items.TryGetValue(key, out value);
            }

        /// <summary>
        /// Gets or sets the <see cref="System.Object"/> with the specified key.
        /// </summary>
        /// <param name="key">Key.</param>
        /// <value>Value associated with key.</value>
        public object this[TKey key]
            {
            get
                {
                object value;

                if (TryGetValue(key, out value))
                    {
                    return value;
                    }
                else
                    {
                    return null;
                    }
                }

            set
                {
                if (value == null)
                    {
                    InternalRemoveItem(key);
                    }
                else
                    {
                    // If the item was to be deleted, the deletion becomes an update.
                    if (removedItems.Remove(key))
                        {
                        InternalAddItemToChangeList(key, modifiedItems);
                        }
                    else
                        {
                        // If the property value was not set, we have a newly set property.
                        if (!ContainsKey(key))
                            {
                            InternalAddItemToChangeList(key, addedItems);
                            }
                        else
                            {
                            // The last case is that we have a modified property.
                            if (!modifiedItems.Contains(key))
                                {
                                InternalAddItemToChangeList(key, modifiedItems);
                                }
                            }
                        }

                    items[key] = value;
                    Changed();
                    }
                }
            }

        /// <summary>
        /// Occurs when Changed.
        /// </summary>
        public event PropertyBagChangedDelegate OnChange;

        #region IEnumerable<KeyValuePair<TKey,object>> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<KeyValuePair<TKey, object>> GetEnumerator()
            {
            return items.GetEnumerator();
            }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
            return items.GetEnumerator();
            }

        #endregion
        }
    }