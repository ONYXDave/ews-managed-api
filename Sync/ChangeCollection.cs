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
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Represents a collection of changes as returned by a synchronization operation.
    /// </summary>
    /// <typeparam name="TChange">Type representing the type of change (e.g. FolderChange or ItemChange)</typeparam>
    public sealed class ChangeCollection<TChange> : IEnumerable<TChange>
        where TChange : Change
        {
        private List<TChange> changes = new();
        private string syncState;
        private bool moreChangesAvailable;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChangeCollection&lt;TChange&gt;"/> class.
        /// </summary>
        internal ChangeCollection()
            {
            }

        /// <summary>
        /// Adds the specified change.
        /// </summary>
        /// <param name="change">The change.</param>
        internal void Add(TChange change)
            {
            EwsUtilities.Assert(
                change != null,
                "ChangeList.Add",
                "change is null");

            changes.Add(change);
            }

        /// <summary>
        /// Gets the number of changes in the collection.
        /// </summary>
        public int Count
            {
            get { return changes.Count; }
            }

        /// <summary>
        /// Gets an individual change from the change collection.
        /// </summary>
        /// <param name="index">Zero-based index.</param>
        /// <returns>An single change.</returns>
        public TChange this[int index]
            {
            get
                {
                if (index < 0 || index >= Count)
                    {
                    throw new ArgumentOutOfRangeException("index", Strings.IndexIsOutOfRange);
                    }

                return changes[index];
                }
            }

        /// <summary>
        /// Gets the SyncState blob returned by a synchronization operation.
        /// </summary>
        public string SyncState
            {
            get { return syncState; }
            internal set { syncState = value; }
            }

        /// <summary>
        /// Gets a value indicating whether the there are more changes to be synchronized from the server.
        /// </summary>
        public bool MoreChangesAvailable
            {
            get { return moreChangesAvailable; }
            internal set { moreChangesAvailable = value; }
            }

        #region IEnumerable<TChange> Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<TChange> GetEnumerator()
            {
            return changes.GetEnumerator();
            }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
            return changes.GetEnumerator();
            }

        #endregion
        }
    }