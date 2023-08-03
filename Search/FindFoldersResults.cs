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
    using System.Collections.ObjectModel;

    /// <summary>
    /// Represents the results of a folder search operation.
    /// </summary>
    public sealed class FindFoldersResults : IEnumerable<Folder>
        {
        private int totalCount;
        private int? nextPageOffset;
        private bool moreAvailable;
        private Collection<Folder> folders = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="FindFoldersResults"/> class.
        /// </summary>
        internal FindFoldersResults()
            {
            }

        /// <summary>
        /// Gets the total number of folders matching the search criteria available in the searched folder.
        /// </summary>
        public int TotalCount
            {
            get { return totalCount; }
            internal set { totalCount = value; }
            }

        /// <summary>
        /// Gets the offset that should be used with FolderView to retrieve the next page of folders in a FindFolders operation.
        /// </summary>
        public int? NextPageOffset
            {
            get { return nextPageOffset; }
            internal set { nextPageOffset = value; }
            }

        /// <summary>
        /// Gets a value indicating whether more folders matching the search criteria.
        /// are available in the searched folder. 
        /// </summary>
        public bool MoreAvailable
            {
            get { return moreAvailable; }
            internal set { moreAvailable = value; }
            }

        /// <summary>
        /// Gets a collection containing the folders that were found by the search operation.
        /// </summary>
        public Collection<Folder> Folders
            {
            get { return folders; }
            }

        #region IEnumerable<Folder> Members

        /// <summary>
        /// Returns an enumerator that iterates through the collection.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.Collections.Generic.IEnumerator`1"/> that can be used to iterate through the collection.
        /// </returns>
        public IEnumerator<Folder> GetEnumerator()
            {
            return folders.GetEnumerator();
            }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Returns an enumerator that iterates through a collection.
        /// </summary>
        /// <returns>
        /// An <see cref="T:System.Collections.IEnumerator"/> object that can be used to iterate through the collection.
        /// </returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
            return folders.GetEnumerator();
            }

        #endregion
        }
    }