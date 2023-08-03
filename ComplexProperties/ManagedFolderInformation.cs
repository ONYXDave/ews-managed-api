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
    /// <summary>
    /// Represents information for a managed folder.
    /// </summary>
    public sealed class ManagedFolderInformation : ComplexProperty
        {
        private bool? canDelete;
        private bool? canRenameOrMove;
        private bool? mustDisplayComment;
        private bool? hasQuota;
        private bool? isManagedFoldersRoot;
        private string managedFolderId;
        private string comment;
        private int? storageQuota;
        private int? folderSize;
        private string homePage;

        /// <summary>
        /// Initializes a new instance of the <see cref="ManagedFolderInformation"/> class.
        /// </summary>
        internal ManagedFolderInformation()
            : base()
            {
            }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
            switch (reader.LocalName)
                {
                case XmlElementNames.CanDelete:
                    canDelete = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.CanRenameOrMove:
                    canRenameOrMove = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.MustDisplayComment:
                    mustDisplayComment = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.HasQuota:
                    hasQuota = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.IsManagedFoldersRoot:
                    isManagedFoldersRoot = reader.ReadValue<bool>();
                    return true;
                case XmlElementNames.ManagedFolderId:
                    managedFolderId = reader.ReadValue();
                    return true;
                case XmlElementNames.Comment:
                    reader.TryReadValue(ref comment);
                    return true;
                case XmlElementNames.StorageQuota:
                    storageQuota = reader.ReadValue<int>();
                    return true;
                case XmlElementNames.FolderSize:
                    folderSize = reader.ReadValue<int>();
                    return true;
                case XmlElementNames.HomePage:
                    reader.TryReadValue(ref homePage);
                    return true;
                default:
                    return false;
                }
            }

        /// <summary>
        /// Gets a value indicating whether the user can delete objects in the folder.
        /// </summary>
        public bool? CanDelete
            {
            get { return canDelete; }
            }

        /// <summary>
        /// Gets a value indicating whether the user can rename or move objects in the folder.
        /// </summary>
        public bool? CanRenameOrMove
            {
            get { return canRenameOrMove; }
            }

        /// <summary>
        /// Gets a value indicating whether the client application must display the Comment property to the user.
        /// </summary>
        public bool? MustDisplayComment
            {
            get { return mustDisplayComment; }
            }

        /// <summary>
        /// Gets a value indicating whether the folder has a quota.
        /// </summary>
        public bool? HasQuota
            {
            get { return hasQuota; }
            }

        /// <summary>
        /// Gets a value indicating whether the folder is the root of the managed folder hierarchy.
        /// </summary>
        public bool? IsManagedFoldersRoot
            {
            get { return isManagedFoldersRoot; }
            }

        /// <summary>
        /// Gets the Managed Folder Id of the folder.
        /// </summary>
        public string ManagedFolderId
            {
            get { return managedFolderId; }
            }

        /// <summary>
        /// Gets the comment associated with the folder.
        /// </summary>
        public string Comment
            {
            get { return comment; }
            }

        /// <summary>
        /// Gets the storage quota of the folder.
        /// </summary>
        public int? StorageQuota
            {
            get { return storageQuota; }
            }

        /// <summary>
        /// Gets the size of the folder.
        /// </summary>
        public int? FolderSize
            {
            get { return folderSize; }
            }

        /// <summary>
        /// Gets the home page associated with the folder.
        /// </summary>
        public string HomePage
            {
            get { return homePage; }
            }
        }
    }