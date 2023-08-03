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
    /// Represents the parameters associated with a search folder.
    /// </summary>
    public sealed class SearchFolderParameters : ComplexProperty
        {
        private SearchFolderTraversal traversal;
        private FolderIdCollection rootFolderIds = new();
        private SearchFilter searchFilter;

        /// <summary>
        /// Initializes a new instance of the <see cref="SearchFolderParameters"/> class.
        /// </summary>
        internal SearchFolderParameters()
            : base()
            {
            rootFolderIds.OnChange += PropertyChanged;
            }

        /// <summary>
        /// Property changed.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        private void PropertyChanged(ComplexProperty complexProperty)
            {
            Changed();
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
                case XmlElementNames.BaseFolderIds:
                    RootFolderIds.InternalClear();
                    RootFolderIds.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Restriction:
                    reader.Read();
                    searchFilter = SearchFilter.LoadFromXml(reader);
                    return true;
                default:
                    return false;
                }
            }

        /// <summary>
        /// Reads the attributes from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal override void ReadAttributesFromXml(EwsServiceXmlReader reader)
            {
            Traversal = reader.ReadAttributeValue<SearchFolderTraversal>(XmlAttributeNames.Traversal);
            }

        /// <summary>
        /// Writes the attributes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteAttributesToXml(EwsServiceXmlWriter writer)
            {
            writer.WriteAttributeValue(XmlAttributeNames.Traversal, Traversal);
            }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
            if (SearchFilter != null)
                {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Restriction);
                SearchFilter.WriteToXml(writer);
                writer.WriteEndElement(); // Restriction
                }

            RootFolderIds.WriteToXml(writer, XmlElementNames.BaseFolderIds);
            }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal void Validate()
            {
            // Search folder must have at least one root folder id.
            if (RootFolderIds.Count == 0)
                {
                throw new ServiceValidationException(Strings.SearchParametersRootFolderIdsEmpty);
                }

            // Validate the search filter
            if (SearchFilter != null)
                {
                SearchFilter.InternalValidate();
                }
            }

        /// <summary>
        /// Gets or sets the traversal mode for the search folder.
        /// </summary>
        public SearchFolderTraversal Traversal
            {
            get { return traversal; }
            set { SetFieldValue<SearchFolderTraversal>(ref traversal, value); }
            }

        /// <summary>
        /// Gets the list of root folders the search folder searches in.
        /// </summary>
        public FolderIdCollection RootFolderIds
            {
            get { return rootFolderIds; }
            }

        /// <summary>
        /// Gets or sets the search filter associated with the search folder. Available search filter classes include
        /// SearchFilter.IsEqualTo, SearchFilter.ContainsSubstring and SearchFilter.SearchFilterCollection.
        /// </summary>
        public SearchFilter SearchFilter
            {
            get
                {
                return searchFilter;
                }

            set
                {
                if (searchFilter != null)
                    {
                    searchFilter.OnChange -= PropertyChanged;
                    }

                SetFieldValue<SearchFilter>(ref searchFilter, value);

                if (searchFilter != null)
                    {
                    searchFilter.OnChange += PropertyChanged;
                    }
                }
            }
        }
    }