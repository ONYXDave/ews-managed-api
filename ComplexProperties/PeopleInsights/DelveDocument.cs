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
    /// Represents the DelveDocument.
    /// </summary>
    public sealed class DelveDocument : InsightValue
        {
        private double rank;
        private string author;
        private string created;
        private string lastModifiedTime;
        private string defaultEncodingURL;
        private string fileType;
        private string title;
        private string lastEditor;
        private string documentId;
        private string previewURL;

        /// <summary>
        /// Gets the Rank
        /// </summary>
        public double Rank
            {
            get
                {
                return rank;
                }

            set
                {
                SetFieldValue<double>(ref rank, value);
                }
            }

        /// <summary>
        /// Gets the Author
        /// </summary>
        public string Author
            {
            get
                {
                return author;
                }

            set
                {
                SetFieldValue<string>(ref author, value);
                }
            }

        /// <summary>
        /// Gets the Created
        /// </summary>
        public string Created
            {
            get
                {
                return created;
                }

            set
                {
                SetFieldValue<string>(ref created, value);
                }
            }

        /// <summary>
        /// Gets the LastModifiedTime
        /// </summary>
        public string LastModifiedTime
            {
            get
                {
                return lastModifiedTime;
                }

            set
                {
                SetFieldValue<string>(ref lastModifiedTime, value);
                }
            }

        /// <summary>
        /// Gets the DefaultEncodingURL
        /// </summary>
        public string DefaultEncodingURL
            {
            get
                {
                return defaultEncodingURL;
                }

            set
                {
                SetFieldValue<string>(ref defaultEncodingURL, value);
                }
            }

        /// <summary>
        /// Gets the FileType
        /// </summary>
        public string FileType
            {
            get
                {
                return fileType;
                }

            set
                {
                SetFieldValue<string>(ref fileType, value);
                }
            }

        /// <summary>
        /// Gets the Title
        /// </summary>
        public string Title
            {
            get
                {
                return title;
                }

            set
                {
                SetFieldValue<string>(ref title, value);
                }
            }

        /// <summary>
        /// Gets the DocumentId
        /// </summary>
        public string DocumentId
            {
            get
                {
                return documentId;
                }

            set
                {
                SetFieldValue<string>(ref documentId, value);
                }
            }

        /// <summary>
        /// Gets the PreviewURL
        /// </summary>
        public string PreviewURL
            {
            get
                {
                return previewURL;
                }

            set
                {
                SetFieldValue<string>(ref previewURL, value);
                }
            }

        /// <summary>
        /// Gets the LastEditor
        /// </summary>
        public string LastEditor
            {
            get
                {
                return lastEditor;
                }

            set
                {
                SetFieldValue<string>(ref lastEditor, value);
                }
            }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">XML reader</param>
        /// <returns>Whether the element was read</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
            switch (reader.LocalName)
                {
                case XmlElementNames.InsightSource:
                    InsightSource = reader.ReadElementValue<string>();
                    break;
                case XmlElementNames.UpdatedUtcTicks:
                    UpdatedUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.Rank:
                    Rank = reader.ReadElementValue<double>();
                    break;
                case XmlElementNames.Author:
                    Author = reader.ReadElementValue();
                    break;
                case XmlElementNames.Created:
                    Created = reader.ReadElementValue();
                    break;
                case XmlElementNames.LastModifiedTime:
                    LastModifiedTime = reader.ReadElementValue();
                    break;
                case XmlElementNames.DefaultEncodingURL:
                    DefaultEncodingURL = reader.ReadElementValue();
                    break;
                case XmlElementNames.FileType:
                    FileType = reader.ReadElementValue();
                    break;
                case XmlElementNames.Title:
                    Title = reader.ReadElementValue();
                    break;
                case XmlElementNames.DocumentId:
                    DocumentId = reader.ReadElementValue();
                    break;
                case XmlElementNames.PreviewURL:
                    PreviewURL = reader.ReadElementValue();
                    break;
                case XmlElementNames.LastEditor:
                    LastEditor = reader.ReadElementValue();
                    break;
                default:
                    return false;
                }

            return true;
            }
        }
    }