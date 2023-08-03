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
    using System.IO;

    /// <summary>
    /// Represents a file attachment.
    /// </summary>
    public sealed class FileAttachment : Attachment
        {
        private string fileName;
        private Stream contentStream;
        private byte[] content;
        private Stream loadToStream;
        private bool isContactPhoto;

        /// <summary>
        /// Initializes a new instance of the <see cref="FileAttachment"/> class.
        /// </summary>
        /// <param name="owner">The owner.</param>
        internal FileAttachment(Item owner)
            : base(owner)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="FileAttachment"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal FileAttachment(ExchangeService service)
            : base(service)
            {
            }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name.</returns>
        internal override string GetXmlElementName()
            {
            return XmlElementNames.FileAttachment;
            }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        /// <param name="attachmentIndex">Index of this attachment.</param>
        internal override void Validate(int attachmentIndex)
            {
            if (string.IsNullOrEmpty(fileName) && (content == null) && (contentStream == null))
                {
                throw new ServiceValidationException(string.Format(Strings.FileAttachmentContentIsNotSet, attachmentIndex));
                }
            }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
            bool result = base.TryReadElementFromXml(reader);

            if (!result)
                {
                if (reader.LocalName == XmlElementNames.IsContactPhoto)
                    {
                    isContactPhoto = reader.ReadElementValue<bool>();
                    }
                else if (reader.LocalName == XmlElementNames.Content)
                    {
                    if (loadToStream != null)
                        {
                        reader.ReadBase64ElementValue(loadToStream);
                        }
                    else
                        {
                        // If there's a file attachment content handler, use it. Otherwise
                        // load the content into a byte array.
                        // TODO: Should we mark the attachment to indicate that content is stored elsewhere?
                        if (reader.Service.FileAttachmentContentHandler != null)
                            {
                            Stream outputStream = reader.Service.FileAttachmentContentHandler.GetOutputStream(Id);

                            if (outputStream != null)
                                {
                                reader.ReadBase64ElementValue(outputStream);
                                }
                            else
                                {
                                content = reader.ReadBase64ElementValue();
                                }
                            }
                        else
                            {
                            content = reader.ReadBase64ElementValue();
                            }
                        }

                    result = true;
                    }
                }

            return result;
            }

        /// <summary>
        /// For FileAttachment, the only thing need to patch is the AttachmentId.
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        internal override bool TryReadElementFromXmlToPatch(EwsServiceXmlReader reader)
            {
            return base.TryReadElementFromXml(reader);
            }

        /// <summary>
        /// Writes elements and content to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
            base.WriteElementsToXml(writer);

            if (writer.Service.RequestedServerVersion > ExchangeVersion.Exchange2007_SP1)
                {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.IsContactPhoto, isContactPhoto);
                }

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Content);

            if (!string.IsNullOrEmpty(FileName))
                {
                using (FileStream fileStream = new(FileName, FileMode.Open, FileAccess.Read))
                    {
                    writer.WriteBase64ElementValue(fileStream);
                    }
                }
            else if (ContentStream != null)
                {
                writer.WriteBase64ElementValue(ContentStream);
                }
            else if (Content != null)
                {
                writer.WriteBase64ElementValue(Content);
                }
            else
                {
                EwsUtilities.Assert(
                    false,
                    "FileAttachment.WriteElementsToXml",
                    "The attachment's content is not set.");
                }

            writer.WriteEndElement();
            }

        /// <summary>
        /// Loads the content of the file attachment into the specified stream. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="stream">The stream to load the content of the attachment into.</param>
        public void Load(Stream stream)
            {
            loadToStream = stream;

            try
                {
                Load();
                }
            finally
                {
                loadToStream = null;
                }
            }

        /// <summary>
        /// Loads the content of the file attachment into the specified file. Calling this method results in a call to EWS.
        /// </summary>
        /// <param name="fileName">The name of the file to load the content of the attachment into. If the file already exists, it is overwritten.</param>
        public void Load(string fileName)
            {
            loadToStream = new FileStream(fileName, FileMode.Create);

            try
                {
                Load();
                }
            finally
                {
                loadToStream.Dispose();
                loadToStream = null;
                }

            this.fileName = fileName;
            content = null;
            contentStream = null;
            }

        /// <summary>
        /// Gets the name of the file the attachment is linked to.
        /// </summary>
        public string FileName
            {
            get
                {
                return fileName;
                }

            internal set
                {
                ThrowIfThisIsNotNew();

                fileName = value;
                content = null;
                contentStream = null;
                }
            }

        /// <summary>
        /// Gets or sets the content stream.
        /// </summary>
        /// <value>The content stream.</value>
        internal Stream ContentStream
            {
            get
                {
                return contentStream;
                }

            set
                {
                ThrowIfThisIsNotNew();

                contentStream = value;
                content = null;
                fileName = null;
                }
            }

        /// <summary>
        /// Gets the content of the attachment into memory. Content is set only when Load() is called.
        /// </summary>
        public byte[] Content
            {
            get
                {
                return content;
                }

            internal set
                {
                ThrowIfThisIsNotNew();

                content = value;
                fileName = null;
                contentStream = null;
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether this attachment is a contact photo.
        /// </summary>
        public bool IsContactPhoto
            {
            get
                {
                EwsUtilities.ValidatePropertyVersion(Service, ExchangeVersion.Exchange2010, "IsContactPhoto");

                return isContactPhoto;
                }

            set
                {
                EwsUtilities.ValidatePropertyVersion(Service, ExchangeVersion.Exchange2010, "IsContactPhoto");

                ThrowIfThisIsNotNew();

                isContactPhoto = value;
                }
            }
        }
    }