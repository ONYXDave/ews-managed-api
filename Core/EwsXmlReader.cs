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
    using System.IO;
    using System.Xml;

    /// <summary>
    /// XML reader.
    /// </summary>
    internal class EwsXmlReader
        {
        private const int ReadWriteBufferSize = 4096;

        #region Private members

        private XmlNodeType prevNodeType = XmlNodeType.None;
        private XmlReader xmlReader;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsXmlReader"/> class.
        /// </summary>
        /// <param name="stream">The stream.</param>
        public EwsXmlReader(Stream stream)
            {
            xmlReader = InitializeXmlReader(stream);
            }

        /// <summary>
        /// Initializes the XML reader.
        /// </summary>
        /// <param name="stream">The stream.</param>
        /// <returns>An XML reader to use.</returns>
        protected virtual XmlReader InitializeXmlReader(Stream stream)
            {
            // The ProhibitDtd property is used to indicate whether XmlReader should process DTDs or not. By default, 
            // it will do so. EWS doesn't use DTD references so we want to turn this off. Also, the XmlResolver property is
            // set to an instance of XmlUrlResolver by default. We don't want XmlTextReader to try to resolve this DTD reference 
            // so we disable the XmlResolver as well.
            XmlReaderSettings settings = new()
                {
                ConformanceLevel = ConformanceLevel.Auto,
                ProhibitDtd = true,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true,
                XmlResolver = null
                };

            XmlTextReader xmlTextReader = SafeXmlFactory.CreateSafeXmlTextReader(stream);
            xmlTextReader.Normalization = false;

            return XmlReader.Create(xmlTextReader, settings);
            }

        #endregion

        /// <summary>
        /// Formats the name of the element.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localElementName">Name of the local element.</param>
        /// <returns>Element name.</returns>
        private static string FormatElementName(string namespacePrefix, string localElementName)
            {
            return string.IsNullOrEmpty(namespacePrefix) ? localElementName : namespacePrefix + ":" + localElementName;
            }

        /// <summary>
        /// Read XML element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        /// <param name="nodeType">Type of the node.</param>
        private void InternalReadElement(
            XmlNamespace xmlNamespace,
            string localName,
            XmlNodeType nodeType)
            {
            if (xmlNamespace == XmlNamespace.NotSpecified)
                {
                InternalReadElement(
                    string.Empty,
                    localName,
                    nodeType);
                }
            else
                {
                Read(nodeType);

                if ((LocalName != localName) || (NamespaceUri != EwsUtilities.GetNamespaceUri(xmlNamespace)))
                    {
                    throw new ServiceXmlDeserializationException(
                        string.Format(
                            Strings.UnexpectedElement,
                            EwsUtilities.GetNamespacePrefix(xmlNamespace),
                            localName,
                            nodeType,
                            xmlReader.Name,
                            NodeType));
                    }
                }
            }

        /// <summary>
        /// Read XML element.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localName">Name of the local.</param>
        /// <param name="nodeType">Type of the node.</param>
        private void InternalReadElement(
            string namespacePrefix,
            string localName,
            XmlNodeType nodeType)
            {
            Read(nodeType);

            if ((LocalName != localName) || (NamespacePrefix != namespacePrefix))
                {
                throw new ServiceXmlDeserializationException(
                                string.Format(
                                    Strings.UnexpectedElement,
                                    namespacePrefix,
                                    localName,
                                    nodeType,
                                    xmlReader.Name,
                                    NodeType));
                }
            }

        /// <summary>
        /// Reads the next node.
        /// </summary>
        public void Read()
            {
            prevNodeType = xmlReader.NodeType;

            // XmlReader.Read returns true if the next node was read successfully; false if there 
            // are no more nodes to read. The caller to EwsXmlReader.Read expects that there's another node to 
            // read. Throw an exception if not true.
            bool nodeRead = xmlReader.Read();
            if (!nodeRead)
                {
                throw new ServiceXmlDeserializationException(Strings.UnexpectedEndOfXmlDocument);
                }
            }

        /// <summary>
        /// Reads the specified node type.
        /// </summary>
        /// <param name="nodeType">Type of the node.</param>
        public void Read(XmlNodeType nodeType)
            {
            Read();
            if (NodeType != nodeType)
                {
                throw new ServiceXmlDeserializationException(
                    string.Format(
                        Strings.UnexpectedElementType,
                        nodeType,
                        NodeType));
                }
            }

        /// <summary>
        /// Reads the attribute value.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="attributeName">Name of the attribute.</param>
        /// <returns>Attribute value.</returns>
        public string ReadAttributeValue(XmlNamespace xmlNamespace, string attributeName)
            {
            if (xmlNamespace == XmlNamespace.NotSpecified)
                {
                return ReadAttributeValue(attributeName);
                }
            else
                {
                return xmlReader.GetAttribute(attributeName, EwsUtilities.GetNamespaceUri(xmlNamespace));
                }
            }

        /// <summary>
        /// Reads the attribute value.
        /// </summary>
        /// <param name="attributeName">Name of the attribute.</param>
        /// <returns>Attribute value.</returns>
        public string ReadAttributeValue(string attributeName)
            {
            return xmlReader.GetAttribute(attributeName);
            }

        /// <summary>
        /// Reads the attribute value.
        /// </summary>
        /// <typeparam name="T">Type of attribute value.</typeparam>
        /// <param name="attributeName">Name of the attribute.</param>
        /// <returns>Attribute value.</returns>
        public T ReadAttributeValue<T>(string attributeName)
            {
            return EwsUtilities.Parse<T>(ReadAttributeValue(attributeName));
            }

        /// <summary>
        /// Reads a nullable attribute value.
        /// </summary>
        /// <typeparam name="T">Type of attribute value.</typeparam>
        /// <param name="attributeName">Name of the attribute.</param>
        /// <returns>Attribute value.</returns>
        public Nullable<T> ReadNullableAttributeValue<T>(string attributeName) where T : struct
            {
            string attributeValue = ReadAttributeValue(attributeName);
            if (attributeValue == null)
                {
                return null;
                }
            else
                {
                return EwsUtilities.Parse<T>(attributeValue);
                }
            }

        /// <summary>
        /// Reads the element value.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>Element value.</returns>
        public string ReadElementValue(string namespacePrefix, string localName)
            {
            if (!IsStartElement(namespacePrefix, localName))
                {
                ReadStartElement(namespacePrefix, localName);
                }

            string value = null;

            if (!IsEmptyElement)
                {
                value = ReadValue();
                }

            return value;
            }

        /// <summary>
        /// Reads the element value.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>Element value.</returns>
        public string ReadElementValue(XmlNamespace xmlNamespace, string localName)
            {
            if (!IsStartElement(xmlNamespace, localName))
                {
                ReadStartElement(xmlNamespace, localName);
                }

            string value = null;

            if (!IsEmptyElement)
                {
                value = ReadValue();
                }

            return value;
            }

        /// <summary>
        /// Reads the element value.
        /// </summary>
        /// <returns>Element value.</returns>
        public string ReadElementValue()
            {
            EnsureCurrentNodeIsStartElement();

            return ReadElementValue(NamespacePrefix, LocalName);
            }

        /// <summary>
        /// Reads the element value.
        /// </summary>
        /// <typeparam name="T">Type of element value.</typeparam>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>Element value.</returns>
        public T ReadElementValue<T>(XmlNamespace xmlNamespace, string localName)
            {
            if (!IsStartElement(xmlNamespace, localName))
                {
                ReadStartElement(xmlNamespace, localName);
                }

            T value = default(T);

            if (!IsEmptyElement)
                {
                value = ReadValue<T>();
                }

            return value;
            }

        /// <summary>
        /// Reads the element value.
        /// </summary>
        /// <typeparam name="T">Type of element value.</typeparam>
        /// <returns>Element value.</returns>
        public T ReadElementValue<T>()
            {
            EnsureCurrentNodeIsStartElement();

            string namespacePrefix = NamespacePrefix;
            string localName = LocalName;

            T value = default(T);

            if (!IsEmptyElement)
                {
                value = ReadValue<T>();
                }

            return value;
            }

        /// <summary>
        /// Reads the value.
        /// </summary>
        /// <returns>Value</returns>
        public string ReadValue()
            {
            return xmlReader.ReadString();
            }

        /// <summary>
        /// Tries to read value.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>True if value was read.</returns>
        public bool TryReadValue(ref string value)
            {
            if (!IsEmptyElement)
                {
                Read();

                if (NodeType == XmlNodeType.Text)
                    {
                    value = xmlReader.Value;
                    return true;
                    }
                else
                    {
                    return false;
                    }
                }
            else
                {
                return false;
                }
            }

        /// <summary>
        /// Reads the value.
        /// </summary>
        /// <typeparam name="T">Type of value.</typeparam>
        /// <returns>Value.</returns>
        public T ReadValue<T>()
            {
            return EwsUtilities.Parse<T>(ReadValue());
            }

        /// <summary>
        /// Reads the base64 element value.
        /// </summary>
        /// <returns>Byte array.</returns>
        public byte[] ReadBase64ElementValue()
            {
            EnsureCurrentNodeIsStartElement();

            byte[] buffer = new byte[ReadWriteBufferSize];
            int bytesRead;

            using (MemoryStream memoryStream = new())
                {
                do
                    {
                    bytesRead = xmlReader.ReadElementContentAsBase64(buffer, 0, ReadWriteBufferSize);

                    if (bytesRead > 0)
                        {
                        memoryStream.Write(buffer, 0, bytesRead);
                        }
                    }
                while (bytesRead > 0);

                // Can use MemoryStream.GetBuffer() if the buffer's capacity and the number of bytes read
                // are identical. Otherwise need to convert to byte array that's the size of the number of bytes read.
                return (memoryStream.Length == memoryStream.Capacity) ? memoryStream.GetBuffer() : memoryStream.ToArray();
                }
            }

        /// <summary>
        /// Reads the base64 element value.
        /// </summary>
        /// <param name="outputStream">The output stream.</param>
        public void ReadBase64ElementValue(Stream outputStream)
            {
            EnsureCurrentNodeIsStartElement();

            byte[] buffer = new byte[ReadWriteBufferSize];
            int bytesRead;

            do
                {
                bytesRead = xmlReader.ReadElementContentAsBase64(buffer, 0, ReadWriteBufferSize);

                if (bytesRead > 0)
                    {
                    outputStream.Write(buffer, 0, bytesRead);
                    }
                }
            while (bytesRead > 0);

            outputStream.Flush();
            }

        /// <summary>
        /// Reads the start element.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localName">Name of the local.</param>
        public void ReadStartElement(string namespacePrefix, string localName)
            {
            InternalReadElement(
                namespacePrefix,
                localName,
                XmlNodeType.Element);
            }

        /// <summary>
        /// Reads the start element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        public void ReadStartElement(XmlNamespace xmlNamespace, string localName)
            {
            InternalReadElement(
                xmlNamespace,
                localName,
                XmlNodeType.Element);
            }

        /// <summary>
        /// Reads the end element.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="elementName">Name of the element.</param>
        public void ReadEndElement(string namespacePrefix, string elementName)
            {
            InternalReadElement(
                namespacePrefix,
                elementName,
                XmlNodeType.EndElement);
            }

        /// <summary>
        /// Reads the end element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        public void ReadEndElement(XmlNamespace xmlNamespace, string localName)
            {
            InternalReadElement(
                xmlNamespace,
                localName,
                XmlNodeType.EndElement);
            }

        /// <summary>
        /// Reads the end element if necessary.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        public void ReadEndElementIfNecessary(XmlNamespace xmlNamespace, string localName)
            {
            if (!(IsStartElement(xmlNamespace, localName) && IsEmptyElement))
                {
                if (!IsEndElement(xmlNamespace, localName))
                    {
                    ReadEndElement(xmlNamespace, localName);
                    }
                }
            }

        /// <summary>
        /// Determines whether current element is a start element.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>
        ///     <c>true</c> if current element is a start element; otherwise, <c>false</c>.
        /// </returns>
        public bool IsStartElement(string namespacePrefix, string localName)
            {
            string fullyQualifiedName = FormatElementName(namespacePrefix, localName);

            return NodeType == XmlNodeType.Element && xmlReader.Name == fullyQualifiedName;
            }

        /// <summary>
        /// Determines whether current element is a start element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>
        ///     <c>true</c> if current element is a start element; otherwise, <c>false</c>.
        /// </returns>
        public bool IsStartElement(XmlNamespace xmlNamespace, string localName)
            {
            return (LocalName == localName) && IsStartElement() &&
                ((NamespacePrefix == EwsUtilities.GetNamespacePrefix(xmlNamespace)) ||
                (NamespaceUri == EwsUtilities.GetNamespaceUri(xmlNamespace)));
            }

        /// <summary>
        /// Determines whether current element is a start element.
        /// </summary>
        /// <returns>
        ///     <c>true</c> if current element is a start element; otherwise, <c>false</c>.
        /// </returns>
        public bool IsStartElement()
            {
            return NodeType == XmlNodeType.Element;
            }

        /// <summary>
        /// Determines whether current element is a end element.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>
        ///     <c>true</c> if current element is an end element; otherwise, <c>false</c>.
        /// </returns>
        public bool IsEndElement(string namespacePrefix, string localName)
            {
            string fullyQualifiedName = FormatElementName(namespacePrefix, localName);

            return NodeType == XmlNodeType.EndElement && xmlReader.Name == fullyQualifiedName;
            }

        /// <summary>
        /// Determines whether current element is a end element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        /// <returns>
        ///     <c>true</c> if current element is an end element; otherwise, <c>false</c>.
        /// </returns>
        public bool IsEndElement(XmlNamespace xmlNamespace, string localName)
            {
            return (LocalName == localName) && (NodeType == XmlNodeType.EndElement) &&
                ((NamespacePrefix == EwsUtilities.GetNamespacePrefix(xmlNamespace)) ||
                (NamespaceUri == EwsUtilities.GetNamespaceUri(xmlNamespace)));
            }

        /// <summary>
        /// Skips the element.
        /// </summary>
        /// <param name="namespacePrefix">The namespace prefix.</param>
        /// <param name="localName">Name of the local.</param>
        public void SkipElement(string namespacePrefix, string localName)
            {
            if (!IsEndElement(namespacePrefix, localName))
                {
                if (!IsStartElement(namespacePrefix, localName))
                    {
                    ReadStartElement(namespacePrefix, localName);
                    }

                if (!IsEmptyElement)
                    {
                    do
                        {
                        Read();
                        }
                    while (!IsEndElement(namespacePrefix, localName));
                    }
                }
            }

        /// <summary>
        /// Skips the element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        public void SkipElement(XmlNamespace xmlNamespace, string localName)
            {
            if (!IsEndElement(xmlNamespace, localName))
                {
                if (!IsStartElement(xmlNamespace, localName))
                    {
                    ReadStartElement(xmlNamespace, localName);
                    }

                if (!IsEmptyElement)
                    {
                    do
                        {
                        Read();
                        }
                    while (!IsEndElement(xmlNamespace, localName));
                    }
                }
            }

        /// <summary>
        /// Skips the current element.
        /// </summary>
        public void SkipCurrentElement()
            {
            SkipElement(NamespacePrefix, LocalName);
            }

        /// <summary>
        /// Ensures the current node is start element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        public void EnsureCurrentNodeIsStartElement(XmlNamespace xmlNamespace, string localName)
            {
            if (!IsStartElement(xmlNamespace, localName))
                {
                throw new ServiceXmlDeserializationException(
                    string.Format(
                        Strings.ElementNotFound,
                        localName,
                        xmlNamespace));
                }
            }

        /// <summary>
        /// Ensures the current node is start element.
        /// </summary>
        public void EnsureCurrentNodeIsStartElement()
            {
            if (NodeType != XmlNodeType.Element)
                {
                throw new ServiceXmlDeserializationException(
                    string.Format(
                        Strings.ExpectedStartElement,
                        xmlReader.Name,
                        NodeType));
                }
            }

        /// <summary>
        /// Ensures the current node is end element.
        /// </summary>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="localName">Name of the local.</param>
        public void EnsureCurrentNodeIsEndElement(XmlNamespace xmlNamespace, string localName)
            {
            if (!IsEndElement(xmlNamespace, localName))
                {
                if (!(IsStartElement(xmlNamespace, localName) && IsEmptyElement))
                    {
                    throw new ServiceXmlDeserializationException(
                        string.Format(
                            Strings.ElementNotFound,
                            localName,
                            xmlNamespace));
                    }
                }
            }

        /// <summary>
        /// Reads the Outer XML at the given location.
        /// </summary>
        /// <returns>
        /// Outer XML as string.
        /// </returns>
        public string ReadOuterXml()
            {
            if (!IsStartElement())
                {
                throw new ServiceXmlDeserializationException(Strings.CurrentPositionNotElementStart);
                }

            return xmlReader.ReadOuterXml();
            }

        /// <summary>
        /// Reads the Inner XML at the given location.
        /// </summary>
        /// <returns>
        /// Inner XML as string.
        /// </returns>
        public string ReadInnerXml()
            {
            if (!IsStartElement())
                {
                throw new ServiceXmlDeserializationException(Strings.CurrentPositionNotElementStart);
                }

            return xmlReader.ReadInnerXml();
            }

        /// <summary>
        /// Gets the XML reader for node.
        /// </summary>
        /// <returns></returns>
        internal XmlReader GetXmlReaderForNode()
            {
            return xmlReader.ReadSubtree();
            }

        /// <summary>
        /// Reads to the next descendant element with the specified local name and namespace.
        /// </summary>
        /// <param name="xmlNamespace">The namespace of the element you with to move to.</param>
        /// <param name="localName">The local name of the element you wish to move to.</param>
        public void ReadToDescendant(XmlNamespace xmlNamespace, string localName)
            {
            xmlReader.ReadToDescendant(localName, EwsUtilities.GetNamespaceUri(xmlNamespace));
            }

        /// <summary>
        /// Gets a value indicating whether this instance has attributes.
        /// </summary>
        /// <value>
        ///     <c>true</c> if this instance has attributes; otherwise, <c>false</c>.
        /// </value>
        public bool HasAttributes
            {
            get { return xmlReader.AttributeCount > 0; }
            }

        /// <summary>
        /// Gets a value indicating whether current element is empty.
        /// </summary>
        /// <value>
        ///     <c>true</c> if current element is empty element; otherwise, <c>false</c>.
        /// </value>
        public bool IsEmptyElement
            {
            get { return xmlReader.IsEmptyElement; }
            }

        /// <summary>
        /// Gets the local name of the current element.
        /// </summary>
        /// <value>The local name of the current element.</value>
        public string LocalName
            {
            get { return xmlReader.LocalName; }
            }

        /// <summary>
        /// Gets the namespace prefix.
        /// </summary>
        /// <value>The namespace prefix.</value>
        public string NamespacePrefix
            {
            get { return xmlReader.Prefix; }
            }

        /// <summary>
        /// Gets the namespace URI.
        /// </summary>
        /// <value>The namespace URI.</value>
        public string NamespaceUri
            {
            get { return xmlReader.NamespaceURI; }
            }

        /// <summary>
        /// Gets the type of the node.
        /// </summary>
        /// <value>The type of the node.</value>
        public XmlNodeType NodeType
            {
            get { return xmlReader.NodeType; }
            }

        /// <summary>
        /// Gets the type of the prev node.
        /// </summary>
        /// <value>The type of the prev node.</value>
        public XmlNodeType PrevNodeType
            {
            get { return prevNodeType; }
            }
        }
    }