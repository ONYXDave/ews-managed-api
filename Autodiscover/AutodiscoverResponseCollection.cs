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

namespace Microsoft.Exchange.WebServices.Autodiscover
    {
    using Microsoft.Exchange.WebServices.Data;
    using System.Collections.Generic;
    using System.Xml;

    /// <summary>
    /// Represents a collection of responses to a call to the Autodiscover service.
    /// </summary>
    /// <typeparam name="TResponse">The type of the responses in the collection.</typeparam>
    public abstract class AutodiscoverResponseCollection<TResponse> : AutodiscoverResponse, IEnumerable<TResponse>
        where TResponse : AutodiscoverResponse
        {
        private List<TResponse> responses;

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverResponseCollection&lt;TResponse&gt;"/> class.
        /// </summary>
        internal AutodiscoverResponseCollection()
            {
            responses = new List<TResponse>();
            }

        /// <summary>
        /// Gets the number of responses in the collection.
        /// </summary>
        public int Count
            {
            get { return responses.Count; }
            }

        /// <summary>
        /// Gets the response at the specified index.
        /// </summary>
        /// <param name="index">Index.</param>
        public TResponse this[int index]
            {
            get { return responses[index]; }
            }

        /// <summary>
        /// Gets the responses list.
        /// </summary>
        internal List<TResponse> Responses
            {
            get { return responses; }
            }

        /// <summary>
        /// Loads response from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="endElementName">End element name.</param>
        internal override void LoadFromXml(EwsXmlReader reader, string endElementName)
            {
            do
                {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                    {
                    if (reader.LocalName == GetResponseCollectionXmlElementName())
                        {
                        LoadResponseCollectionFromXml(reader);
                        }
                    else
                        {
                        base.LoadFromXml(reader, endElementName);
                        }
                    }
                }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, endElementName));
            }

        /// <summary>
        /// Loads the response collection from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void LoadResponseCollectionFromXml(EwsXmlReader reader)
            {
            if (!reader.IsEmptyElement)
                {
                do
                    {
                    reader.Read();
                    if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == GetResponseInstanceXmlElementName()))
                        {
                        TResponse response = CreateResponseInstance();
                        response.LoadFromXml(reader, GetResponseInstanceXmlElementName());
                        Responses.Add(response);
                        }
                    }
                while (!reader.IsEndElement(XmlNamespace.Autodiscover, GetResponseCollectionXmlElementName()));
                }
            }

        /// <summary>
        /// Gets the name of the response collection XML element.
        /// </summary>
        /// <returns>Response collection XMl element name.</returns>
        internal abstract string GetResponseCollectionXmlElementName();

        /// <summary>
        /// Gets the name of the response instance XML element.
        /// </summary>
        /// <returns>Response instance XMl element name.</returns>
        internal abstract string GetResponseInstanceXmlElementName();

        /// <summary>
        /// Create a response instance.
        /// </summary>
        /// <returns>TResponse.</returns>
        internal abstract TResponse CreateResponseInstance();

        #region IEnumerable<TResponse>

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        public IEnumerator<TResponse> GetEnumerator()
            {
            return responses.GetEnumerator();
            }

        #endregion

        #region IEnumerable Members

        /// <summary>
        /// Gets an enumerator that iterates through the elements of the collection.
        /// </summary>
        /// <returns>An IEnumerator for the collection.</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
            return (responses as System.Collections.IEnumerable).GetEnumerator();
            }

        #endregion
        }
    }