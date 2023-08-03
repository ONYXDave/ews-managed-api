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
    using System.Linq;

    /// <summary>
    /// Represents an abstract Subscribe request.
    /// </summary>
    /// <typeparam name="TSubscription">The type of the subscription.</typeparam>
    internal abstract class SubscribeRequest<TSubscription> : MultiResponseServiceRequest<SubscribeResponse<TSubscription>>
        where TSubscription : SubscriptionBase
        {
        /// <summary>
        /// Validate request.
        /// </summary>
        internal override void Validate()
            {
            base.Validate();
            EwsUtilities.ValidateParam(FolderIds, "FolderIds");
            EwsUtilities.ValidateParamCollection(EventTypes, "EventTypes");
            FolderIds.Validate(Service.RequestedServerVersion);

            // Check that caller isn't trying to subscribe to Status events.
            if (EventTypes.Count<EventType>(eventType => eventType == EventType.Status) > 0)
                {
                throw new ServiceValidationException(Strings.CannotSubscribeToStatusEvents);
                }

            // If Watermark was specified, make sure it's not a blank string.
            if (!string.IsNullOrEmpty(Watermark))
                {
                EwsUtilities.ValidateNonBlankStringParam(Watermark, "Watermark");
                }

            EventTypes.ForEach(eventType => EwsUtilities.ValidateEnumVersionValue(eventType, Service.RequestedServerVersion));
            }

        /// <summary>
        /// Gets the name of the subscription XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal abstract string GetSubscriptionXmlElementName();

        /// <summary>
        /// Gets the expected response message count.
        /// </summary>
        /// <returns>Number of expected response messages.</returns>
        internal override int GetExpectedResponseMessageCount()
            {
            return 1;
            }

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetXmlElementName()
            {
            return XmlElementNames.Subscribe;
            }

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseXmlElementName()
            {
            return XmlElementNames.SubscribeResponse;
            }

        /// <summary>
        /// Gets the name of the response message XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal override string GetResponseMessageXmlElementName()
            {
            return XmlElementNames.SubscribeResponseMessage;
            }

        /// <summary>
        /// Internal method to write XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void InternalWriteElementsToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
            writer.WriteStartElement(XmlNamespace.Messages, GetSubscriptionXmlElementName());

            if (FolderIds.Count == 0)
                {
                writer.WriteAttributeValue(
                    XmlAttributeNames.SubscribeToAllFolders,
                    true);
                }

            FolderIds.WriteToXml(
                writer,
                XmlNamespace.Types,
                XmlElementNames.FolderIds);

            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.EventTypes);
            foreach (EventType eventType in EventTypes)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.EventType,
                    eventType);
                }
            writer.WriteEndElement();

            if (!string.IsNullOrEmpty(Watermark))
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Watermark,
                    Watermark);
                }

            InternalWriteElementsToXml(writer);

            writer.WriteEndElement();
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="SubscribeRequest&lt;TSubscription&gt;"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal SubscribeRequest(ExchangeService service)
            : base(service, ServiceErrorHandling.ThrowOnError)
            {
            FolderIds = new FolderIdWrapperList();
            EventTypes = new List<EventType>();
            }

        /// <summary>
        /// Gets the folder ids.
        /// </summary>
        public FolderIdWrapperList FolderIds
            {
            get; private set;
            }

        /// <summary>
        /// Gets the event types.
        /// </summary>
        public List<EventType> EventTypes
            {
            get; private set;
            }

        /// <summary>
        /// Gets or sets the watermark.
        /// </summary>
        public string Watermark
            {
            get; set;
            }
        }
    }