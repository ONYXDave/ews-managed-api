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
    using System.Collections.ObjectModel;
    using System.Xml;

    /// <summary>
    /// Represents the response to a GetDomainSettings call for an individual domain.
    /// </summary>
    public sealed class GetDomainSettingsResponse : AutodiscoverResponse
        {
        private string domain;
        private string redirectTarget;
        private Dictionary<DomainSettingName, object> settings;
        private Collection<DomainSettingError> domainSettingErrors;

        /// <summary>
        /// Initializes a new instance of the <see cref="GetDomainSettingsResponse"/> class.
        /// </summary>
        public GetDomainSettingsResponse()
            : base()
            {
            domain = string.Empty;
            settings = new Dictionary<DomainSettingName, object>();
            domainSettingErrors = new Collection<DomainSettingError>();
            }

        /// <summary>
        /// Gets the domain this response applies to.
        /// </summary>
        public string Domain
            {
            get { return domain; }
            internal set { domain = value; }
            }

        /// <summary>
        /// Gets the redirectionTarget (URL or email address)
        /// </summary>
        public string RedirectTarget
            {
            get { return redirectTarget; }
            }

        /// <summary>
        /// Gets the requested settings for the domain.
        /// </summary>
        public IDictionary<DomainSettingName, object> Settings
            {
            get { return settings; }
            }

        /// <summary>
        /// Gets error information for settings that could not be returned.
        /// </summary>
        public Collection<DomainSettingError> DomainSettingErrors
            {
            get { return domainSettingErrors; }
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
                    switch (reader.LocalName)
                        {
                        case XmlElementNames.RedirectTarget:
                            redirectTarget = reader.ReadElementValue();
                            break;
                        case XmlElementNames.DomainSettingErrors:
                            LoadDomainSettingErrorsFromXml(reader);
                            break;
                        case XmlElementNames.DomainSettings:
                            LoadDomainSettingsFromXml(reader);
                            break;
                        default:
                            base.LoadFromXml(reader, endElementName);
                            break;
                        }
                    }
                }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, endElementName));
            }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadDomainSettingsFromXml(EwsXmlReader reader)
            {
            if (!reader.IsEmptyElement)
                {
                do
                    {
                    reader.Read();

                    if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.DomainSetting))
                        {
                        string settingClass = reader.ReadAttributeValue(XmlNamespace.XmlSchemaInstance, XmlAttributeNames.Type);

                        switch (settingClass)
                            {
                            case XmlElementNames.DomainStringSetting:
                                ReadSettingFromXml(reader);
                                break;

                            default:
                                EwsUtilities.Assert(
                                    false,
                                    "GetDomainSettingsResponse.LoadDomainSettingsFromXml",
                                    string.Format("Invalid setting class '{0}' returned", settingClass));
                                break;
                            }
                        }
                    }
                while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DomainSettings));
                }
            }

        /// <summary>
        /// Reads domain setting from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ReadSettingFromXml(EwsXmlReader reader)
            {
            DomainSettingName? name = null;
            object value = null;

            do
                {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                    {
                    switch (reader.LocalName)
                        {
                        case XmlElementNames.Name:
                            name = reader.ReadElementValue<DomainSettingName>();
                            break;
                        case XmlElementNames.Value:
                            value = reader.ReadElementValue();
                            break;
                        }
                    }
                }
            while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DomainSetting));

            EwsUtilities.Assert(
                name.HasValue,
                "GetDomainSettingsResponse.ReadSettingFromXml",
                "Missing name element in domain setting");

            settings.Add(name.Value, value);
            }

        /// <summary>
        /// Loads the domain setting errors.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void LoadDomainSettingErrorsFromXml(EwsXmlReader reader)
            {
            if (!reader.IsEmptyElement)
                {
                do
                    {
                    reader.Read();

                    if ((reader.NodeType == XmlNodeType.Element) && (reader.LocalName == XmlElementNames.DomainSettingError))
                        {
                        DomainSettingError error = new();
                        error.LoadFromXml(reader);
                        domainSettingErrors.Add(error);
                        }
                    }
                while (!reader.IsEndElement(XmlNamespace.Autodiscover, XmlElementNames.DomainSettingErrors));
                }
            }
        }
    }