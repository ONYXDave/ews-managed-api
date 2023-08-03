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
    /// Represents the CompanyInsightValue.
    /// </summary>
    public sealed class CompanyInsightValue : InsightValue
        {
        private string name;
        private string satoriId;
        private string description;
        private string descriptionAttribution;
        private string imageUrl;
        private string imageUrlAttribution;
        private string yearFound;
        private string financeSymbol;
        private string websiteUrl;

        /// <summary>
        /// Gets the Name
        /// </summary>
        public string Name
            {
            get
                {
                return name;
                }

            set
                {
                SetFieldValue<string>(ref name, value);
                }
            }

        /// <summary>
        /// Gets the SatoriId
        /// </summary>
        public string SatoriId
            {
            get
                {
                return satoriId;
                }

            set
                {
                SetFieldValue<string>(ref satoriId, value);
                }
            }

        /// <summary>
        /// Gets the Description
        /// </summary>
        public string Description
            {
            get
                {
                return description;
                }

            set
                {
                SetFieldValue<string>(ref description, value);
                }
            }

        /// <summary>
        /// Gets the DescriptionAttribution
        /// </summary>
        public string DescriptionAttribution
            {
            get
                {
                return descriptionAttribution;
                }

            set
                {
                SetFieldValue<string>(ref descriptionAttribution, value);
                }
            }

        /// <summary>
        /// Gets the ImageUrl
        /// </summary>
        public string ImageUrl
            {
            get
                {
                return imageUrl;
                }

            set
                {
                SetFieldValue<string>(ref imageUrl, value);
                }
            }

        /// <summary>
        /// Gets the ImageUrlAttribution
        /// </summary>
        public string ImageUrlAttribution
            {
            get
                {
                return imageUrlAttribution;
                }

            set
                {
                SetFieldValue<string>(ref imageUrlAttribution, value);
                }
            }

        /// <summary>
        /// Gets the YearFound
        /// </summary>
        public string YearFound
            {
            get
                {
                return yearFound;
                }

            set
                {
                SetFieldValue<string>(ref yearFound, value);
                }
            }

        /// <summary>
        /// Gets the FinanceSymbol
        /// </summary>
        public string FinanceSymbol
            {
            get
                {
                return financeSymbol;
                }

            set
                {
                SetFieldValue<string>(ref financeSymbol, value);
                }
            }

        /// <summary>
        /// Gets the WebsiteUrl
        /// </summary>
        public string WebsiteUrl
            {
            get
                {
                return websiteUrl;
                }

            set
                {
                SetFieldValue<string>(ref websiteUrl, value);
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
                case XmlElementNames.Name:
                    Name = reader.ReadElementValue();
                    break;
                case XmlElementNames.SatoriId:
                    SatoriId = reader.ReadElementValue();
                    break;
                case XmlElementNames.Description:
                    Description = reader.ReadElementValue();
                    break;
                case XmlElementNames.DescriptionAttribution:
                    DescriptionAttribution = reader.ReadElementValue();
                    break;
                case XmlElementNames.ImageUrl:
                    ImageUrl = reader.ReadElementValue();
                    break;
                case XmlElementNames.ImageUrlAttribution:
                    ImageUrlAttribution = reader.ReadElementValue();
                    break;
                case XmlElementNames.YearFound:
                    YearFound = reader.ReadElementValue();
                    break;
                case XmlElementNames.FinanceSymbol:
                    FinanceSymbol = reader.ReadElementValue();
                    break;
                case XmlElementNames.WebsiteUrl:
                    WebsiteUrl = reader.ReadElementValue();
                    break;
                default:
                    return false;
                }

            return true;
            }
        }
    }