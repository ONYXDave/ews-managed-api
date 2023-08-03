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
    /// Represents the SingleValueInsightContent.
    /// </summary>
    public sealed class SingleValueInsightContent : ComplexProperty
        {
        /// <summary>
        /// Gets the Item
        /// </summary>
        public InsightValue Item { get; internal set; }

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
            switch (reader.ReadAttributeValue("xsi:type"))
                {
                case XmlElementNames.StringInsightValue:
                    Item = new StringInsightValue();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.ProfileInsightValue:
                    Item = new ProfileInsightValue();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.JobInsightValue:
                    Item = new JobInsightValue();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.UserProfilePicture:
                    Item = new UserProfilePicture();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.EducationInsightValue:
                    Item = new EducationInsightValue();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.SkillInsightValue:
                    Item = new SkillInsightValue();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.DelveDocument:
                    Item = new DelveDocument();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.CompanyInsightValue:
                    Item = new CompanyInsightValue();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.ComputedInsightValue:
                    Item = new ComputedInsightValue();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                case XmlElementNames.OutOfOfficeInsightValue:
                    Item = new OutOfOfficeInsightValue();
                    Item.LoadFromXml(reader, reader.LocalName);
                    break;
                default:
                    return false;
                }

            return true;
            }
        }
    }