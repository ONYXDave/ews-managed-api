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
    /// Represents the EducationInsightValue.
    /// </summary>
    public sealed class EducationInsightValue : InsightValue
        {
        private string institute;
        private string degree;
        private long startUtcTicks;
        private long endUtcTicks;

        /// <summary>
        /// Gets the Institute
        /// </summary>
        public string Institute
            {
            get
                {
                return institute;
                }

            set
                {
                SetFieldValue<string>(ref institute, value);
                }
            }

        /// <summary>
        /// Gets the Degree
        /// </summary>
        public string Degree
            {
            get
                {
                return degree;
                }

            set
                {
                SetFieldValue<string>(ref degree, value);
                }
            }

        /// <summary>
        /// Gets the StartUtcTicks
        /// </summary>
        public long StartUtcTicks
            {
            get
                {
                return startUtcTicks;
                }

            set
                {
                SetFieldValue<long>(ref startUtcTicks, value);
                }
            }

        /// <summary>
        /// Gets the EndUtcTicks
        /// </summary>
        public long EndUtcTicks
            {
            get
                {
                return endUtcTicks;
                }

            set
                {
                SetFieldValue<long>(ref endUtcTicks, value);
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
                case XmlElementNames.Institute:
                    Institute = reader.ReadElementValue();
                    break;
                case XmlElementNames.Degree:
                    Degree = reader.ReadElementValue();
                    break;
                case XmlElementNames.StartUtcTicks:
                    StartUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.EndUtcTicks:
                    EndUtcTicks = reader.ReadElementValue<long>();
                    break;
                default:
                    return false;
                }

            return true;
            }
        }
    }