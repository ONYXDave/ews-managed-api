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
    /// Represents the minimum and maximum size of a message.
    /// </summary>
    public sealed class RulePredicateSizeRange : ComplexProperty
        {
        /// <summary>
        /// Minimum Size.
        /// </summary>
        private int? minimumSize;

        /// <summary>
        /// Mamixmum Size.
        /// </summary>
        private int? maximumSize;

        /// <summary>
        /// Initializes a new instance of the <see cref="RulePredicateSizeRange"/> class.
        /// </summary>
        internal RulePredicateSizeRange()
            : base()
            {
            }

        /// <summary>
        /// Gets or sets the minimum size, in kilobytes. If MinimumSize is set to 
        /// null, no minimum size applies.
        /// </summary>
        public int? MinimumSize
            {
            get
                {
                return minimumSize;
                }

            set
                {
                SetFieldValue<int?>(ref minimumSize, value);
                }
            }

        /// <summary>
        /// Gets or sets the maximum size, in kilobytes. If MaximumSize is set to 
        /// null, no maximum size applies.
        /// </summary>
        public int? MaximumSize
            {
            get
                {
                return maximumSize;
                }

            set
                {
                SetFieldValue<int?>(ref maximumSize, value);
                }
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
                case XmlElementNames.MinimumSize:
                    minimumSize = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.MaximumSize:
                    maximumSize = reader.ReadElementValue<int>();
                    return true;
                default:
                    return false;
                }
            }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
            if (MinimumSize.HasValue)
                {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MinimumSize, MinimumSize.Value);
                }
            if (MaximumSize.HasValue)
                {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MaximumSize, MaximumSize.Value);
                }
            }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
            {
            base.InternalValidate();
            if (minimumSize.HasValue &&
                maximumSize.HasValue &&
                minimumSize.Value > maximumSize.Value)
                {
                throw new ServiceValidationException("MinimumSize cannot be larger than MaximumSize.");
                }
            }
        }
    }