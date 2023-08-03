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

    /// <summary>
    /// Encapsulates information on the occurrence of a recurring appointment.
    /// </summary>
    public sealed class Flag : ComplexProperty
        {
        private ItemFlagStatus flagStatus;
        private DateTime startDate;
        private DateTime dueDate;
        private DateTime completeDate;

        /// <summary>
        /// Initializes a new instance of the <see cref="Flag"/> class.
        /// </summary>
        public Flag()
            {
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
                case XmlElementNames.FlagStatus:
                    flagStatus = reader.ReadElementValue<ItemFlagStatus>();
                    return true;
                case XmlElementNames.StartDate:
                    startDate = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.DueDate:
                    dueDate = reader.ReadElementValueAsDateTime().Value;
                    return true;
                case XmlElementNames.CompleteDate:
                    completeDate = reader.ReadElementValueAsDateTime().Value;
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
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FlagStatus, FlagStatus);

            if (FlagStatus == ItemFlagStatus.Flagged)
                {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.StartDate, StartDate);
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.DueDate, DueDate);
                }
            else if (FlagStatus == ItemFlagStatus.Complete)
                {
                writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.CompleteDate, CompleteDate);
                }
            }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal void Validate()
            {
            EwsUtilities.ValidateParam(flagStatus, "FlagStatus");
            }

        /// <summary>
        /// Gets or sets the flag status.
        /// </summary>
        public ItemFlagStatus FlagStatus
            {
            get
                {
                return flagStatus;
                }

            set
                {
                SetFieldValue<ItemFlagStatus>(ref flagStatus, value);
                }
            }

        /// <summary>
        /// Gets the start date.
        /// </summary>
        public DateTime StartDate
            {
            get
                {
                return startDate;
                }

            set
                {
                SetFieldValue<DateTime>(ref startDate, value);
                }
            }

        /// <summary>
        /// Gets the due date.
        /// </summary>
        public DateTime DueDate
            {
            get
                {
                return dueDate;
                }

            set
                {
                SetFieldValue<DateTime>(ref dueDate, value);
                }
            }

        /// <summary>
        /// Gets the complete date.
        /// </summary>
        public DateTime CompleteDate
            {
            get
                {
                return completeDate;
                }

            set
                {
                SetFieldValue<DateTime>(ref completeDate, value);
                }
            }
        }
    }