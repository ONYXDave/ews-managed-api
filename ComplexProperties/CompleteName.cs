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
    /// Represents the complete name of a contact.
    /// </summary>
    public sealed class CompleteName : ComplexProperty
        {
        private string title;
        private string givenName;
        private string middleName;
        private string surname;
        private string suffix;
        private string initials;
        private string fullName;
        private string nickname;
        private string yomiGivenName;
        private string yomiSurname;

        #region Properties

        /// <summary>
        /// Gets the contact's title.
        /// </summary>
        public string Title
            {
            get { return title; }
            }

        /// <summary>
        /// Gets the given name (first name) of the contact.
        /// </summary>
        public string GivenName
            {
            get { return givenName; }
            }

        /// <summary>
        /// Gets the middle name of the contact.
        /// </summary>
        public string MiddleName
            {
            get { return middleName; }
            }

        /// <summary>
        /// Gets the surname (last name) of the contact.
        /// </summary>
        public string Surname
            {
            get { return surname; }
            }

        /// <summary>
        /// Gets the suffix of the contact.
        /// </summary>
        public string Suffix
            {
            get { return suffix; }
            }

        /// <summary>
        /// Gets the initials of the contact.
        /// </summary>
        public string Initials
            {
            get { return initials; }
            }

        /// <summary>
        /// Gets the full name of the contact.
        /// </summary>
        public string FullName
            {
            get { return fullName; }
            }

        /// <summary>
        /// Gets the nickname of the contact.
        /// </summary>
        public string NickName
            {
            get { return nickname; }
            }

        /// <summary>
        /// Gets the Yomi given name (first name) of the contact.
        /// </summary>
        public string YomiGivenName
            {
            get { return yomiGivenName; }
            }

        /// <summary>
        /// Gets the Yomi surname (last name) of the contact.
        /// </summary>
        public string YomiSurname
            {
            get { return yomiSurname; }
            }

        #endregion

        /// <summary>
        /// Tries to read element from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>True if element was read.</returns>
        internal override bool TryReadElementFromXml(EwsServiceXmlReader reader)
            {
            switch (reader.LocalName)
                {
                case XmlElementNames.Title:
                    title = reader.ReadElementValue();
                    return true;
                case XmlElementNames.FirstName:
                    givenName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.MiddleName:
                    middleName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.LastName:
                    surname = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Suffix:
                    suffix = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Initials:
                    initials = reader.ReadElementValue();
                    return true;
                case XmlElementNames.FullName:
                    fullName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.NickName:
                    nickname = reader.ReadElementValue();
                    return true;
                case XmlElementNames.YomiFirstName:
                    yomiGivenName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.YomiLastName:
                    yomiSurname = reader.ReadElementValue();
                    return true;
                default:
                    return false;
                }
            }

        /// <summary>
        /// Writes the elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Title, Title);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FirstName, GivenName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.MiddleName, MiddleName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LastName, Surname);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Suffix, Suffix);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Initials, Initials);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FullName, FullName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.NickName, NickName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.YomiFirstName, YomiGivenName);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.YomiLastName, YomiSurname);
            }
        }
    }