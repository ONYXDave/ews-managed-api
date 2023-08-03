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
    /// Represents the ProfileInsightValue.
    /// </summary>
    public sealed class ProfileInsightValue : InsightValue
        {
        private string fullName;
        private string firstName;
        private string lastName;
        private string emailAddress;
        private string avatar;
        private long joinedUtcTicks;
        private UserProfilePicture profilePicture;
        private string title;

        /// <summary>
        /// Gets the FullName
        /// </summary>
        public string FullName
            {
            get
                {
                return fullName;
                }
            }

        /// <summary>
        /// Gets the FirstName
        /// </summary>
        public string FirstName
            {
            get
                {
                return firstName;
                }
            }

        /// <summary>
        /// Gets the LastName
        /// </summary>
        public string LastName
            {
            get
                {
                return lastName;
                }
            }

        /// <summary>
        /// Gets the EmailAddress
        /// </summary>
        public string EmailAddress
            {
            get
                {
                return emailAddress;
                }
            }

        /// <summary>
        /// Gets the Avatar
        /// </summary>
        public string Avatar
            {
            get
                {
                return avatar;
                }
            }

        /// <summary>
        /// Gets the JoinedUtcTicks
        /// </summary>
        public long JoinedUtcTicks
            {
            get
                {
                return joinedUtcTicks;
                }
            }

        /// <summary>
        /// Gets the ProfilePicture
        /// </summary>
        public UserProfilePicture ProfilePicture
            {
            get
                {
                return profilePicture;
                }
            }

        /// <summary>
        /// Gets the Title
        /// </summary>
        public string Title
            {
            get
                {
                return title;
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
                case XmlElementNames.FullName:
                    fullName = reader.ReadElementValue();
                    break;
                case XmlElementNames.FirstName:
                    firstName = reader.ReadElementValue();
                    break;
                case XmlElementNames.LastName:
                    lastName = reader.ReadElementValue();
                    break;
                case XmlElementNames.EmailAddress:
                    emailAddress = reader.ReadElementValue();
                    break;
                case XmlElementNames.Avatar:
                    avatar = reader.ReadElementValue();
                    break;
                case XmlElementNames.JoinedUtcTicks:
                    joinedUtcTicks = reader.ReadElementValue<long>();
                    break;
                case XmlElementNames.ProfilePicture:
                    UserProfilePicture picture = new();
                    picture.LoadFromXml(reader, XmlNamespace.Types, XmlElementNames.ProfilePicture);
                    profilePicture = picture;
                    break;
                case XmlElementNames.Title:
                    title = reader.ReadElementValue();
                    break;
                default:
                    return false;
                }

            return true;
            }
        }
    }