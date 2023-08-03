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
    using System.Xml;

    /// <summary>
    /// Represents PersonaPostalAddress.
    /// </summary>
    public sealed class PersonaPostalAddress : ComplexProperty
        {
        private string street;
        private string city;
        private string state;
        private string country;
        private string postalCode;
        private string postOfficeBox;
        private string type;
        private double? latitude;
        private double? longitude;
        private double? accuracy;
        private double? altitude;
        private double? altitudeAccuracy;
        private string formattedAddress;
        private string uri;
        private LocationSource source;

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonaPostalAddress"/> class.
        /// </summary>
        internal PersonaPostalAddress()
            : base()
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="PersonaPostalAddress"/> class.
        /// </summary>
        /// <param name="street">The Street Address.</param>
        /// <param name="city">The City value.</param>
        /// <param name="state">The State value.</param>
        /// <param name="country">The country value.</param>
        /// <param name="postalCode">The postal code value.</param>
        /// <param name="postOfficeBox">The Post Office Box.</param>
        /// <param name="locationSource">The location Source.</param>
        /// <param name="locationUri">The location Uri.</param>
        /// <param name="formattedAddress">The location street Address in formatted address.</param>
        /// <param name="latitude">The location latitude.</param>
        /// <param name="longitude">The location longitude.</param>
        /// <param name="accuracy">The location accuracy.</param>
        /// <param name="altitude">The location altitude.</param>
        /// <param name="altitudeAccuracy">The location altitude Accuracy.</param>
        public PersonaPostalAddress(
            string street,
            string city,
            string state,
            string country,
            string postalCode,
            string postOfficeBox,
            LocationSource locationSource,
            string locationUri,
            string formattedAddress,
            double latitude,
            double longitude,
            double accuracy,
            double altitude,
            double altitudeAccuracy)
            : this()
            {
            this.street = street;
            this.city = city;
            this.state = state;
            this.country = country;
            this.postalCode = postalCode;
            this.postOfficeBox = postOfficeBox;
            this.latitude = latitude;
            this.longitude = longitude;
            source = locationSource;
            uri = locationUri;
            this.formattedAddress = formattedAddress;
            this.accuracy = accuracy;
            this.altitude = altitude;
            this.altitudeAccuracy = altitudeAccuracy;
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
                case XmlElementNames.Street:
                    street = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.City:
                    city = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.State:
                    state = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.Country:
                    country = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.PostalCode:
                    postalCode = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.PostOfficeBox:
                    postOfficeBox = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.PostalAddressType:
                    type = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.Latitude:
                    latitude = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.Longitude:
                    longitude = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.Accuracy:
                    accuracy = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.Altitude:
                    altitude = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.AltitudeAccuracy:
                    altitudeAccuracy = reader.ReadValue<double>();
                    return true;
                case XmlElementNames.FormattedAddress:
                    formattedAddress = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.LocationUri:
                    uri = reader.ReadValue<string>();
                    return true;
                case XmlElementNames.LocationSource:
                    source = reader.ReadValue<LocationSource>();
                    return true;
                default:
                    return false;
                }
            }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsServiceXmlReader reader)
            {
            do
                {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                    {
                    TryReadElementFromXml(reader);
                    }
                }
            while (!reader.IsEndElement(XmlNamespace.Types, XmlElementNames.PersonaPostalAddress));
            }

        /// <summary>
        /// Gets or sets the street.
        /// </summary>
        public string Street
            {
            get { return street; }
            set { SetFieldValue<string>(ref street, value); }
            }

        /// <summary>
        /// Gets or sets the City.
        /// </summary>
        public string City
            {
            get { return city; }
            set { SetFieldValue<string>(ref city, value); }
            }

        /// <summary>
        /// Gets or sets the state.
        /// </summary>
        public string State
            {
            get { return state; }
            set { SetFieldValue<string>(ref state, value); }
            }

        /// <summary>
        /// Gets or sets the Country.
        /// </summary>
        public string Country
            {
            get { return country; }
            set { SetFieldValue<string>(ref country, value); }
            }

        /// <summary>
        /// Gets or sets the postalCode.
        /// </summary>
        public string PostalCode
            {
            get { return postalCode; }
            set { SetFieldValue<string>(ref postalCode, value); }
            }

        /// <summary>
        /// Gets or sets the postOfficeBox.
        /// </summary>
        public string PostOfficeBox
            {
            get { return postOfficeBox; }
            set { SetFieldValue<string>(ref postOfficeBox, value); }
            }

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        public string Type
            {
            get { return type; }
            set { SetFieldValue<string>(ref type, value); }
            }

        /// <summary>
        /// Gets or sets the location source type.
        /// </summary>
        public LocationSource Source
            {
            get { return source; }
            set { SetFieldValue<LocationSource>(ref source, value); }
            }

        /// <summary>
        /// Gets or sets the location Uri.
        /// </summary>
        public string Uri
            {
            get { return uri; }
            set { SetFieldValue<string>(ref uri, value); }
            }

        /// <summary>
        /// Gets or sets a value indicating location latitude.
        /// </summary>
        public double? Latitude
            {
            get { return latitude; }
            set { SetFieldValue<double?>(ref latitude, value); }
            }

        /// <summary>
        /// Gets or sets a value indicating location longitude.
        /// </summary>
        public double? Longitude
            {
            get { return longitude; }
            set { SetFieldValue<double?>(ref longitude, value); }
            }

        /// <summary>
        /// Gets or sets the location accuracy.
        /// </summary>
        public double? Accuracy
            {
            get { return accuracy; }
            set { SetFieldValue<double?>(ref accuracy, value); }
            }

        /// <summary>
        /// Gets or sets the location altitude.
        /// </summary>
        public double? Altitude
            {
            get { return altitude; }
            set { SetFieldValue<double?>(ref altitude, value); }
            }

        /// <summary>
        /// Gets or sets the location altitude accuracy.
        /// </summary>
        public double? AltitudeAccuracy
            {
            get { return altitudeAccuracy; }
            set { SetFieldValue<double?>(ref altitudeAccuracy, value); }
            }

        /// <summary>
        /// Gets or sets the street address.
        /// </summary>
        public string FormattedAddress
            {
            get { return formattedAddress; }
            set { SetFieldValue<string>(ref formattedAddress, value); }
            }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Street, street);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.City, city);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.State, state);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Country, country);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostalCode, postalCode);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostOfficeBox, postOfficeBox);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.PostalAddressType, type);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Latitude, latitude);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Longitude, longitude);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Accuracy, accuracy);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.Altitude, altitude);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.AltitudeAccuracy, altitudeAccuracy);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.FormattedAddress, formattedAddress);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationUri, uri);
            writer.WriteElementValue(XmlNamespace.Types, XmlElementNames.LocationSource, source);
            }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
            {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.PersonaPostalAddress);

            WriteElementsToXml(writer);

            writer.WriteEndElement(); // xmlElementName
            }
        }
    }