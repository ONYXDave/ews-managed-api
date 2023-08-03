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
    using System.Text;

    /// <summary>
    /// Represents an extended property.
    /// </summary>
    public sealed class ExtendedProperty : ComplexProperty
        {
        private ExtendedPropertyDefinition propertyDefinition;
        private object value;

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        internal ExtendedProperty()
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExtendedProperty"/> class.
        /// </summary>
        /// <param name="propertyDefinition">The definition of the extended property.</param>
        internal ExtendedProperty(ExtendedPropertyDefinition propertyDefinition)
            : this()
            {
            EwsUtilities.ValidateParam(propertyDefinition, "propertyDefinition");

            this.propertyDefinition = propertyDefinition;
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
                case XmlElementNames.ExtendedFieldURI:
                    propertyDefinition = new ExtendedPropertyDefinition();
                    propertyDefinition.LoadFromXml(reader);
                    return true;
                case XmlElementNames.Value:
                    EwsUtilities.Assert(
                        PropertyDefinition != null,
                        "ExtendedProperty.TryReadElementFromXml",
                        "PropertyDefintion is missing");

                    string stringValue = reader.ReadElementValue();
                    value = MapiTypeConverter.ConvertToValue(PropertyDefinition.MapiType, stringValue);
                    return true;
                case XmlElementNames.Values:
                    EwsUtilities.Assert(
                        PropertyDefinition != null,
                        "ExtendedProperty.TryReadElementFromXml",
                        "PropertyDefintion is missing");

                    StringList stringList = new(XmlElementNames.Value);
                    stringList.LoadFromXml(reader, reader.LocalName);
                    value = MapiTypeConverter.ConvertToValue(PropertyDefinition.MapiType, stringList);
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
            PropertyDefinition.WriteToXml(writer);

            if (MapiTypeConverter.IsArrayType(PropertyDefinition.MapiType))
                {
                Array array = Value as Array;
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.Values);
                for (int index = array.GetLowerBound(0); index <= array.GetUpperBound(0); index++)
                    {
                    writer.WriteElementValue(
                        XmlNamespace.Types,
                        XmlElementNames.Value,
                        MapiTypeConverter.ConvertToString(PropertyDefinition.MapiType, array.GetValue(index)));
                    }
                writer.WriteEndElement();
                }
            else
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Value,
                    MapiTypeConverter.ConvertToString(PropertyDefinition.MapiType, Value));
                }
            }

        /// <summary>
        /// Gets the definition of the extended property.
        /// </summary>
        public ExtendedPropertyDefinition PropertyDefinition
            {
            get { return propertyDefinition; }
            }

        /// <summary>
        /// Gets or sets the value of the extended property.
        /// </summary>
        public object Value
            {
            get
                {
                return value;
                }

            set
                {
                EwsUtilities.ValidateParam(value, "value");
                SetFieldValue<object>(
                    ref this.value,
                    MapiTypeConverter.ChangeType(PropertyDefinition.MapiType, value));
                }
            }

        /// <summary>
        /// Gets the string value.
        /// </summary>
        /// <returns>Value as string.</returns>
        private string GetStringValue()
            {
            if (MapiTypeConverter.IsArrayType(PropertyDefinition.MapiType))
                {
                Array array = Value as Array;
                if (array == null)
                    {
                    return string.Empty;
                    }
                else
                    {
                    StringBuilder sb = new();
                    sb.Append("[");
                    for (int index = array.GetLowerBound(0); index <= array.GetUpperBound(0); index++)
                        {
                        sb.Append(
                            MapiTypeConverter.ConvertToString(
                                PropertyDefinition.MapiType,
                                array.GetValue(index)));
                        sb.Append(",");
                        }
                    sb.Append("]");

                    return sb.ToString();
                    }
                }
            else
                {
                return MapiTypeConverter.ConvertToString(PropertyDefinition.MapiType, Value);
                }
            }

        /// <summary>
        /// Determines whether the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>.
        /// </summary>
        /// <param name="obj">The <see cref="T:System.Object"/> to compare with the current <see cref="T:System.Object"/>.</param>
        /// <returns>
        /// true if the specified <see cref="T:System.Object"/> is equal to the current <see cref="T:System.Object"/>; otherwise, false.
        /// </returns>
        /// <exception cref="T:System.NullReferenceException">The <paramref name="obj"/> parameter is null.</exception>
        public override bool Equals(object obj)
            {
            ExtendedProperty other = obj as ExtendedProperty;
            if ((other != null) && other.PropertyDefinition.Equals(PropertyDefinition))
                {
                return GetStringValue().Equals(other.GetStringValue());
                }
            else
                {
                return false;
                }
            }

        /// <summary>
        /// Serves as a hash function for a particular type.
        /// </summary>
        /// <returns>
        /// A hash code for the current <see cref="T:System.Object"/>.
        /// </returns>
        public override int GetHashCode()
            {
            return string.Concat(
                (PropertyDefinition != null) ? PropertyDefinition.GetPrintableName() : string.Empty,
                GetStringValue()).GetHashCode();
            }
        }
    }