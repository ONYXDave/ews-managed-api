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
    using System.Collections.Generic;
    using System.ComponentModel;

    /// <summary>
    /// Represents a generic dictionary that can be sent to or retrieved from EWS.
    /// </summary>
    /// <typeparam name="TKey">The type of key.</typeparam>
    /// <typeparam name="TEntry">The type of entry.</typeparam>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class DictionaryProperty<TKey, TEntry> : ComplexProperty, ICustomUpdateSerializer
        where TEntry : DictionaryEntryProperty<TKey>
        {
        private Dictionary<TKey, TEntry> entries = new();
        private Dictionary<TKey, TEntry> removedEntries = new();
        private List<TKey> addedEntries = new();
        private List<TKey> modifiedEntries = new();

        /// <summary>
        /// Entry was changed.
        /// </summary>
        /// <param name="complexProperty">The complex property.</param>
        private void EntryChanged(ComplexProperty complexProperty)
            {
            TKey key = (complexProperty as TEntry).Key;

            if (!addedEntries.Contains(key) && !modifiedEntries.Contains(key))
                {
                modifiedEntries.Add(key);
                Changed();
                }
            }

        /// <summary>
        /// Writes the URI to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="key">The key.</param>
        private void WriteUriToXml(EwsServiceXmlWriter writer, TKey key)
            {
            writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.IndexedFieldURI);
            writer.WriteAttributeValue(XmlAttributeNames.FieldURI, GetFieldURI());
            writer.WriteAttributeValue(XmlAttributeNames.FieldIndex, GetFieldIndex(key));
            writer.WriteEndElement();
            }

        /// <summary>
        /// Gets the index of the field.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns>Key index.</returns>
        internal virtual string GetFieldIndex(TKey key)
            {
            return key.ToString();
            }

        /// <summary>
        /// Gets the field URI.
        /// </summary>
        /// <returns>Field URI.</returns>
        internal virtual string GetFieldURI()
            {
            return null;
            }

        /// <summary>
        /// Creates the entry.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Dictionary entry.</returns>
        internal virtual TEntry CreateEntry(EwsServiceXmlReader reader)
            {
            if (reader.LocalName == XmlElementNames.Entry)
                {
                return CreateEntryInstance();
                }
            else
                {
                return null;
                }
            }

        /// <summary>
        /// Creates instance of dictionary entry.
        /// </summary>
        /// <returns>New instance.</returns>
        internal abstract TEntry CreateEntryInstance();

        /// <summary>
        /// Gets the name of the entry XML element.
        /// </summary>
        /// <param name="entry">The entry.</param>
        /// <returns>XML element name.</returns>
        internal virtual string GetEntryXmlElementName(TEntry entry)
            {
            return XmlElementNames.Entry;
            }

        /// <summary>
        /// Clears the change log.
        /// </summary>
        internal override void ClearChangeLog()
            {
            addedEntries.Clear();
            removedEntries.Clear();
            modifiedEntries.Clear();

            foreach (TEntry entry in entries.Values)
                {
                entry.ClearChangeLog();
                }
            }

        /// <summary>
        /// Add entry.
        /// </summary>
        /// <param name="entry">The entry.</param>
        internal void InternalAdd(TEntry entry)
            {
            entry.OnChange += EntryChanged;

            entries.Add(entry.Key, entry);
            addedEntries.Add(entry.Key);
            removedEntries.Remove(entry.Key);

            Changed();
            }

        /// <summary>
        /// Add or replace entry.
        /// </summary>
        /// <param name="entry">The entry.</param>
        internal void InternalAddOrReplace(TEntry entry)
            {
            TEntry oldEntry;

            if (entries.TryGetValue(entry.Key, out oldEntry))
                {
                oldEntry.OnChange -= EntryChanged;

                entry.OnChange += EntryChanged;

                if (!addedEntries.Contains(entry.Key))
                    {
                    if (!modifiedEntries.Contains(entry.Key))
                        {
                        modifiedEntries.Add(entry.Key);
                        }
                    }

                Changed();
                }
            else
                {
                InternalAdd(entry);
                }
            }

        /// <summary>
        /// Remove entry based on key.
        /// </summary>
        /// <param name="key">The key.</param>
        internal void InternalRemove(TKey key)
            {
            TEntry entry;

            if (entries.TryGetValue(key, out entry))
                {
                entry.OnChange -= EntryChanged;

                entries.Remove(key);
                removedEntries.Add(key, entry);

                Changed();
                }

            addedEntries.Remove(key);
            modifiedEntries.Remove(key);
            }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="localElementName">Name of the local element.</param>
        internal override void LoadFromXml(EwsServiceXmlReader reader, string localElementName)
            {
            reader.EnsureCurrentNodeIsStartElement(XmlNamespace.Types, localElementName);

            if (!reader.IsEmptyElement)
                {
                do
                    {
                    reader.Read();

                    if (reader.IsStartElement())
                        {
                        TEntry entry = CreateEntry(reader);

                        if (entry != null)
                            {
                            entry.LoadFromXml(reader, reader.LocalName);
                            InternalAdd(entry);
                            }
                        else
                            {
                            reader.SkipCurrentElement();
                            }
                        }
                    }
                while (!reader.IsEndElement(XmlNamespace.Types, localElementName));
                }
            }

        /// <summary>
        /// Writes to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="xmlNamespace">The XML namespace.</param>
        /// <param name="xmlElementName">Name of the XML element.</param>
        internal override void WriteToXml(
            EwsServiceXmlWriter writer,
            XmlNamespace xmlNamespace,
            string xmlElementName)
            {
            // Only write collection if it has at least one element.
            if (entries.Count > 0)
                {
                base.WriteToXml(
                    writer,
                    xmlNamespace,
                    xmlElementName);
                }
            }

        /// <summary>
        /// Writes elements to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal override void WriteElementsToXml(EwsServiceXmlWriter writer)
            {
            foreach (KeyValuePair<TKey, TEntry> keyValuePair in entries)
                {
                keyValuePair.Value.WriteToXml(writer, GetEntryXmlElementName(keyValuePair.Value));
                }
            }

        /// <summary>
        /// Gets the entries.
        /// </summary>
        /// <value>The entries.</value>
        internal Dictionary<TKey, TEntry> Entries
            {
            get { return entries; }
            }

        /// <summary>
        /// Determines whether this instance contains the specified key.
        /// </summary>
        /// <param name="key">The key.</param>
        /// <returns>
        ///     <c>true</c> if this instance contains the specified key; otherwise, <c>false</c>.
        /// </returns>
        public bool Contains(TKey key)
            {
            return Entries.ContainsKey(key);
            }

        #region ICustomXmlUpdateSerializer Members

        /// <summary>
        /// Writes updates to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <param name="propertyDefinition">Property definition.</param>
        /// <returns>
        /// True if property generated serialization.
        /// </returns>
        bool ICustomUpdateSerializer.WriteSetUpdateToXml(
            EwsServiceXmlWriter writer,
            ServiceObject ewsObject,
            PropertyDefinition propertyDefinition)
            {
            List<TEntry> tempEntries = new();

            foreach (TKey key in addedEntries)
                {
                tempEntries.Add(entries[key]);
                }
            foreach (TKey key in modifiedEntries)
                {
                tempEntries.Add(entries[key]);
                }

            foreach (TEntry entry in tempEntries)
                {
                if (!entry.WriteSetUpdateToXml(
                    writer,
                    ewsObject,
                    propertyDefinition.XmlElementName))
                    {
                    writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetSetFieldXmlElementName());
                    WriteUriToXml(writer, entry.Key);

                    writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetXmlElementName());
                    writer.WriteStartElement(XmlNamespace.Types, propertyDefinition.XmlElementName);
                    entry.WriteToXml(writer, GetEntryXmlElementName(entry));
                    writer.WriteEndElement();
                    writer.WriteEndElement();

                    writer.WriteEndElement();
                    }
                }

            foreach (TEntry entry in removedEntries.Values)
                {
                if (!entry.WriteDeleteUpdateToXml(writer, ewsObject))
                    {
                    writer.WriteStartElement(XmlNamespace.Types, ewsObject.GetDeleteFieldXmlElementName());
                    WriteUriToXml(writer, entry.Key);
                    writer.WriteEndElement();
                    }
                }

            return true;
            }

        /// <summary>
        /// Writes deletion update to XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        /// <param name="ewsObject">The ews object.</param>
        /// <returns>
        /// True if property generated serialization.
        /// </returns>
        bool ICustomUpdateSerializer.WriteDeleteUpdateToXml(EwsServiceXmlWriter writer, ServiceObject ewsObject)
            {
            return false;
            }

        #endregion
        }
    }