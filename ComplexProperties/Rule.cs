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
    /// Represents a rule that automatically handles incoming messages.
    /// A rule consists of a set of conditions and exceptions that determine whether or 
    /// not a set of actions should be executed on incoming messages.
    /// </summary>
    public sealed class Rule : ComplexProperty
        {
        /// <summary>
        /// The rule ID.
        /// </summary>
        private string ruleId;

        /// <summary>
        /// The rule display name.
        /// </summary>
        private string displayName;

        /// <summary>
        /// The rule priority.
        /// </summary>
        private int priority;

        /// <summary>
        /// The rule status of enabled or not.
        /// </summary>
        private bool isEnabled;

        /// <summary>
        /// The rule status of is supported or not.
        /// </summary>
        private bool isNotSupported;

        /// <summary>
        /// The rule status of in error or not.
        /// </summary>
        private bool isInError;

        /// <summary>
        /// The rule conditions. 
        /// </summary>
        private RulePredicates conditions;

        /// <summary>
        /// The rule actions. 
        /// </summary>
        private RuleActions actions;

        /// <summary>
        /// The rule exceptions. 
        /// </summary>
        private RulePredicates exceptions;

        /// <summary>
        /// Initializes a new instance of the <see cref="Rule"/> class.
        /// </summary>
        public Rule()
            : base()
            {
            //// New rule has priority as 0 by default
            priority = 1;
            //// New rule is enabled by default
            isEnabled = true;
            conditions = new RulePredicates();
            actions = new RuleActions();
            exceptions = new RulePredicates();
            }

        /// <summary>
        /// Gets or sets the Id of this rule.
        /// </summary>
        public string Id
            {
            get
                {
                return ruleId;
                }

            set
                {
                SetFieldValue<string>(ref ruleId, value);
                }
            }

        /// <summary>
        /// Gets or sets the name of this rule as it should be displayed to the user.
        /// </summary>
        public string DisplayName
            {
            get
                {
                return displayName;
                }

            set
                {
                SetFieldValue<string>(ref displayName, value);
                }
            }

        /// <summary>
        /// Gets or sets the priority of this rule, which determines its execution order.
        /// </summary>
        public int Priority
            {
            get
                {
                return priority;
                }

            set
                {
                SetFieldValue<int>(ref priority, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether this rule is enabled.
        /// </summary>
        public bool IsEnabled
            {
            get
                {
                return isEnabled;
                }

            set
                {
                SetFieldValue<bool>(ref isEnabled, value);
                }
            }

        /// <summary>
        /// Gets a value indicating whether this rule can be modified via EWS. 
        /// If IsNotSupported is true, the rule cannot be modified via EWS.
        /// </summary>
        public bool IsNotSupported
            {
            get
                {
                return isNotSupported;
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether this rule has errors. A rule that is in error 
        /// cannot be processed unless it is updated and the error is corrected.
        /// </summary>
        public bool IsInError
            {
            get
                {
                return isInError;
                }

            set
                {
                SetFieldValue<bool>(ref isInError, value);
                }
            }

        /// <summary>
        /// Gets the conditions that determine whether or not this rule should be
        /// executed against incoming messages.
        /// </summary>
        public RulePredicates Conditions
            {
            get
                {
                return conditions;
                }
            }

        /// <summary>
        /// Gets the actions that should be executed against incoming messages if the
        /// conditions evaluate as true.
        /// </summary>
        public RuleActions Actions
            {
            get
                {
                return actions;
                }
            }

        /// <summary>
        /// Gets the exceptions that determine if this rule should be skipped even if 
        /// its conditions evaluate to true.
        /// </summary>
        public RulePredicates Exceptions
            {
            get
                {
                return exceptions;
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
                case XmlElementNames.DisplayName:
                    displayName = reader.ReadElementValue();
                    return true;
                case XmlElementNames.RuleId:
                    ruleId = reader.ReadElementValue();
                    return true;
                case XmlElementNames.Priority:
                    priority = reader.ReadElementValue<int>();
                    return true;
                case XmlElementNames.IsEnabled:
                    isEnabled = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsNotSupported:
                    isNotSupported = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsInError:
                    isInError = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Conditions:
                    conditions.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Actions:
                    actions.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.Exceptions:
                    exceptions.LoadFromXml(reader, reader.LocalName);
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
            if (!string.IsNullOrEmpty(Id))
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.RuleId,
                    Id);
                }

            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.DisplayName,
                DisplayName);
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.Priority,
                Priority);
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.IsEnabled,
                IsEnabled);
            writer.WriteElementValue(
                XmlNamespace.Types,
                XmlElementNames.IsInError,
                IsInError);
            Conditions.WriteToXml(writer, XmlElementNames.Conditions);
            Exceptions.WriteToXml(writer, XmlElementNames.Exceptions);
            Actions.WriteToXml(writer, XmlElementNames.Actions);
            }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
            {
            base.InternalValidate();
            EwsUtilities.ValidateParam(displayName, "DisplayName");
            EwsUtilities.ValidateParam(conditions, "Conditions");
            EwsUtilities.ValidateParam(exceptions, "Exceptions");
            EwsUtilities.ValidateParam(actions, "Actions");
            }
        }
    }