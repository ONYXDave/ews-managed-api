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
    /// Represents the set of conditions and exceptions available for a rule.
    /// </summary>
    public sealed class RulePredicates : ComplexProperty
        {
        /// <summary>
        /// The HasCategories predicate.
        /// </summary>
        private StringList categories;

        /// <summary>
        /// The ContainsBodyStrings predicate.
        /// </summary>
        private StringList containsBodyStrings;

        /// <summary>
        /// The ContainsHeaderStrings predicate.
        /// </summary>
        private StringList containsHeaderStrings;

        /// <summary>
        /// The ContainsRecipientStrings predicate.
        /// </summary>
        private StringList containsRecipientStrings;

        /// <summary>
        /// The ContainsSenderStrings predicate.
        /// </summary>
        private StringList containsSenderStrings;

        /// <summary>
        /// The ContainsSubjectOrBodyStrings predicate.
        /// </summary>
        private StringList containsSubjectOrBodyStrings;

        /// <summary>
        /// The ContainsSubjectStrings predicate.
        /// </summary>
        private StringList containsSubjectStrings;

        /// <summary>
        /// The FlaggedForAction predicate.
        /// </summary>
        private FlaggedForAction? flaggedForAction;

        /// <summary>
        /// The FromAddresses predicate.
        /// </summary>
        private EmailAddressCollection fromAddresses;

        /// <summary>
        /// The FromConnectedAccounts predicate.
        /// </summary>
        private StringList fromConnectedAccounts;

        /// <summary>
        /// The HasAttachments predicate.
        /// </summary>
        private bool hasAttachments;

        /// <summary>
        /// The Importance predicate.
        /// </summary>
        private Importance? importance;

        /// <summary>
        /// The IsApprovalRequest predicate.
        /// </summary>
        private bool isApprovalRequest;

        /// <summary>
        /// The IsAutomaticForward predicate.
        /// </summary>
        private bool isAutomaticForward;

        /// <summary>
        /// The IsAutomaticReply predicate.
        /// </summary>
        private bool isAutomaticReply;

        /// <summary>
        /// The IsEncrypted predicate.
        /// </summary>
        private bool isEncrypted;

        /// <summary>
        /// The IsMeetingRequest predicate.
        /// </summary>
        private bool isMeetingRequest;

        /// <summary>
        /// The IsMeetingResponse predicate.
        /// </summary>
        private bool isMeetingResponse;

        /// <summary>
        /// The IsNDR predicate.
        /// </summary>
        private bool isNonDeliveryReport;

        /// <summary>
        /// The IsPermissionControlled predicate.
        /// </summary>
        private bool isPermissionControlled;

        /// <summary>
        /// The IsSigned predicate.
        /// </summary>
        private bool isSigned;

        /// <summary>
        /// The IsVoicemail predicate.
        /// </summary>
        private bool isVoicemail;

        /// <summary>
        /// The IsReadReceipt  predicate.
        /// </summary>
        private bool isReadReceipt;

        /// <summary>
        /// ItemClasses predicate.
        /// </summary>
        private StringList itemClasses;

        /// <summary>
        /// The MessageClassifications predicate.
        /// </summary>
        private StringList messageClassifications;

        /// <summary>
        /// The NotSentToMe predicate.
        /// </summary>
        private bool notSentToMe;

        /// <summary>
        /// SentCcMe predicate.
        /// </summary>
        private bool sentCcMe;

        /// <summary>
        /// The SentOnlyToMe predicate.
        /// </summary>
        private bool sentOnlyToMe;

        /// <summary>
        /// The SentToAddresses predicate.
        /// </summary>
        private EmailAddressCollection sentToAddresses;

        /// <summary>
        /// The SentToMe predicate.
        /// </summary>
        private bool sentToMe;

        /// <summary>
        /// The SentToOrCcMe predicate.
        /// </summary>
        private bool sentToOrCcMe;

        /// <summary>
        /// The Sensitivity predicate.
        /// </summary>
        private Sensitivity? sensitivity;

        /// <summary>
        /// The Sensitivity predicate.
        /// </summary>
        private RulePredicateDateRange withinDateRange;

        /// <summary>
        /// The Sensitivity predicate.
        /// </summary>
        private RulePredicateSizeRange withinSizeRange;

        /// <summary>
        /// Initializes a new instance of the <see cref="RulePredicates"/> class.
        /// </summary>
        internal RulePredicates()
            : base()
            {
            categories = new StringList();
            containsBodyStrings = new StringList();
            containsHeaderStrings = new StringList();
            containsRecipientStrings = new StringList();
            containsSenderStrings = new StringList();
            containsSubjectOrBodyStrings = new StringList();
            containsSubjectStrings = new StringList();
            fromAddresses = new EmailAddressCollection(XmlElementNames.Address);
            fromConnectedAccounts = new StringList();
            itemClasses = new StringList();
            messageClassifications = new StringList();
            sentToAddresses = new EmailAddressCollection(XmlElementNames.Address);
            withinDateRange = new RulePredicateDateRange();
            withinSizeRange = new RulePredicateSizeRange();
            }

        /// <summary>
        /// Gets the categories that an incoming message should be stamped with 
        /// for the condition or exception to apply. To disable this predicate,
        /// empty the list.
        /// </summary>
        public StringList Categories
            {
            get
                {
                return categories;
                }
            }

        /// <summary>
        /// Gets the strings that should appear in the body of incoming messages 
        /// for the condition or exception to apply.
        /// To disable this predicate, empty the list.
        /// </summary>
        public StringList ContainsBodyStrings
            {
            get
                {
                return containsBodyStrings;
                }
            }

        /// <summary>
        /// Gets the strings that should appear in the headers of incoming messages 
        /// for the condition or exception to apply. To disable this predicate, empty 
        /// the list.
        /// </summary>
        public StringList ContainsHeaderStrings
            {
            get
                {
                return containsHeaderStrings;
                }
            }

        /// <summary>
        /// Gets the strings that should appear in either the To or Cc fields of 
        /// incoming messages for the condition or exception to apply. To disable this
        /// predicate, empty the list.
        /// </summary>
        public StringList ContainsRecipientStrings
            {
            get
                {
                return containsRecipientStrings;
                }
            }

        /// <summary>
        /// Gets the strings that should appear in the From field of incoming messages 
        /// for the condition or exception to apply. To disable this predicate, empty 
        /// the list.
        /// </summary>
        public StringList ContainsSenderStrings
            {
            get
                {
                return containsSenderStrings;
                }
            }

        /// <summary>
        /// Gets the strings that should appear in either the body or the subject 
        /// of incoming messages for the condition or exception to apply.
        /// To disable this predicate, empty the list.
        /// </summary>
        public StringList ContainsSubjectOrBodyStrings
            {
            get
                {
                return containsSubjectOrBodyStrings;
                }
            }

        /// <summary>
        /// Gets the strings that should appear in the subject of incoming messages 
        /// for the condition or exception to apply. To disable this predicate, 
        /// empty the list.
        /// </summary>
        public StringList ContainsSubjectStrings
            {
            get
                {
                return containsSubjectStrings;
                }
            }

        /// <summary>
        /// Gets or sets the flag for action value that should appear on incoming 
        /// messages for the condition or execption to apply. To disable this 
        /// predicate, set it to null. 
        /// </summary>
        public FlaggedForAction? FlaggedForAction
            {
            get
                {
                return flaggedForAction;
                }

            set
                {
                SetFieldValue<FlaggedForAction?>(ref flaggedForAction, value);
                }
            }

        /// <summary>
        /// Gets the e-mail addresses of the senders of incoming messages for the 
        /// condition or exception to apply. To disable this predicate, empty the 
        /// list.
        /// </summary>
        public EmailAddressCollection FromAddresses
            {
            get
                {
                return fromAddresses;
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must have
        /// attachments for the condition or exception to apply.  
        /// </summary>
        public bool HasAttachments
            {
            get
                {
                return hasAttachments;
                }

            set
                {
                SetFieldValue<bool>(ref hasAttachments, value);
                }
            }

        /// <summary>
        /// Gets or sets the importance that should be stamped on incoming messages 
        /// for the condition or exception to apply. To disable this predicate, set 
        /// it to null.
        /// </summary>
        public Importance? Importance
            {
            get
                {
                return importance;
                }

            set
                {
                SetFieldValue<Importance?>(ref importance, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// approval requests for the condition or exception to apply. 
        /// </summary>
        public bool IsApprovalRequest
            {
            get
                {
                return isApprovalRequest;
                }

            set
                {
                SetFieldValue<bool>(ref isApprovalRequest, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// automatic forwards for the condition or exception to apply.
        /// </summary>
        public bool IsAutomaticForward
            {
            get
                {
                return isAutomaticForward;
                }

            set
                {
                SetFieldValue<bool>(ref isAutomaticForward, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// automatic replies for the condition or exception to apply. 
        /// </summary>
        public bool IsAutomaticReply
            {
            get
                {
                return isAutomaticReply;
                }

            set
                {
                SetFieldValue<bool>(ref isAutomaticReply, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// S/MIME encrypted for the condition or exception to apply.
        /// </summary>
        public bool IsEncrypted
            {
            get
                {
                return isEncrypted;
                }

            set
                {
                SetFieldValue<bool>(ref isEncrypted, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// meeting requests for the condition or exception to apply. 
        /// </summary>
        public bool IsMeetingRequest
            {
            get
                {
                return isMeetingRequest;
                }

            set
                {
                SetFieldValue<bool>(ref isMeetingRequest, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// meeting responses for the condition or exception to apply. 
        /// </summary>
        public bool IsMeetingResponse
            {
            get
                {
                return isMeetingResponse;
                }

            set
                {
                SetFieldValue<bool>(ref isMeetingResponse, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// non-delivery reports (NDR) for the condition or exception to apply. 
        /// </summary>
        public bool IsNonDeliveryReport
            {
            get
                {
                return isNonDeliveryReport;
                }

            set
                {
                SetFieldValue<bool>(ref isNonDeliveryReport, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// permission controlled (RMS protected) for the condition or exception 
        /// to apply. 
        /// </summary>
        public bool IsPermissionControlled
            {
            get
                {
                return isPermissionControlled;
                }

            set
                {
                SetFieldValue<bool>(ref isPermissionControlled, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// S/MIME signed for the condition or exception to apply. 
        /// </summary>
        public bool IsSigned
            {
            get
                {
                return isSigned;
                }

            set
                {
                SetFieldValue<bool>(ref isSigned, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// voice mails for the condition or exception to apply. 
        /// </summary>
        public bool IsVoicemail
            {
            get
                {
                return isVoicemail;
                }

            set
                {
                SetFieldValue<bool>(ref isVoicemail, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether incoming messages must be 
        /// read receipts for the condition or exception to apply. 
        /// </summary>
        public bool IsReadReceipt
            {
            get
                {
                return isReadReceipt;
                }

            set
                {
                SetFieldValue<bool>(ref isReadReceipt, value);
                }
            }

        /// <summary>
        /// Gets the e-mail account names from which incoming messages must have 
        /// been aggregated for the condition or exception to apply. To disable 
        /// this predicate, empty the list.
        /// </summary>
        public StringList FromConnectedAccounts
            {
            get
                {
                return fromConnectedAccounts;
                }
            }

        /// <summary>
        /// Gets the item classes that must be stamped on incoming messages for
        /// the condition or exception to apply. To disable this predicate, 
        /// empty the list.
        /// </summary>
        public StringList ItemClasses
            {
            get
                {
                return itemClasses;
                }
            }

        /// <summary>
        /// Gets the message classifications that must be stamped on incoming messages
        /// for the condition or exception to apply. To disable this predicate, 
        /// empty the list.
        /// </summary>
        public StringList MessageClassifications
            {
            get
                {
                return messageClassifications;
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must 
        /// NOT be a To recipient of the incoming messages for the condition or 
        /// exception to apply.
        /// </summary>
        public bool NotSentToMe
            {
            get
                {
                return notSentToMe;
                }

            set
                {
                SetFieldValue<bool>(ref notSentToMe, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must be 
        /// a Cc recipient of incoming messages for the condition or exception to apply. 
        /// </summary>
        public bool SentCcMe
            {
            get
                {
                return sentCcMe;
                }

            set
                {
                SetFieldValue<bool>(ref sentCcMe, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must be 
        /// the only To recipient of incoming messages for the condition or exception 
        /// to apply.
        /// </summary>
        public bool SentOnlyToMe
            {
            get
                {
                return sentOnlyToMe;
                }

            set
                {
                SetFieldValue<bool>(ref sentOnlyToMe, value);
                }
            }

        /// <summary>
        /// Gets the e-mail addresses incoming messages must have been sent to for 
        /// the condition or exception to apply. To disable this predicate, empty 
        /// the list.
        /// </summary>
        public EmailAddressCollection SentToAddresses
            {
            get
                {
                return sentToAddresses;
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must be 
        /// a To recipient of incoming messages for the condition or exception to apply. 
        /// </summary>
        public bool SentToMe
            {
            get
                {
                return sentToMe;
                }

            set
                {
                SetFieldValue<bool>(ref sentToMe, value);
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether the owner of the mailbox must be
        /// either a To or Cc recipient of incoming messages for the condition or
        /// exception to apply. 
        /// </summary>
        public bool SentToOrCcMe
            {
            get
                {
                return sentToOrCcMe;
                }

            set
                {
                SetFieldValue<bool>(ref sentToOrCcMe, value);
                }
            }

        /// <summary>
        /// Gets or sets the sensitivity that must be stamped on incoming messages 
        /// for the condition or exception to apply. To disable this predicate, set it
        /// to null.
        /// </summary>
        public Sensitivity? Sensitivity
            {
            get
                {
                return sensitivity;
                }

            set
                {
                SetFieldValue<Sensitivity?>(ref sensitivity, value);
                }
            }

        /// <summary>
        /// Gets the date range within which incoming messages must have been received 
        /// for the condition or exception to apply. To disable this predicate, set both 
        /// its Start and End properties to null.
        /// </summary>
        public RulePredicateDateRange WithinDateRange
            {
            get
                {
                return withinDateRange;
                }
            }

        /// <summary>
        /// Gets the minimum and maximum sizes incoming messages must have for the 
        /// condition or exception to apply. To disable this predicate, set both its 
        /// MinimumSize and MaximumSize properties to null.
        /// </summary>
        public RulePredicateSizeRange WithinSizeRange
            {
            get
                {
                return withinSizeRange;
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
                case XmlElementNames.Categories:
                    categories.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsBodyStrings:
                    containsBodyStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsHeaderStrings:
                    containsHeaderStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsRecipientStrings:
                    containsRecipientStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsSenderStrings:
                    containsSenderStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsSubjectOrBodyStrings:
                    containsSubjectOrBodyStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.ContainsSubjectStrings:
                    containsSubjectStrings.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.FlaggedForAction:
                    flaggedForAction = reader.ReadElementValue<FlaggedForAction>();
                    return true;
                case XmlElementNames.FromAddresses:
                    fromAddresses.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.FromConnectedAccounts:
                    fromConnectedAccounts.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.HasAttachments:
                    hasAttachments = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Importance:
                    importance = reader.ReadElementValue<Importance>();
                    return true;
                case XmlElementNames.IsApprovalRequest:
                    isApprovalRequest = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsAutomaticForward:
                    isAutomaticForward = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsAutomaticReply:
                    isAutomaticReply = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsEncrypted:
                    isEncrypted = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsMeetingRequest:
                    isMeetingRequest = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsMeetingResponse:
                    isMeetingResponse = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsNDR:
                    isNonDeliveryReport = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsPermissionControlled:
                    isPermissionControlled = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsSigned:
                    isSigned = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsVoicemail:
                    isVoicemail = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.IsReadReceipt:
                    isReadReceipt = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.ItemClasses:
                    itemClasses.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.MessageClassifications:
                    messageClassifications.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.NotSentToMe:
                    notSentToMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SentCcMe:
                    sentCcMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SentOnlyToMe:
                    sentOnlyToMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SentToAddresses:
                    sentToAddresses.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.SentToMe:
                    sentToMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.SentToOrCcMe:
                    sentToOrCcMe = reader.ReadElementValue<bool>();
                    return true;
                case XmlElementNames.Sensitivity:
                    sensitivity = reader.ReadElementValue<Sensitivity>();
                    return true;
                case XmlElementNames.WithinDateRange:
                    withinDateRange.LoadFromXml(reader, reader.LocalName);
                    return true;
                case XmlElementNames.WithinSizeRange:
                    withinSizeRange.LoadFromXml(reader, reader.LocalName);
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
            if (Categories.Count > 0)
                {
                Categories.WriteToXml(writer, XmlElementNames.Categories);
                }

            if (ContainsBodyStrings.Count > 0)
                {
                ContainsBodyStrings.WriteToXml(writer, XmlElementNames.ContainsBodyStrings);
                }

            if (ContainsHeaderStrings.Count > 0)
                {
                ContainsHeaderStrings.WriteToXml(writer, XmlElementNames.ContainsHeaderStrings);
                }

            if (ContainsRecipientStrings.Count > 0)
                {
                ContainsRecipientStrings.WriteToXml(writer, XmlElementNames.ContainsRecipientStrings);
                }

            if (ContainsSenderStrings.Count > 0)
                {
                ContainsSenderStrings.WriteToXml(writer, XmlElementNames.ContainsSenderStrings);
                }

            if (ContainsSubjectOrBodyStrings.Count > 0)
                {
                ContainsSubjectOrBodyStrings.WriteToXml(writer, XmlElementNames.ContainsSubjectOrBodyStrings);
                }

            if (ContainsSubjectStrings.Count > 0)
                {
                ContainsSubjectStrings.WriteToXml(writer, XmlElementNames.ContainsSubjectStrings);
                }

            if (FlaggedForAction.HasValue)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.FlaggedForAction,
                    FlaggedForAction.Value);
                }

            if (FromAddresses.Count > 0)
                {
                FromAddresses.WriteToXml(writer, XmlElementNames.FromAddresses);
                }

            if (FromConnectedAccounts.Count > 0)
                {
                FromConnectedAccounts.WriteToXml(writer, XmlElementNames.FromConnectedAccounts);
                }

            if (HasAttachments != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.HasAttachments,
                    HasAttachments);
                }

            if (Importance.HasValue)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Importance,
                    Importance.Value);
                }

            if (IsApprovalRequest != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsApprovalRequest,
                    IsApprovalRequest);
                }

            if (IsAutomaticForward != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsAutomaticForward,
                    IsAutomaticForward);
                }

            if (IsAutomaticReply != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsAutomaticReply,
                    IsAutomaticReply);
                }

            if (IsEncrypted != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsEncrypted,
                    IsEncrypted);
                }

            if (IsMeetingRequest != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsMeetingRequest,
                    IsMeetingRequest);
                }

            if (IsMeetingResponse != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsMeetingResponse,
                    IsMeetingResponse);
                }

            if (IsNonDeliveryReport != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsNDR,
                    IsNonDeliveryReport);
                }

            if (IsPermissionControlled != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsPermissionControlled,
                    IsPermissionControlled);
                }

            if (isReadReceipt != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsReadReceipt,
                    IsReadReceipt);
                }

            if (IsSigned != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsSigned,
                    IsSigned);
                }

            if (IsVoicemail != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.IsVoicemail,
                    IsVoicemail);
                }

            if (ItemClasses.Count > 0)
                {
                ItemClasses.WriteToXml(writer, XmlElementNames.ItemClasses);
                }

            if (MessageClassifications.Count > 0)
                {
                MessageClassifications.WriteToXml(writer, XmlElementNames.MessageClassifications);
                }

            if (NotSentToMe != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.NotSentToMe,
                    NotSentToMe);
                }

            if (SentCcMe != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.SentCcMe,
                    SentCcMe);
                }

            if (SentOnlyToMe != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.SentOnlyToMe,
                    SentOnlyToMe);
                }

            if (SentToAddresses.Count > 0)
                {
                SentToAddresses.WriteToXml(writer, XmlElementNames.SentToAddresses);
                }

            if (SentToMe != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.SentToMe,
                    SentToMe);
                }

            if (SentToOrCcMe != false)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.SentToOrCcMe,
                    SentToOrCcMe);
                }

            if (Sensitivity.HasValue)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.Sensitivity,
                    Sensitivity.Value);
                }

            if (WithinDateRange.Start.HasValue || WithinDateRange.End.HasValue)
                {
                WithinDateRange.WriteToXml(writer, XmlElementNames.WithinDateRange);
                }

            if (WithinSizeRange.MaximumSize.HasValue || WithinSizeRange.MinimumSize.HasValue)
                {
                WithinSizeRange.WriteToXml(writer, XmlElementNames.WithinSizeRange);
                }
            }

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal override void InternalValidate()
            {
            base.InternalValidate();
            EwsUtilities.ValidateParam(fromAddresses, "FromAddresses");
            EwsUtilities.ValidateParam(sentToAddresses, "SentToAddresses");
            EwsUtilities.ValidateParam(withinDateRange, "WithinDateRange");
            EwsUtilities.ValidateParam(withinSizeRange, "WithinSizeRange");
            }
        }
    }