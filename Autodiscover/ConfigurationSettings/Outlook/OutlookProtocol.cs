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

namespace Microsoft.Exchange.WebServices.Autodiscover
    {
    using Microsoft.Exchange.WebServices.Data;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Linq;
    using System.Xml;
    using ConverterDictionary = System.Collections.Generic.Dictionary<UserSettingName, System.Func<OutlookProtocol, object>>;
    using ConverterPair = System.Collections.Generic.KeyValuePair<UserSettingName, System.Func<OutlookProtocol, object>>;

    /// <summary>
    /// Represents a supported Outlook protocol in an Outlook configurations settings account.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    internal sealed class OutlookProtocol
        {
        #region Private constants
        private const string EXCH = "EXCH";
        private const string EXPR = "EXPR";
        private const string WEB = "WEB";
        #endregion

        #region Private static fields
        /// <summary>
        /// Converters to translate common Outlook protocol settings.
        /// Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance. 
        /// </summary>
        private static LazyMember<ConverterDictionary> commonProtocolSettings = new(
            () =>
            {
                ConverterDictionary results = new();
                results.Add(UserSettingName.EcpDeliveryReportUrlFragment, p => p.ecpUrlMt);
                results.Add(UserSettingName.EcpEmailSubscriptionsUrlFragment, p => p.ecpUrlAggr);
                results.Add(UserSettingName.EcpPublishingUrlFragment, p => p.ecpUrlPublish);
                results.Add(UserSettingName.EcpPhotoUrlFragment, p => p.ecpUrlPhoto);
                results.Add(UserSettingName.EcpRetentionPolicyTagsUrlFragment, p => p.ecpUrlRet);
                results.Add(UserSettingName.EcpTextMessagingUrlFragment, p => p.ecpUrlSms);
                results.Add(UserSettingName.EcpVoicemailUrlFragment, p => p.ecpUrlUm);
                results.Add(UserSettingName.EcpConnectUrlFragment, p => p.ecpUrlConnect);
                results.Add(UserSettingName.EcpTeamMailboxUrlFragment, p => p.ecpUrlTm);
                results.Add(UserSettingName.EcpTeamMailboxCreatingUrlFragment, p => p.ecpUrlTmCreating);
                results.Add(UserSettingName.EcpTeamMailboxEditingUrlFragment, p => p.ecpUrlTmEditing);
                results.Add(UserSettingName.EcpExtensionInstallationUrlFragment, p => p.ecpUrlExtInstall);
                results.Add(UserSettingName.SiteMailboxCreationURL, p => p.siteMailboxCreationURL);
                return results;
            });

        /// <summary>
        /// Converters to translate internal (EXCH) Outlook protocol settings.
        /// Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance. 
        /// </summary>
        private static LazyMember<ConverterDictionary> internalProtocolSettings = new(
            () =>
            {
                ConverterDictionary results = new();
                results.Add(UserSettingName.ActiveDirectoryServer, p => p.activeDirectoryServer);
                results.Add(UserSettingName.CrossOrganizationSharingEnabled, p => p.sharingEnabled.ToString());
                results.Add(UserSettingName.InternalEcpUrl, p => p.ecpUrl);
                results.Add(UserSettingName.InternalEcpDeliveryReportUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlMt));
                results.Add(UserSettingName.InternalEcpEmailSubscriptionsUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlAggr));
                results.Add(UserSettingName.InternalEcpPublishingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlPublish));
                results.Add(UserSettingName.InternalEcpPhotoUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlPhoto));
                results.Add(UserSettingName.InternalEcpRetentionPolicyTagsUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlRet));
                results.Add(UserSettingName.InternalEcpTextMessagingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlSms));
                results.Add(UserSettingName.InternalEcpVoicemailUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlUm));
                results.Add(UserSettingName.InternalEcpConnectUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlConnect));
                results.Add(UserSettingName.InternalEcpTeamMailboxUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlTm));
                results.Add(UserSettingName.InternalEcpTeamMailboxCreatingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlTmCreating));
                results.Add(UserSettingName.InternalEcpTeamMailboxEditingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlTmEditing));
                results.Add(UserSettingName.InternalEcpTeamMailboxHidingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlTmHiding));
                results.Add(UserSettingName.InternalEcpExtensionInstallationUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlExtInstall));
                results.Add(UserSettingName.InternalEwsUrl, p => p.exchangeWebServicesUrl ?? p.availabilityServiceUrl);
                results.Add(UserSettingName.InternalEmwsUrl, p => p.exchangeManagementWebServicesUrl);
                results.Add(UserSettingName.InternalMailboxServerDN, p => p.serverDN);
                results.Add(UserSettingName.InternalRpcClientServer, p => p.server);
                results.Add(UserSettingName.InternalOABUrl, p => p.offlineAddressBookUrl);
                results.Add(UserSettingName.InternalUMUrl, p => p.unifiedMessagingUrl);
                results.Add(UserSettingName.MailboxDN, p => p.mailboxDN);
                results.Add(UserSettingName.PublicFolderServer, p => p.publicFolderServer);
                results.Add(UserSettingName.InternalServerExclusiveConnect, p => p.serverExclusiveConnect);
                return results;
            });

        /// <summary>
        /// Converters to translate external (EXPR) Outlook protocol settings.
        /// Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance. 
        /// </summary>
        private static LazyMember<ConverterDictionary> externalProtocolSettings = new(
            () =>
            {
                ConverterDictionary results = new();
                results.Add(UserSettingName.ExternalEcpDeliveryReportUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlRet));
                results.Add(UserSettingName.ExternalEcpEmailSubscriptionsUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlAggr));
                results.Add(UserSettingName.ExternalEcpPublishingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlPublish));
                results.Add(UserSettingName.ExternalEcpPhotoUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlPhoto));
                results.Add(UserSettingName.ExternalEcpRetentionPolicyTagsUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlRet));
                results.Add(UserSettingName.ExternalEcpTextMessagingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlSms));
                results.Add(UserSettingName.ExternalEcpUrl, p => p.ecpUrl);
                results.Add(UserSettingName.ExternalEcpVoicemailUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlUm));
                results.Add(UserSettingName.ExternalEcpConnectUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlConnect));
                results.Add(UserSettingName.ExternalEcpTeamMailboxUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlTm));
                results.Add(UserSettingName.ExternalEcpTeamMailboxCreatingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlTmCreating));
                results.Add(UserSettingName.ExternalEcpTeamMailboxEditingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlTmEditing));
                results.Add(UserSettingName.ExternalEcpTeamMailboxHidingUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlTmHiding));
                results.Add(UserSettingName.ExternalEcpExtensionInstallationUrl, p => p.ConvertEcpFragmentToUrl(p.ecpUrlExtInstall));
                results.Add(UserSettingName.ExternalEwsUrl, p => p.exchangeWebServicesUrl ?? p.availabilityServiceUrl);
                results.Add(UserSettingName.ExternalEmwsUrl, p => p.exchangeManagementWebServicesUrl);
                results.Add(UserSettingName.ExternalMailboxServer, p => p.server);
                results.Add(UserSettingName.ExternalMailboxServerAuthenticationMethods, p => p.authPackage);
                results.Add(UserSettingName.ExternalMailboxServerRequiresSSL, p => p.sslEnabled.ToString());
                results.Add(UserSettingName.ExternalOABUrl, p => p.offlineAddressBookUrl);
                results.Add(UserSettingName.ExternalUMUrl, p => p.unifiedMessagingUrl);
                results.Add(UserSettingName.ExchangeRpcUrl, p => p.exchangeRpcUrl);
                results.Add(UserSettingName.EwsPartnerUrl, p => p.exchangeWebServicesPartnerUrl);
                results.Add(UserSettingName.ExternalServerExclusiveConnect, p => p.serverExclusiveConnect.ToString());
                results.Add(UserSettingName.CertPrincipalName, p => p.certPrincipalName);
                results.Add(UserSettingName.GroupingInformation, p => p.groupingInformation);
                return results;
            });

        /// <summary>
        /// Merged converter dictionary for translating internal (EXCH) Outlook protocol settings.
        /// Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance. 
        /// </summary>
        private static LazyMember<ConverterDictionary> internalProtocolConverterDictionary = new(
            () =>
            {
                ConverterDictionary results = new();
                commonProtocolSettings.Member.ToList<ConverterPair>().ForEach((kv) => results.Add(kv.Key, kv.Value));
                internalProtocolSettings.Member.ToList<ConverterPair>().ForEach((kv) => results.Add(kv.Key, kv.Value));
                return results;
            });

        /// <summary>
        /// Merged converter dictionary for translating external (EXPR) Outlook protocol settings.
        /// Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance. 
        /// </summary>
        private static LazyMember<ConverterDictionary> externalProtocolConverterDictionary = new(
            () =>
            {
                ConverterDictionary results = new();
                commonProtocolSettings.Member.ToList<ConverterPair>().ForEach((kv) => results.Add(kv.Key, kv.Value));
                externalProtocolSettings.Member.ToList<ConverterPair>().ForEach((kv) => results.Add(kv.Key, kv.Value));
                return results;
            });

        /// <summary>
        /// Converters to translate Web (WEB) Outlook protocol settings.
        /// Each entry maps to a lambda expression used to get the matching property from the OutlookProtocol instance. 
        /// </summary>
        private static LazyMember<ConverterDictionary> webProtocolConverterDictionary = new(
            () =>
            {
                ConverterDictionary results = new();
                results.Add(UserSettingName.InternalWebClientUrls, p => p.internalOutlookWebAccessUrls);
                results.Add(UserSettingName.ExternalWebClientUrls, p => p.externalOutlookWebAccessUrls);
                return results;
            });

        /// <summary>
        /// The collection of available user settings for all OutlookProtocol types.
        /// </summary>
        private static LazyMember<List<UserSettingName>> availableUserSettings = new(
            () =>
            {
                List<UserSettingName> results = new();
                results.AddRange(commonProtocolSettings.Member.Keys);
                results.AddRange(internalProtocolSettings.Member.Keys);
                results.AddRange(externalProtocolSettings.Member.Keys);
                results.AddRange(webProtocolConverterDictionary.Member.Keys);
                return results;
            });

        /// <summary>
        /// Map Outlook protocol name to type.
        /// </summary>
        private static LazyMember<Dictionary<string, OutlookProtocolType>> protocolNameToTypeMap = new(
            delegate ()
            {
                Dictionary<string, OutlookProtocolType> results = new();
                results.Add(OutlookProtocol.EXCH, OutlookProtocolType.Rpc);
                results.Add(OutlookProtocol.EXPR, OutlookProtocolType.RpcOverHttp);
                results.Add(OutlookProtocol.WEB, OutlookProtocolType.Web);
                return results;
                });
        #endregion

        #region Private fields
        private string activeDirectoryServer;
        private string authPackage;
        private string availabilityServiceUrl;
        private string ecpUrl;
        private string ecpUrlAggr;
        private string ecpUrlMt;
        private string ecpUrlPublish;
        private string ecpUrlPhoto;
        private string ecpUrlConnect;
        private string ecpUrlRet;
        private string ecpUrlSms;
        private string ecpUrlUm;
        private string ecpUrlTm;
        private string ecpUrlTmCreating;
        private string ecpUrlTmEditing;
        private string ecpUrlTmHiding;
        private string siteMailboxCreationURL;
        private string ecpUrlExtInstall;
        private string exchangeWebServicesUrl;
        private string exchangeManagementWebServicesUrl;
        private string mailboxDN;
        private string offlineAddressBookUrl;
        private string exchangeRpcUrl;
        private string exchangeWebServicesPartnerUrl;
        private string publicFolderServer;
        private string server;
        private string serverDN;
        private string unifiedMessagingUrl;
        private bool sharingEnabled;
        private bool sslEnabled;
        private bool serverExclusiveConnect;
        private string certPrincipalName;
        private string groupingInformation;
        private WebClientUrlCollection externalOutlookWebAccessUrls;
        private WebClientUrlCollection internalOutlookWebAccessUrls;
        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="OutlookProtocol"/> class.
        /// </summary>
        internal OutlookProtocol()
            {
            internalOutlookWebAccessUrls = new WebClientUrlCollection();
            externalOutlookWebAccessUrls = new WebClientUrlCollection();
            }

        /// <summary>
        /// Loads from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        internal void LoadFromXml(EwsXmlReader reader)
            {
            do
                {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                    {
                    switch (reader.LocalName)
                        {
                        case XmlElementNames.Type:
                            ProtocolType = OutlookProtocol.ProtocolNameToType(reader.ReadElementValue());
                            break;
                        case XmlElementNames.AuthPackage:
                            authPackage = reader.ReadElementValue();
                            break;
                        case XmlElementNames.Server:
                            server = reader.ReadElementValue();
                            break;
                        case XmlElementNames.ServerDN:
                            serverDN = reader.ReadElementValue();
                            break;
                        case XmlElementNames.ServerVersion:
                            // just read it out
                            reader.ReadElementValue();
                            break;
                        case XmlElementNames.AD:
                            activeDirectoryServer = reader.ReadElementValue();
                            break;
                        case XmlElementNames.MdbDN:
                            mailboxDN = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EWSUrl:
                            exchangeWebServicesUrl = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EmwsUrl:
                            exchangeManagementWebServicesUrl = reader.ReadElementValue();
                            break;
                        case XmlElementNames.ASUrl:
                            availabilityServiceUrl = reader.ReadElementValue();
                            break;
                        case XmlElementNames.OOFUrl:
                            // just read it out
                            reader.ReadElementValue();
                            break;
                        case XmlElementNames.UMUrl:
                            unifiedMessagingUrl = reader.ReadElementValue();
                            break;
                        case XmlElementNames.OABUrl:
                            offlineAddressBookUrl = reader.ReadElementValue();
                            break;
                        case XmlElementNames.PublicFolderServer:
                            publicFolderServer = reader.ReadElementValue();
                            break;
                        case XmlElementNames.Internal:
                            OutlookProtocol.LoadWebClientUrlsFromXml(reader, internalOutlookWebAccessUrls, reader.LocalName);
                            break;
                        case XmlElementNames.External:
                            OutlookProtocol.LoadWebClientUrlsFromXml(reader, externalOutlookWebAccessUrls, reader.LocalName);
                            break;
                        case XmlElementNames.Ssl:
                            string sslStr = reader.ReadElementValue();
                            sslEnabled = sslStr.Equals("On", StringComparison.OrdinalIgnoreCase);
                            break;
                        case XmlElementNames.SharingUrl:
                            sharingEnabled = reader.ReadElementValue().Length > 0;
                            break;
                        case XmlElementNames.EcpUrl:
                            ecpUrl = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_um:
                            ecpUrlUm = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_aggr:
                            ecpUrlAggr = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_sms:
                            ecpUrlSms = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_mt:
                            ecpUrlMt = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_ret:
                            ecpUrlRet = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_publish:
                            ecpUrlPublish = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_photo:
                            ecpUrlPhoto = reader.ReadElementValue();
                            break;
                        case XmlElementNames.ExchangeRpcUrl:
                            exchangeRpcUrl = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EwsPartnerUrl:
                            exchangeWebServicesPartnerUrl = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_connect:
                            ecpUrlConnect = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_tm:
                            ecpUrlTm = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_tmCreating:
                            ecpUrlTmCreating = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_tmEditing:
                            ecpUrlTmEditing = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_tmHiding:
                            ecpUrlTmHiding = reader.ReadElementValue();
                            break;
                        case XmlElementNames.SiteMailboxCreationURL:
                            siteMailboxCreationURL = reader.ReadElementValue();
                            break;
                        case XmlElementNames.EcpUrl_extinstall:
                            ecpUrlExtInstall = reader.ReadElementValue();
                            break;
                        case XmlElementNames.ServerExclusiveConnect:
                            string serverExclusiveConnectStr = reader.ReadElementValue();
                            serverExclusiveConnect = serverExclusiveConnectStr.Equals("On", StringComparison.OrdinalIgnoreCase);
                            break;
                        case XmlElementNames.CertPrincipalName:
                            certPrincipalName = reader.ReadElementValue();
                            break;
                        case XmlElementNames.GroupingInformation:
                            groupingInformation = reader.ReadElementValue();
                            break;
                        default:
                            reader.SkipCurrentElement();
                            break;
                        }
                    }
                }
            while (!reader.IsEndElement(XmlNamespace.NotSpecified, XmlElementNames.Protocol));
            }

        /// <summary>
        /// Convert protocol name to protocol type.
        /// </summary>
        /// <param name="protocolName">Name of the protocol.</param>
        /// <returns>OutlookProtocolType</returns>
        private static OutlookProtocolType ProtocolNameToType(string protocolName)
            {
            OutlookProtocolType protocolType;
            if (!protocolNameToTypeMap.Member.TryGetValue(protocolName, out protocolType))
                {
                protocolType = OutlookProtocolType.Unknown;
                }
            return protocolType;
            }

        /// <summary>
        /// Loads web client urls from XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="webClientUrls">The web client urls.</param>
        /// <param name="elementName">Name of the element.</param>
        private static void LoadWebClientUrlsFromXml(EwsXmlReader reader, WebClientUrlCollection webClientUrls, string elementName)
            {
            do
                {
                reader.Read();

                if (reader.NodeType == XmlNodeType.Element)
                    {
                    switch (reader.LocalName)
                        {
                        case XmlElementNames.OWAUrl:
                            string authMethod = reader.ReadAttributeValue(XmlAttributeNames.AuthenticationMethod);
                            string owaUrl = reader.ReadElementValue();
                            WebClientUrl webClientUrl = new(authMethod, owaUrl);
                            webClientUrls.Urls.Add(webClientUrl);
                            break;
                        default:
                            reader.SkipCurrentElement();
                            break;
                        }
                    }
                }
            while (!reader.IsEndElement(XmlNamespace.NotSpecified, elementName));
            }

        /// <summary>
        /// Convert ECP fragment to full ECP URL.
        /// </summary>
        /// <param name="fragment">The fragment.</param>
        /// <returns>Full URL string (or null if either portion is empty.</returns>
        private string ConvertEcpFragmentToUrl(string fragment)
            {
            return (string.IsNullOrEmpty(ecpUrl) || string.IsNullOrEmpty(fragment)) ? null : (ecpUrl + fragment);
            }

        /// <summary>
        /// Convert OutlookProtocol to GetUserSettings response.
        /// </summary>
        /// <param name="requestedSettings">The requested settings.</param>
        /// <param name="response">The response.</param>
        internal void ConvertToUserSettings(
            List<UserSettingName> requestedSettings,
            GetUserSettingsResponse response)
            {
            if (ConverterDictionary != null)
                {
                // In English: collect converters that are contained in the requested settings.
                IEnumerable<ConverterPair> converterQuery = from converter in ConverterDictionary
                                                            where requestedSettings.Contains(converter.Key)
                                                            select converter;

                foreach (ConverterPair kv in converterQuery)
                    {
                    object value = kv.Value(this);
                    if (value != null)
                        {
                        response.Settings[kv.Key] = value;
                        }
                    }
                }
            }

        /// <summary>
        /// Gets the type of the protocol.
        /// </summary>
        /// <value>The type of the protocol.</value>
        internal OutlookProtocolType ProtocolType
            {
            get; set;
            }

        /// <summary>
        /// Gets the converter dictionary for protocol type.
        /// </summary>
        /// <value>The converter dictionary.</value>
        private ConverterDictionary ConverterDictionary
            {
            get
                {
                switch (ProtocolType)
                    {
                    case OutlookProtocolType.Rpc:
                        return internalProtocolConverterDictionary.Member;
                    case OutlookProtocolType.RpcOverHttp:
                        return externalProtocolConverterDictionary.Member;
                    case OutlookProtocolType.Web:
                        return webProtocolConverterDictionary.Member;
                    default:
                        return null;
                    }
                }
            }

        /// <summary>
        /// Gets the available user settings.
        /// </summary>
        internal static List<UserSettingName> AvailableUserSettings
            {
            get
                {
                return availableUserSettings.Member;
                }
            }
        }
    }