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
    using System.IO;
    using System.Net;
    using System.Text.RegularExpressions;
    using System.Xml;

    /// <summary>
    /// Defines a delegate that is used by the AutodiscoverService to ask whether a redirectionUrl can be used.
    /// </summary>
    /// <param name="redirectionUrl">Redirection URL that Autodiscover wants to use.</param>
    /// <returns>Delegate returns true if Autodiscover is allowed to use this URL.</returns>
    public delegate bool AutodiscoverRedirectionUrlValidationCallback(string redirectionUrl);

    /// <summary>
    /// Represents a binding to the Exchange Autodiscover Service.
    /// </summary>
    public sealed class AutodiscoverService : ExchangeServiceBase
        {
        #region Static members

        /// <summary>
        /// Autodiscover legacy path
        /// </summary>
        private const string AutodiscoverLegacyPath = "/autodiscover/autodiscover.xml";

        /// <summary>
        /// Autodiscover legacy Url with protocol fill-in
        /// </summary>
        private const string AutodiscoverLegacyUrl = "{0}://{1}" + AutodiscoverLegacyPath;

        /// <summary>
        /// Autodiscover legacy HTTPS Url
        /// </summary>
        private const string AutodiscoverLegacyHttpsUrl = "https://{0}" + AutodiscoverLegacyPath;

        /// <summary>
        /// Autodiscover legacy HTTP Url
        /// </summary>
        private const string AutodiscoverLegacyHttpUrl = "http://{0}" + AutodiscoverLegacyPath;

        /// <summary>
        /// Autodiscover SOAP HTTPS Url
        /// </summary>
        private const string AutodiscoverSoapHttpsUrl = "https://{0}/autodiscover/autodiscover.svc";

        /// <summary>
        /// Autodiscover SOAP WS-Security HTTPS Url
        /// </summary>
        private const string AutodiscoverSoapWsSecurityHttpsUrl = AutodiscoverSoapHttpsUrl + "/wssecurity";

        /// <summary>
        /// Autodiscover SOAP WS-Security symmetrickey HTTPS Url
        /// </summary>
        private const string AutodiscoverSoapWsSecuritySymmetricKeyHttpsUrl = AutodiscoverSoapHttpsUrl + "/wssecurity/symmetrickey";

        /// <summary>
        /// Autodiscover SOAP WS-Security x509cert HTTPS Url
        /// </summary>
        private const string AutodiscoverSoapWsSecurityX509CertHttpsUrl = AutodiscoverSoapHttpsUrl + "/wssecurity/x509cert";

        /// <summary>
        /// Autodiscover request namespace
        /// </summary>
        private const string AutodiscoverRequestNamespace = "http://schemas.microsoft.com/exchange/autodiscover/outlook/requestschema/2006";

        /// <summary>
        /// Legacy path regular expression.
        /// </summary>
        private static readonly Regex LegacyPathRegex = new(@"/autodiscover/([^/]+/)*autodiscover.xml", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Maximum number of Url (or address) redirections that will be followed by an Autodiscover call
        /// </summary>
        internal const int AutodiscoverMaxRedirections = 10;

        /// <summary>
        /// HTTP header indicating that SOAP Autodiscover service is enabled.
        /// </summary>
        private const string AutodiscoverSoapEnabledHeaderName = "X-SOAP-Enabled";

        /// <summary>
        /// HTTP header indicating that WS-Security Autodiscover service is enabled.
        /// </summary>
        private const string AutodiscoverWsSecurityEnabledHeaderName = "X-WSSecurity-Enabled";

        /// <summary>
        /// HTTP header indicating that WS-Security/SymmetricKey Autodiscover service is enabled.
        /// </summary>
        private const string AutodiscoverWsSecuritySymmetricKeyEnabledHeaderName = "X-WSSecurity-SymmetricKey-Enabled";

        /// <summary>
        /// HTTP header indicating that WS-Security/X509Cert Autodiscover service is enabled.
        /// </summary>
        private const string AutodiscoverWsSecurityX509CertEnabledHeaderName = "X-WSSecurity-X509Cert-Enabled";

        /// <summary>
        /// HTTP header indicating that OAuth Autodiscover service is enabled.
        /// </summary>
        private const string AutodiscoverOAuthEnabledHeaderName = "X-OAuth-Enabled";

        /// <summary>
        /// Minimum request version for Autodiscover SOAP service.
        /// </summary>
        private const ExchangeVersion MinimumRequestVersionForAutoDiscoverSoapService = ExchangeVersion.Exchange2010;

        #endregion

        #region Private members

        private string domain;
        private bool? isExternal = true;
        private Uri url;
        private AutodiscoverRedirectionUrlValidationCallback redirectionUrlValidationCallback;
        private AutodiscoverDnsClient dnsClient;
        private IPAddress dnsServerAddress;
        private bool enableScpLookup = true;

        private delegate TGetSettingsResponseCollection GetSettingsMethod<TGetSettingsResponseCollection, TSettingName>(
            List<string> smtpAddresses,
            List<TSettingName> settings,
            ExchangeVersion? requestedVersion,
            ref Uri autodiscoverUrl);

        #endregion

        /// <summary>
        /// Default implementation of AutodiscoverRedirectionUrlValidationCallback.
        /// Always returns true indicating that the URL can be used.
        /// </summary>
        /// <param name="redirectionUrl">The redirection URL.</param>
        /// <returns>Returns true.</returns>
        private bool DefaultAutodiscoverRedirectionUrlValidationCallback(string redirectionUrl)
            {
            throw new AutodiscoverLocalException(string.Format(Strings.AutodiscoverRedirectBlocked, redirectionUrl));
            }

        #region Legacy Autodiscover

        /// <summary>
        /// Calls the Autodiscover service to get configuration settings at the specified URL.
        /// </summary>
        /// <typeparam name="TSettings">The type of the settings to retrieve.</typeparam>
        /// <param name="emailAddress">The email address to retrieve configuration settings for.</param>
        /// <param name="url">The URL of the Autodiscover service.</param>
        /// <returns>The requested configuration settings.</returns>
        private TSettings GetLegacyUserSettingsAtUrl<TSettings>(string emailAddress, Uri url)
            where TSettings : ConfigurationSettingsBase, new()
            {
            TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                string.Format("Trying to call Autodiscover for {0} on {1}.", emailAddress, url));

            TSettings settings = new();

            IEwsHttpWebRequest request = PrepareHttpWebRequestForUrl(url);

            TraceHttpRequestHeaders(TraceFlags.AutodiscoverRequestHttpHeaders, request);

            using (Stream requestStream = request.GetRequestStream())
                {
                Stream writerStream = requestStream;

                // If tracing is enabled, we generate the request in-memory so that we
                // can pass it along to the ITraceListener. Then we copy the stream to
                // the request stream.
                if (IsTraceEnabledFor(TraceFlags.AutodiscoverRequest))
                    {
                    using (MemoryStream memoryStream = new())
                        {
                        using (StreamWriter writer = new(memoryStream))
                            {
                            WriteLegacyAutodiscoverRequest(emailAddress, settings, writer);
                            writer.Flush();

                            TraceXml(TraceFlags.AutodiscoverRequest, memoryStream);

                            EwsUtilities.CopyStream(memoryStream, requestStream);
                            }
                        }
                    }
                else
                    {
                    using (StreamWriter writer = new(requestStream))
                        {
                        WriteLegacyAutodiscoverRequest(emailAddress, settings, writer);
                        }
                    }
                }

            using (IEwsHttpWebResponse webResponse = request.GetResponse())
                {
                Uri redirectUrl;
                if (TryGetRedirectionResponse(webResponse, out redirectUrl))
                    {
                    settings.MakeRedirectionResponse(redirectUrl);
                    return settings;
                    }

                using (Stream responseStream = webResponse.GetResponseStream())
                    {
                    // If tracing is enabled, we read the entire response into a MemoryStream so that we
                    // can pass it along to the ITraceListener. Then we parse the response from the 
                    // MemoryStream.
                    if (IsTraceEnabledFor(TraceFlags.AutodiscoverResponse))
                        {
                        using (MemoryStream memoryStream = new())
                            {
                            // Copy response stream to in-memory stream and reset to start
                            EwsUtilities.CopyStream(responseStream, memoryStream);
                            memoryStream.Position = 0;

                            TraceResponse(webResponse, memoryStream);

                            EwsXmlReader reader = new(memoryStream);
                            reader.Read(XmlNodeType.XmlDeclaration);
                            settings.LoadFromXml(reader);
                            }
                        }
                    else
                        {
                        EwsXmlReader reader = new(responseStream);
                        reader.Read(XmlNodeType.XmlDeclaration);
                        settings.LoadFromXml(reader);
                        }
                    }

                return settings;
                }
            }

        /// <summary>
        /// Writes the autodiscover request.
        /// </summary>
        /// <param name="emailAddress">The email address.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="writer">The writer.</param>
        private void WriteLegacyAutodiscoverRequest(
            string emailAddress,
            ConfigurationSettingsBase settings,
            StreamWriter writer)
            {
            writer.Write(string.Format("<Autodiscover xmlns=\"{0}\">", AutodiscoverRequestNamespace));
            writer.Write("<Request>");
            writer.Write(string.Format("<EMailAddress>{0}</EMailAddress>", emailAddress));
            writer.Write(string.Format("<AcceptableResponseSchema>{0}</AcceptableResponseSchema>", settings.GetNamespace()));
            writer.Write("</Request>");
            writer.Write("</Autodiscover>");
            }

        /// <summary>
        /// Gets a redirection URL to an SSL-enabled Autodiscover service from the standard non-SSL Autodiscover URL.
        /// </summary>
        /// <param name="domainName">The name of the domain to call Autodiscover on.</param>
        /// <returns>A valid SSL-enabled redirection URL. (May be null).</returns>
        private Uri GetRedirectUrl(string domainName)
            {
            string url = string.Format(AutodiscoverLegacyHttpUrl, "autodiscover." + domainName);

            TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                string.Format("Trying to get Autodiscover redirection URL from {0}.", url));

            IEwsHttpWebRequest request = HttpWebRequestFactory.CreateRequest(new Uri(url));

            request.Method = "GET";
            request.AllowAutoRedirect = false;
            request.PreAuthenticate = false;

            IEwsHttpWebResponse response = null;

            try
                {
                response = request.GetResponse();
                }
            catch (WebException ex)
                {
                TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    string.Format("Request error: {0}", ex.Message));

                // The exception response factory requires a valid HttpWebResponse, 
                // but there will be no web response if the web request couldn't be
                // actually be issued (e.g. due to DNS error).
                if (ex.Response != null)
                    {
                    response = HttpWebRequestFactory.CreateExceptionResponse(ex);
                    }
                }
            catch (IOException ex)
                {
                TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    string.Format("I/O error: {0}", ex.Message));
                }

            if (response != null)
                {
                using (response)
                    {
                    Uri redirectUrl;
                    if (TryGetRedirectionResponse(response, out redirectUrl))
                        {
                        return redirectUrl;
                        }
                    }
                }

            TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                "No Autodiscover redirection URL was returned.");

            return null;
            }

        /// <summary>
        /// Tries the get redirection response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="redirectUrl">The redirect URL.</param>
        /// <returns>True if a valid redirection URL was found.</returns>
        private bool TryGetRedirectionResponse(IEwsHttpWebResponse response, out Uri redirectUrl)
            {
            redirectUrl = null;
            if (AutodiscoverRequest.IsRedirectionResponse(response))
                {
                // Get the redirect location and verify that it's valid.
                string location = response.Headers[HttpResponseHeader.Location];

                if (!string.IsNullOrEmpty(location))
                    {
                    try
                        {
                        redirectUrl = new Uri(response.ResponseUri, location);

                        // Check if URL is SSL and that the path matches.
                        Match match = LegacyPathRegex.Match(redirectUrl.AbsolutePath);
                        if ((redirectUrl.Scheme == Uri.UriSchemeHttps) &&
                            match.Success)
                            {
                            TraceMessage(
                                TraceFlags.AutodiscoverConfiguration,
                                string.Format("Redirection URL found: '{0}'", redirectUrl));

                            return true;
                            }
                        }
                    catch (UriFormatException)
                        {
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            string.Format("Invalid redirection URL was returned: '{0}'", location));
                        return false;
                        }
                    }
                }

            return false;
            }

        /// <summary>
        /// Calls the legacy Autodiscover service to retrieve configuration settings.
        /// </summary>
        /// <typeparam name="TSettings">The type of the settings to retrieve.</typeparam>
        /// <param name="emailAddress">The email address to retrieve configuration settings for.</param>
        /// <returns>The requested configuration settings.</returns>
        internal TSettings GetLegacyUserSettings<TSettings>(string emailAddress)
            where TSettings : ConfigurationSettingsBase, new()
            {
            // If Url is specified, call service directly.
            if (Url != null)
                {
                Match match = LegacyPathRegex.Match(Url.AbsolutePath);
                if (match.Success)
                    {
                    return GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, Url);
                    }

                // this.Uri is intended for Autodiscover SOAP service, convert to Legacy endpoint URL.
                Uri autodiscoverUrl = new(Url, AutodiscoverLegacyPath);
                return GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, autodiscoverUrl);
                }

            // If Domain is specified, figure out the endpoint Url and call service.
            else if (!string.IsNullOrEmpty(Domain))
                {
                Uri autodiscoverUrl = new(string.Format(AutodiscoverLegacyHttpsUrl, Domain));
                return GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, autodiscoverUrl);
                }
            else
                {
                // No Url or Domain specified, need to figure out which endpoint to use.
                int currentHop = 1;
                List<string> redirectionEmailAddresses = new();
                return InternalGetLegacyUserSettings<TSettings>(
                    emailAddress,
                    redirectionEmailAddresses,
                    ref currentHop);
                }
            }

        /// <summary>
        /// Calls the legacy Autodiscover service to retrieve configuration settings.
        /// </summary>
        /// <typeparam name="TSettings">The type of the settings to retrieve.</typeparam>
        /// <param name="emailAddress">The email address to retrieve configuration settings for.</param>
        /// <param name="redirectionEmailAddresses">List of previous email addresses.</param>
        /// <param name="currentHop">Current number of redirection urls/addresses attempted so far.</param>
        /// <returns>The requested configuration settings.</returns>
        private TSettings InternalGetLegacyUserSettings<TSettings>(
            string emailAddress,
            List<string> redirectionEmailAddresses,
            ref int currentHop)
            where TSettings : ConfigurationSettingsBase, new()
            {
            string domainName = EwsUtilities.DomainFromEmailAddress(emailAddress);

            int scpUrlCount;
            List<Uri> urls = GetAutodiscoverServiceUrls(domainName, out scpUrlCount);

            if (urls.Count == 0)
                {
                throw new ServiceValidationException(Strings.AutodiscoverServiceRequestRequiresDomainOrUrl);
                }

            // Assume caller is not inside the Intranet, regardless of whether SCP Urls 
            // were returned or not. SCP Urls are only relevant if one of them returns
            // valid Autodiscover settings.
            isExternal = true;

            int currentUrlIndex = 0;

            // Used to save exception for later reporting.
            Exception delayedException = null;
            TSettings settings = null;

            do
                {
                Uri autodiscoverUrl = urls[currentUrlIndex];
                bool isScpUrl = currentUrlIndex < scpUrlCount;

                try
                    {
                    settings = GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, autodiscoverUrl);

                    switch (settings.ResponseType)
                        {
                        case AutodiscoverResponseType.Success:
                            // Not external if Autodiscover endpoint found via SCP returned the settings.
                            if (isScpUrl)
                                {
                                IsExternal = false;
                                }
                            Url = autodiscoverUrl;
                            return settings;
                        case AutodiscoverResponseType.RedirectUrl:
                            if (currentHop < AutodiscoverMaxRedirections)
                                {
                                currentHop++;
                                TraceMessage(
                                    TraceFlags.AutodiscoverResponse,
                                    string.Format("Autodiscover service returned redirection URL '{0}'.", settings.RedirectTarget));

                                urls[currentUrlIndex] = new Uri(settings.RedirectTarget);
                                break;
                                }
                            else
                                {
                                throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
                                }
                        case AutodiscoverResponseType.RedirectAddress:
                            if (currentHop < AutodiscoverMaxRedirections)
                                {
                                currentHop++;
                                TraceMessage(
                                    TraceFlags.AutodiscoverResponse,
                                    string.Format("Autodiscover service returned redirection email address '{0}'.", settings.RedirectTarget));

                                // If this email address was already tried, we may have a loop
                                // in SCP lookups. Disable consideration of SCP records.
                                DisableScpLookupIfDuplicateRedirection(settings.RedirectTarget, redirectionEmailAddresses);

                                return InternalGetLegacyUserSettings<TSettings>(
                                                settings.RedirectTarget,
                                                redirectionEmailAddresses,
                                                ref currentHop);
                                }
                            else
                                {
                                throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
                                }
                        case AutodiscoverResponseType.Error:
                            // Don't treat errors from an SCP-based Autodiscover service to be conclusive.
                            // We'll try the next one and record the error for later.
                            if (isScpUrl)
                                {
                                TraceMessage(
                                    TraceFlags.AutodiscoverConfiguration,
                                    "Error returned by Autodiscover service found via SCP, treating as inconclusive.");

                                delayedException = new AutodiscoverRemoteException(Strings.AutodiscoverError, settings.Error);
                                currentUrlIndex++;
                                }
                            else
                                {
                                throw new AutodiscoverRemoteException(Strings.AutodiscoverError, settings.Error);
                                }
                            break;
                        default:
                            EwsUtilities.Assert(
                                false,
                                "Autodiscover.GetConfigurationSettings",
                                "An unexpected error has occurred. This code path should never be reached.");
                            break;
                        }
                    }
                catch (WebException ex)
                    {
                    if (ex.Response != null)
                        {
                        IEwsHttpWebResponse response = HttpWebRequestFactory.CreateExceptionResponse(ex);
                        Uri redirectUrl;
                        if (TryGetRedirectionResponse(response, out redirectUrl))
                            {
                            TraceMessage(
                                TraceFlags.AutodiscoverConfiguration,
                                string.Format("Host returned a redirection to url {0}", redirectUrl));

                            currentHop++;
                            urls[currentUrlIndex] = redirectUrl;
                            }
                        else
                            {
                            ProcessHttpErrorResponse(response, ex);

                            TraceMessage(
                                TraceFlags.AutodiscoverConfiguration,
                                string.Format("{0} failed: {1} ({2})", url, ex.GetType().Name, ex.Message));

                            // The url did not work, let's try the next.
                            currentUrlIndex++;
                            }
                        }
                    else
                        {
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            string.Format("{0} failed: {1} ({2})", url, ex.GetType().Name, ex.Message));

                        // The url did not work, let's try the next.
                        currentUrlIndex++;
                        }
                    }
                catch (XmlException ex)
                    {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("{0} failed: XML parsing error: {1}", url, ex.Message));

                    // The content at the URL wasn't a valid response, let's try the next.
                    currentUrlIndex++;
                    }
                catch (IOException ex)
                    {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("{0} failed: I/O error: {1}", url, ex.Message));

                    // The content at the URL wasn't a valid response, let's try the next.
                    currentUrlIndex++;
                    }
                }
            while (currentUrlIndex < urls.Count);

            // If we got this far it's because none of the URLs we tried have worked. As a next-to-last chance, use GetRedirectUrl to 
            // try to get a redirection URL using an HTTP GET on a non-SSL Autodiscover endpoint. If successful, use this 
            // redirection URL to get the configuration settings for this email address. (This will be a common scenario for 
            // DataCenter deployments).
            Uri redirectionUrl = GetRedirectUrl(domainName);
            if ((redirectionUrl != null) &&
                TryLastChanceHostRedirection<TSettings>(
                    emailAddress,
                    redirectionUrl,
                    out settings))
                {
                return settings;
                }
            else
                {
                // Getting a redirection URL from an HTTP GET failed too. As a last chance, try to get an appropriate SRV Record
                // using DnsQuery. If successful, use this redirection URL to get the configuration settings for this email address.
                redirectionUrl = GetRedirectionUrlFromDnsSrvRecord(domainName);
                if ((redirectionUrl != null) &&
                    TryLastChanceHostRedirection<TSettings>(
                        emailAddress,
                        redirectionUrl,
                        out settings))
                    {
                    return settings;
                    }

                // If there was an earlier exception, throw it.
                else if (delayedException != null)
                    {
                    throw delayedException;
                    }
                else
                    {
                    throw new AutodiscoverLocalException(Strings.AutodiscoverCouldNotBeLocated);
                    }
                }
            }

        /// <summary>
        /// Get an autodiscover SRV record in DNS and construct autodiscover URL.
        /// </summary>
        /// <param name="domainName">Name of the domain.</param>
        /// <returns>Autodiscover URL (may be null if lookup failed)</returns>
        internal Uri GetRedirectionUrlFromDnsSrvRecord(string domainName)
            {
            TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                string.Format("Trying to get Autodiscover host from DNS SRV record for {0}.", domainName));

            string hostname = dnsClient.FindAutodiscoverHostFromSrv(domainName);
            if (!string.IsNullOrEmpty(hostname))
                {
                TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    string.Format("Autodiscover host {0} was returned.", hostname));

                return new Uri(string.Format(AutodiscoverLegacyHttpsUrl, hostname));
                }
            else
                {
                TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    "No matching Autodiscover DNS SRV records were found.");

                return null;
                }
            }

        /// <summary>
        /// Tries to get Autodiscover settings using redirection Url.
        /// </summary>
        /// <typeparam name="TSettings">The type of the settings.</typeparam>
        /// <param name="emailAddress">The email address.</param>
        /// <param name="redirectionUrl">Redirection Url.</param>
        /// <param name="settings">The settings.</param>
        private bool TryLastChanceHostRedirection<TSettings>(
            string emailAddress,
            Uri redirectionUrl,
            out TSettings settings) where TSettings : ConfigurationSettingsBase, new()
            {
            settings = null;

            List<string> redirectionEmailAddresses = new();

            // Bug 60274: Performing a non-SSL HTTP GET to retrieve a redirection URL is potentially unsafe. We allow the caller 
            // to specify delegate to be called to determine whether we are allowed to use the redirection URL. 
            if (CallRedirectionUrlValidationCallback(redirectionUrl.ToString()))
                {
                for (int currentHop = 0; currentHop < AutodiscoverService.AutodiscoverMaxRedirections; currentHop++)
                    {
                    try
                        {
                        settings = GetLegacyUserSettingsAtUrl<TSettings>(emailAddress, redirectionUrl);

                        switch (settings.ResponseType)
                            {
                            case AutodiscoverResponseType.Success:
                                return true;
                            case AutodiscoverResponseType.Error:
                                throw new AutodiscoverRemoteException(Strings.AutodiscoverError, settings.Error);
                            case AutodiscoverResponseType.RedirectAddress:

                                // If this email address was already tried, we may have a loop
                                // in SCP lookups. Disable consideration of SCP records.
                                DisableScpLookupIfDuplicateRedirection(settings.RedirectTarget, redirectionEmailAddresses);

                                settings = InternalGetLegacyUserSettings<TSettings>(
                                    settings.RedirectTarget,
                                    redirectionEmailAddresses,
                                    ref currentHop);
                                return true;
                            case AutodiscoverResponseType.RedirectUrl:
                                try
                                    {
                                    redirectionUrl = new Uri(settings.RedirectTarget);
                                    }
                                catch (UriFormatException)
                                    {
                                    TraceMessage(
                                        TraceFlags.AutodiscoverConfiguration,
                                        string.Format(
                                            "Service returned invalid redirection URL {0}",
                                            settings.RedirectTarget));
                                    return false;
                                    }
                                break;
                            default:
                                string failureMessage = string.Format(
                                    "Autodiscover call at {0} failed with error {1}, target {2}",
                                    redirectionUrl,
                                    settings.ResponseType,
                                    settings.RedirectTarget);
                                TraceMessage(TraceFlags.AutodiscoverConfiguration, failureMessage);
                                return false;
                            }
                        }
                    catch (WebException ex)
                        {
                        if (ex.Response != null)
                            {
                            IEwsHttpWebResponse response = HttpWebRequestFactory.CreateExceptionResponse(ex);
                            if (TryGetRedirectionResponse(response, out redirectionUrl))
                                {
                                TraceMessage(
                                    TraceFlags.AutodiscoverConfiguration,
                                    string.Format("Host returned a redirection to url {0}", redirectionUrl));
                                continue;
                                }
                            else
                                {
                                ProcessHttpErrorResponse(response, ex);
                                }
                            }

                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            string.Format("{0} failed: {1} ({2})", url, ex.GetType().Name, ex.Message));

                        return false;
                        }
                    catch (XmlException ex)
                        {
                        // If the response is malformed, it wasn't a valid Autodiscover endpoint.
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            string.Format("{0} failed: XML parsing error: {1}", redirectionUrl, ex.Message));
                        return false;
                        }
                    catch (IOException ex)
                        {
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            string.Format("{0} failed: I/O error: {1}", redirectionUrl, ex.Message));
                        return false;
                        }
                    }
                }

            return false;
            }

        /// <summary>
        /// Disables SCP lookup if duplicate email address redirection.
        /// </summary>
        /// <param name="emailAddress">The email address to use.</param>
        /// <param name="redirectionEmailAddresses">The list of prior redirection email addresses.</param>
        private void DisableScpLookupIfDuplicateRedirection(string emailAddress, List<string> redirectionEmailAddresses)
            {
            // SMTP addresses are case-insensitive so entries are converted to lower-case.
            emailAddress = emailAddress.ToLowerInvariant();

            if (redirectionEmailAddresses.Contains(emailAddress))
                {
                EnableScpLookup = false;
                }
            else
                {
                redirectionEmailAddresses.Add(emailAddress);
                }
            }

        /// <summary>
        /// Gets user settings from Autodiscover legacy endpoint.
        /// </summary>
        /// <param name="emailAddress">The email address.</param>
        /// <param name="requestedSettings">The requested settings.</param>
        /// <returns>GetUserSettingsResponse</returns>
        internal GetUserSettingsResponse InternalGetLegacyUserSettings(string emailAddress, List<UserSettingName> requestedSettings)
            {
            // Cannot call legacy Autodiscover service with WindowsLive and other WSSecurity-based credentials
            if ((Credentials != null) && (Credentials is WSSecurityBasedCredentials))
                {
                throw new AutodiscoverLocalException(Strings.WLIDCredentialsCannotBeUsedWithLegacyAutodiscover);
                }

            OutlookConfigurationSettings settings = GetLegacyUserSettings<OutlookConfigurationSettings>(emailAddress);

            return settings.ConvertSettings(emailAddress, requestedSettings);
            }
        #endregion

        #region SOAP-based Autodiscover

        /// <summary>
        /// Calls the SOAP Autodiscover service for user settings for a single SMTP address.
        /// </summary>
        /// <param name="smtpAddress">SMTP address.</param>
        /// <param name="requestedSettings">The requested settings.</param>
        /// <returns></returns>
        internal GetUserSettingsResponse InternalGetSoapUserSettings(string smtpAddress, List<UserSettingName> requestedSettings)
            {
            List<string> smtpAddresses = new();
            smtpAddresses.Add(smtpAddress);

            List<string> redirectionEmailAddresses = new();
            redirectionEmailAddresses.Add(smtpAddress.ToLowerInvariant());

            for (int currentHop = 0; currentHop < AutodiscoverService.AutodiscoverMaxRedirections; currentHop++)
                {
                GetUserSettingsResponse response = GetUserSettings(smtpAddresses, requestedSettings)[0];

                switch (response.ErrorCode)
                    {
                    case AutodiscoverErrorCode.RedirectAddress:
                        TraceMessage(
                            TraceFlags.AutodiscoverResponse,
                            string.Format("Autodiscover service returned redirection email address '{0}'.", response.RedirectTarget));

                        smtpAddresses.Clear();
                        smtpAddresses.Add(response.RedirectTarget.ToLowerInvariant());
                        Url = null;
                        Domain = null;

                        // If this email address was already tried, we may have a loop
                        // in SCP lookups. Disable consideration of SCP records.
                        DisableScpLookupIfDuplicateRedirection(response.RedirectTarget, redirectionEmailAddresses);
                        break;

                    case AutodiscoverErrorCode.RedirectUrl:
                        TraceMessage(
                            TraceFlags.AutodiscoverResponse,
                            string.Format("Autodiscover service returned redirection URL '{0}'.", response.RedirectTarget));

                        Url = Credentials.AdjustUrl(new Uri(response.RedirectTarget));
                        break;

                    case AutodiscoverErrorCode.NoError:
                    default:
                        return response;
                    }
                }

            throw new AutodiscoverLocalException(Strings.AutodiscoverCouldNotBeLocated);
            }

        /// <summary>
        /// Gets the user settings using Autodiscover SOAP service.
        /// </summary>
        /// <param name="smtpAddresses">The SMTP addresses of the users.</param>
        /// <param name="settings">The settings.</param>
        /// <returns></returns>
        internal GetUserSettingsResponseCollection GetUserSettings(
            List<string> smtpAddresses,
            List<UserSettingName> settings)
            {
            EwsUtilities.ValidateParam(smtpAddresses, "smtpAddresses");
            EwsUtilities.ValidateParam(settings, "settings");

            return GetSettings<GetUserSettingsResponseCollection, UserSettingName>(
                smtpAddresses,
                settings,
                null,
                InternalGetUserSettings,
                delegate () { return EwsUtilities.DomainFromEmailAddress(smtpAddresses[0]); });
            }

        /// <summary>
        /// Gets user or domain settings using Autodiscover SOAP service.
        /// </summary>
        /// <typeparam name="TGetSettingsResponseCollection">Type of response collection to return.</typeparam>
        /// <typeparam name="TSettingName">Type of setting name.</typeparam>
        /// <param name="identities">Either the domains or the SMTP addresses of the users.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="requestedVersion">Requested version of the Exchange service.</param>
        /// <param name="getSettingsMethod">The method to use.</param>
        /// <param name="getDomainMethod">The method to calculate the domain value.</param>
        /// <returns></returns>
        private TGetSettingsResponseCollection GetSettings<TGetSettingsResponseCollection, TSettingName>(
            List<string> identities,
            List<TSettingName> settings,
            ExchangeVersion? requestedVersion,
            GetSettingsMethod<TGetSettingsResponseCollection, TSettingName> getSettingsMethod,
            System.Func<string> getDomainMethod)
            {
            TGetSettingsResponseCollection response;

            // Autodiscover service only exists in E14 or later.
            if (RequestedServerVersion < MinimumRequestVersionForAutoDiscoverSoapService)
                {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.AutodiscoverServiceIncompatibleWithRequestVersion,
                        MinimumRequestVersionForAutoDiscoverSoapService));
                }

            // If Url is specified, call service directly.
            if (Url != null)
                {
                Uri autodiscoverUrl = Url;

                response = getSettingsMethod(
                            identities,
                            settings,
                            requestedVersion,
                            ref autodiscoverUrl);

                Url = autodiscoverUrl;
                return response;
                }

            // If Domain is specified, determine endpoint Url and call service.
            else if (!string.IsNullOrEmpty(Domain))
                {
                Uri autodiscoverUrl = GetAutodiscoverEndpointUrl(Domain);
                response = getSettingsMethod(
                                identities,
                                settings,
                                requestedVersion,
                                ref autodiscoverUrl);

                // If we got this far, response was successful, set Url.
                Url = autodiscoverUrl;
                return response;
                }

            // No Url or Domain specified, need to figure out which endpoint(s) to try.
            else
                {
                // Assume caller is not inside the Intranet, regardless of whether SCP Urls 
                // were returned or not. SCP Urls are only relevent if one of them returns
                // valid Autodiscover settings.
                IsExternal = true;

                Uri autodiscoverUrl;

                string domainName = getDomainMethod();
                int scpHostCount;
                List<string> hosts = GetAutodiscoverServiceHosts(domainName, out scpHostCount);

                if (hosts.Count == 0)
                    {
                    throw new ServiceValidationException(Strings.AutodiscoverServiceRequestRequiresDomainOrUrl);
                    }

                for (int currentHostIndex = 0; currentHostIndex < hosts.Count; currentHostIndex++)
                    {
                    string host = hosts[currentHostIndex];
                    bool isScpHost = currentHostIndex < scpHostCount;

                    if (TryGetAutodiscoverEndpointUrl(host, out autodiscoverUrl))
                        {
                        try
                            {
                            response = getSettingsMethod(
                                            identities,
                                            settings,
                                            requestedVersion,
                                            ref autodiscoverUrl);

                            // If we got this far, the response was successful, set Url.
                            Url = autodiscoverUrl;

                            // Not external if Autodiscover endpoint found via SCP returned the settings.
                            if (isScpHost)
                                {
                                IsExternal = false;
                                }

                            return response;
                            }
                        catch (AutodiscoverResponseException)
                            {
                            // skip
                            }
                        catch (ServiceRequestException)
                            {
                            // skip
                            }
                        }
                    }

                // Next-to-last chance: try unauthenticated GET over HTTP to be redirected to appropriate service endpoint.
                autodiscoverUrl = GetRedirectUrl(domainName);
                if ((autodiscoverUrl != null) &&
                    CallRedirectionUrlValidationCallback(autodiscoverUrl.ToString()) &&
                    TryGetAutodiscoverEndpointUrl(autodiscoverUrl.Host, out autodiscoverUrl))
                    {
                    response = getSettingsMethod(
                                    identities,
                                    settings,
                                    requestedVersion,
                                    ref autodiscoverUrl);

                    // If we got this far, the response was successful, set Url.
                    Url = autodiscoverUrl;

                    return response;
                    }

                // Last Chance: try to read autodiscover SRV Record from DNS. If we find one, use
                // the hostname returned to construct an Autodiscover endpoint URL.
                autodiscoverUrl = GetRedirectionUrlFromDnsSrvRecord(domainName);
                if ((autodiscoverUrl != null) &&
                    CallRedirectionUrlValidationCallback(autodiscoverUrl.ToString()) &&
                        TryGetAutodiscoverEndpointUrl(autodiscoverUrl.Host, out autodiscoverUrl))
                    {
                    response = getSettingsMethod(
                                    identities,
                                    settings,
                                    requestedVersion,
                                    ref autodiscoverUrl);

                    // If we got this far, the response was successful, set Url.
                    Url = autodiscoverUrl;

                    return response;
                    }
                else
                    {
                    throw new AutodiscoverLocalException(Strings.AutodiscoverCouldNotBeLocated);
                    }
                }
            }

        /// <summary>
        /// Gets settings for one or more users.
        /// </summary>
        /// <param name="smtpAddresses">The SMTP addresses of the users.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="requestedVersion">Requested version of the Exchange service.</param>
        /// <param name="autodiscoverUrl">The autodiscover URL.</param>
        /// <returns>GetUserSettingsResponse collection.</returns>
        private GetUserSettingsResponseCollection InternalGetUserSettings(
            List<string> smtpAddresses,
            List<UserSettingName> settings,
            ExchangeVersion? requestedVersion,
            ref Uri autodiscoverUrl)
            {
            // The response to GetUserSettings can be a redirection. Execute GetUserSettings until we get back 
            // a valid response or we've followed too many redirections.
            for (int currentHop = 0; currentHop < AutodiscoverService.AutodiscoverMaxRedirections; currentHop++)
                {
                GetUserSettingsRequest request = new(this, autodiscoverUrl);
                request.SmtpAddresses = smtpAddresses;
                request.Settings = settings;
                GetUserSettingsResponseCollection response = request.Execute();

                // Did we get redirected?
                if (response.ErrorCode == AutodiscoverErrorCode.RedirectUrl && response.RedirectionUrl != null)
                    {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("Request to {0} returned redirection to {1}", autodiscoverUrl.ToString(), response.RedirectionUrl));

                    // this url need be brought back to the caller.
                    //
                    autodiscoverUrl = response.RedirectionUrl;
                    }
                else
                    {
                    return response;
                    }
                }

            TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                string.Format("Maximum number of redirection hops {0} exceeded", AutodiscoverMaxRedirections));

            throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
            }

        /// <summary>
        /// Gets the domain settings using Autodiscover SOAP service.
        /// </summary>
        /// <param name="domains">The domains.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="requestedVersion">Requested version of the Exchange service.</param>
        /// <returns>GetDomainSettingsResponse collection.</returns>
        internal GetDomainSettingsResponseCollection GetDomainSettings(
            List<string> domains,
            List<DomainSettingName> settings,
            ExchangeVersion? requestedVersion)
            {
            EwsUtilities.ValidateParam(domains, "domains");
            EwsUtilities.ValidateParam(settings, "settings");

            return GetSettings<GetDomainSettingsResponseCollection, DomainSettingName>(
                domains,
                settings,
                requestedVersion,
                InternalGetDomainSettings,
                delegate () { return domains[0]; });
            }

        /// <summary>
        /// Gets settings for one or more domains.
        /// </summary>
        /// <param name="domains">The domains.</param>
        /// <param name="settings">The settings.</param>
        /// <param name="requestedVersion">Requested version of the Exchange service.</param>
        /// <param name="autodiscoverUrl">The autodiscover URL.</param>
        /// <returns>GetDomainSettingsResponse collection.</returns>
        private GetDomainSettingsResponseCollection InternalGetDomainSettings(
            List<string> domains,
            List<DomainSettingName> settings,
            ExchangeVersion? requestedVersion,
            ref Uri autodiscoverUrl)
            {
            // The response to GetDomainSettings can be a redirection. Execute GetDomainSettings until we get back 
            // a valid response or we've followed too many redirections.
            for (int currentHop = 0; currentHop < AutodiscoverService.AutodiscoverMaxRedirections; currentHop++)
                {
                GetDomainSettingsRequest request = new(this, autodiscoverUrl);
                request.Domains = domains;
                request.Settings = settings;
                request.RequestedVersion = requestedVersion;
                GetDomainSettingsResponseCollection response = request.Execute();

                // Did we get redirected?
                if (response.ErrorCode == AutodiscoverErrorCode.RedirectUrl && response.RedirectionUrl != null)
                    {
                    autodiscoverUrl = response.RedirectionUrl;
                    }
                else
                    {
                    return response;
                    }
                }

            TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                string.Format("Maximum number of redirection hops {0} exceeded", AutodiscoverMaxRedirections));

            throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
            }

        /// <summary>
        /// Gets the autodiscover endpoint URL.
        /// </summary>
        /// <param name="host">The host.</param>
        /// <returns></returns>
        private Uri GetAutodiscoverEndpointUrl(string host)
            {
            Uri autodiscoverUrl;

            if (TryGetAutodiscoverEndpointUrl(host, out autodiscoverUrl))
                {
                return autodiscoverUrl;
                }
            else
                {
                throw new AutodiscoverLocalException(Strings.NoSoapOrWsSecurityEndpointAvailable);
                }
            }

        /// <summary>
        /// Tries the get Autodiscover Service endpoint URL.
        /// </summary>
        /// <param name="host">The host.</param>
        /// <param name="url">The URL.</param>
        /// <returns></returns>
        private bool TryGetAutodiscoverEndpointUrl(string host, out Uri url)
            {
            url = null;

            AutodiscoverEndpoints endpoints;
            if (TryGetEnabledEndpointsForHost(ref host, out endpoints))
                {
                url = new Uri(string.Format(AutodiscoverSoapHttpsUrl, host));

                // Make sure that at least one of the non-legacy endpoints is available.
                if (((endpoints & AutodiscoverEndpoints.Soap) != AutodiscoverEndpoints.Soap) &&
                    ((endpoints & AutodiscoverEndpoints.WsSecurity) != AutodiscoverEndpoints.WsSecurity) &&
                    ((endpoints & AutodiscoverEndpoints.WSSecuritySymmetricKey) != AutodiscoverEndpoints.WSSecuritySymmetricKey) &&
                    ((endpoints & AutodiscoverEndpoints.WSSecurityX509Cert) != AutodiscoverEndpoints.WSSecurityX509Cert) &&
                    ((endpoints & AutodiscoverEndpoints.OAuth) != AutodiscoverEndpoints.OAuth))
                    {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("No Autodiscover endpoints are available  for host {0}", host));

                    return false;
                    }

                // If we have WLID credentials, make sure that we have a WS-Security endpoint
                if (Credentials is WindowsLiveCredentials)
                    {
                    if ((endpoints & AutodiscoverEndpoints.WsSecurity) != AutodiscoverEndpoints.WsSecurity)
                        {
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            string.Format("No Autodiscover WS-Security endpoint is available for host {0}", host));

                        return false;
                        }
                    else
                        {
                        url = new Uri(string.Format(AutodiscoverSoapWsSecurityHttpsUrl, host));
                        }
                    }
                else if (Credentials is PartnerTokenCredentials)
                    {
                    if ((endpoints & AutodiscoverEndpoints.WSSecuritySymmetricKey) != AutodiscoverEndpoints.WSSecuritySymmetricKey)
                        {
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            string.Format("No Autodiscover WS-Security/SymmetricKey endpoint is available for host {0}", host));

                        return false;
                        }
                    else
                        {
                        url = new Uri(string.Format(AutodiscoverSoapWsSecuritySymmetricKeyHttpsUrl, host));
                        }
                    }
                else if (Credentials is X509CertificateCredentials)
                    {
                    if ((endpoints & AutodiscoverEndpoints.WSSecurityX509Cert) != AutodiscoverEndpoints.WSSecurityX509Cert)
                        {
                        TraceMessage(
                            TraceFlags.AutodiscoverConfiguration,
                            string.Format("No Autodiscover WS-Security/X509Cert endpoint is available for host {0}", host));

                        return false;
                        }
                    else
                        {
                        url = new Uri(string.Format(AutodiscoverSoapWsSecurityX509CertHttpsUrl, host));
                        }
                    }
                else if (Credentials is OAuthCredentials)
                    {
                    // If the credential is OAuthCredentials, no matter whether we have
                    // the corresponding x-header, we will go with OAuth. 
                    url = new Uri(string.Format(AutodiscoverSoapHttpsUrl, host));
                    }

                return true;
                }
            else
                {
                TraceMessage(
                    TraceFlags.AutodiscoverConfiguration,
                    string.Format("No Autodiscover endpoints are available for host {0}", host));

                return false;
                }
            }

        /// <summary>
        /// Defaults the get autodiscover service urls for domain.
        /// </summary>
        /// <param name="domainName">Name of the domain.</param>
        /// <returns></returns>
        private ICollection<string> DefaultGetScpUrlsForDomain(string domainName)
            {
            DirectoryHelper helper = new(this);
            return helper.GetAutodiscoverScpUrlsForDomain(domainName);
            }

        /// <summary>
        /// Gets the list of autodiscover service URLs.
        /// </summary>
        /// <param name="domainName">Domain name.</param>
        /// <param name="scpHostCount">Count of hosts found via SCP lookup.</param>
        /// <returns>List of Autodiscover URLs.</returns>
        internal List<Uri> GetAutodiscoverServiceUrls(string domainName, out int scpHostCount)
            {
            List<Uri> urls = new();

            if (enableScpLookup)
                {
                // Get SCP URLs
                Func<string, ICollection<string>> callback = GetScpUrlsForDomainCallback ?? DefaultGetScpUrlsForDomain;
                ICollection<string> scpUrls = callback(domainName);
                foreach (string str in scpUrls)
                    {
                    urls.Add(new Uri(str));
                    }
                }

            scpHostCount = urls.Count;

            // As a fallback, add autodiscover URLs base on the domain name.
            urls.Add(new Uri(string.Format(AutodiscoverLegacyHttpsUrl, domainName)));
            urls.Add(new Uri(string.Format(AutodiscoverLegacyHttpsUrl, "autodiscover." + domainName)));

            return urls;
            }

        /// <summary>
        /// Gets the list of autodiscover service hosts.
        /// </summary>
        /// <param name="domainName">Name of the domain.</param>
        /// <param name="scpHostCount">Count of SCP hosts that were found.</param>
        /// <returns>List of host names.</returns>
        internal List<string> GetAutodiscoverServiceHosts(string domainName, out int scpHostCount)
            {
            List<string> serviceHosts = new();
            foreach (Uri url in GetAutodiscoverServiceUrls(domainName, out scpHostCount))
                {
                serviceHosts.Add(url.Host);
                }

            return serviceHosts;
            }

        /// <summary>
        /// Gets the enabled autodiscover endpoints on a specific host.
        /// </summary>
        /// <param name="host">The host.</param>
        /// <param name="endpoints">Endpoints found for host.</param>
        /// <returns>Flags indicating which endpoints are enabled.</returns>
        private bool TryGetEnabledEndpointsForHost(ref string host, out AutodiscoverEndpoints endpoints)
            {
            TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                string.Format("Determining which endpoints are enabled for host {0}", host));

            // We may get redirected to another host. And therefore need to limit the number
            // of redirections we'll tolerate.
            for (int currentHop = 0; currentHop < AutodiscoverMaxRedirections; currentHop++)
                {
                Uri autoDiscoverUrl = new(string.Format(AutodiscoverLegacyHttpsUrl, host));

                endpoints = AutodiscoverEndpoints.None;

                IEwsHttpWebRequest request = HttpWebRequestFactory.CreateRequest(autoDiscoverUrl);

                request.Method = "GET";
                request.AllowAutoRedirect = false;
                request.PreAuthenticate = false;
                request.UseDefaultCredentials = false;

                IEwsHttpWebResponse response = null;

                try
                    {
                    response = request.GetResponse();
                    }
                catch (WebException ex)
                    {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("Request error: {0}", ex.Message));

                    // The exception response factory requires a valid HttpWebResponse, 
                    // but there will be no web response if the web request couldn't be
                    // actually be issued (e.g. due to DNS error).
                    if (ex.Response != null)
                        {
                        response = HttpWebRequestFactory.CreateExceptionResponse(ex);
                        }
                    }
                catch (IOException ex)
                    {
                    TraceMessage(
                        TraceFlags.AutodiscoverConfiguration,
                        string.Format("I/O error: {0}", ex.Message));
                    }

                if (response != null)
                    {
                    using (response)
                        {
                        Uri redirectUrl;
                        if (TryGetRedirectionResponse(response, out redirectUrl))
                            {
                            TraceMessage(
                                TraceFlags.AutodiscoverConfiguration,
                                string.Format("Host returned redirection to host '{0}'", redirectUrl.Host));

                            host = redirectUrl.Host;
                            }
                        else
                            {
                            endpoints = GetEndpointsFromHttpWebResponse(response);

                            TraceMessage(
                                TraceFlags.AutodiscoverConfiguration,
                                string.Format("Host returned enabled endpoint flags: {0}", endpoints));

                            return true;
                            }
                        }
                    }
                else
                    {
                    return false;
                    }
                }

            TraceMessage(
                TraceFlags.AutodiscoverConfiguration,
                string.Format("Maximum number of redirection hops {0} exceeded", AutodiscoverMaxRedirections));

            throw new AutodiscoverLocalException(Strings.MaximumRedirectionHopsExceeded);
            }

        /// <summary>
        /// Gets the endpoints from HTTP web response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <returns>Endpoints enabled.</returns>
        private AutodiscoverEndpoints GetEndpointsFromHttpWebResponse(IEwsHttpWebResponse response)
            {
            AutodiscoverEndpoints endpoints = AutodiscoverEndpoints.Legacy;
            if (!string.IsNullOrEmpty(response.Headers[AutodiscoverSoapEnabledHeaderName]))
                {
                endpoints |= AutodiscoverEndpoints.Soap;
                }
            if (!string.IsNullOrEmpty(response.Headers[AutodiscoverWsSecurityEnabledHeaderName]))
                {
                endpoints |= AutodiscoverEndpoints.WsSecurity;
                }
            if (!string.IsNullOrEmpty(response.Headers[AutodiscoverWsSecuritySymmetricKeyEnabledHeaderName]))
                {
                endpoints |= AutodiscoverEndpoints.WSSecuritySymmetricKey;
                }
            if (!string.IsNullOrEmpty(response.Headers[AutodiscoverWsSecurityX509CertEnabledHeaderName]))
                {
                endpoints |= AutodiscoverEndpoints.WSSecurityX509Cert;
                }
            if (!string.IsNullOrEmpty(response.Headers[AutodiscoverOAuthEnabledHeaderName]))
                {
                endpoints |= AutodiscoverEndpoints.OAuth;
                }
            return endpoints;
            }

        /// <summary>
        /// Traces the response.
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="memoryStream">The response content in a MemoryStream.</param>
        internal void TraceResponse(IEwsHttpWebResponse response, MemoryStream memoryStream)
            {
            ProcessHttpResponseHeaders(TraceFlags.AutodiscoverResponseHttpHeaders, response);

            if (TraceEnabled)
                {
                if (!string.IsNullOrEmpty(response.ContentType) &&
                    (response.ContentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) ||
                     response.ContentType.StartsWith("application/soap", StringComparison.OrdinalIgnoreCase)))
                    {
                    TraceXml(TraceFlags.AutodiscoverResponse, memoryStream);
                    }
                else
                    {
                    TraceMessage(TraceFlags.AutodiscoverResponse, "Non-textual response");
                    }
                }
            }

        #endregion

        #region Utilities
        /// <summary>
        /// Creates an HttpWebRequest instance and initializes it with the appropriate parameters,
        /// based on the configuration of this service object.
        /// </summary>
        /// <param name="url">The URL that the HttpWebRequest should target.</param>
        internal IEwsHttpWebRequest PrepareHttpWebRequestForUrl(Uri url)
            {
            return PrepareHttpWebRequestForUrl(
                            url,
                            false,      // acceptGzipEncoding
                            false);     // allowAutoRedirect
            }

        /// <summary>
        /// Calls the redirection URL validation callback.
        /// </summary>
        /// <param name="redirectionUrl">The redirection URL.</param>
        /// <remarks>
        /// If the redirection URL validation callback is null, use the default callback which
        /// does not allow following any redirections.
        /// </remarks>
        /// <returns>True if redirection should be followed.</returns>
        private bool CallRedirectionUrlValidationCallback(string redirectionUrl)
            {
            AutodiscoverRedirectionUrlValidationCallback callback = (RedirectionUrlValidationCallback == null)
                                                                        ? DefaultAutodiscoverRedirectionUrlValidationCallback
                                                                        : RedirectionUrlValidationCallback;
            return callback(redirectionUrl);
            }

        /// <summary>
        /// Processes an HTTP error response.
        /// </summary>
        /// <param name="httpWebResponse">The HTTP web response.</param>
        /// <param name="webException">The web exception.</param>
        internal override void ProcessHttpErrorResponse(IEwsHttpWebResponse httpWebResponse, WebException webException)
            {
            InternalProcessHttpErrorResponse(
                httpWebResponse,
                webException,
                TraceFlags.AutodiscoverResponseHttpHeaders,
                TraceFlags.AutodiscoverResponse);
            }
        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        public AutodiscoverService()
            : this(ExchangeVersion.Exchange2010)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="requestedServerVersion">The requested server version.</param>
        public AutodiscoverService(ExchangeVersion requestedServerVersion)
            : this(null, null, requestedServerVersion)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="domain">The domain that will be used to determine the URL of the service.</param>
        public AutodiscoverService(string domain)
            : this(null, domain)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="domain">The domain that will be used to determine the URL of the service.</param>
        /// <param name="requestedServerVersion">The requested server version.</param>
        public AutodiscoverService(string domain, ExchangeVersion requestedServerVersion)
            : this(null, domain, requestedServerVersion)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="url">The URL of the service.</param>
        public AutodiscoverService(Uri url)
            : this(url, url.Host)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="url">The URL of the service.</param>
        /// <param name="requestedServerVersion">The requested server version.</param>
        public AutodiscoverService(Uri url, ExchangeVersion requestedServerVersion)
            : this(url, url.Host, requestedServerVersion)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="url">The URL of the service.</param>
        /// <param name="domain">The domain that will be used to determine the URL of the service.</param>
        internal AutodiscoverService(Uri url, string domain)
            : base()
            {
            EwsUtilities.ValidateDomainNameAllowNull(domain, "domain");

            this.url = url;
            this.domain = domain;
            dnsClient = new AutodiscoverDnsClient(this);
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="url">The URL of the service.</param>
        /// <param name="domain">The domain that will be used to determine the URL of the service.</param>
        /// <param name="requestedServerVersion">The requested server version.</param>
        internal AutodiscoverService(
            Uri url,
            string domain,
            ExchangeVersion requestedServerVersion)
            : base(requestedServerVersion)
            {
            EwsUtilities.ValidateDomainNameAllowNull(domain, "domain");

            this.url = url;
            this.domain = domain;
            dnsClient = new AutodiscoverDnsClient(this);
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="service">The other service.</param>
        /// <param name="requestedServerVersion">The requested server version.</param>
        internal AutodiscoverService(ExchangeServiceBase service, ExchangeVersion requestedServerVersion)
            : base(service, requestedServerVersion)
            {
            dnsClient = new AutodiscoverDnsClient(this);
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="AutodiscoverService"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal AutodiscoverService(ExchangeServiceBase service)
            : this(service, service.RequestedServerVersion)
            {
            }

        #endregion

        #region Public Methods
        /// <summary>
        /// Retrieves the specified settings for single SMTP address.
        /// </summary>
        /// <param name="userSmtpAddress">The SMTP addresses of the user.</param>
        /// <param name="userSettingNames">The user setting names.</param>
        /// <returns>A UserResponse object containing the requested settings for the specified user.</returns>
        /// <remarks>
        /// This method handles will run the entire Autodiscover "discovery" algorithm and will follow address and URL redirections.
        /// </remarks>
        public GetUserSettingsResponse GetUserSettings(
            string userSmtpAddress,
            params UserSettingName[] userSettingNames)
            {
            List<UserSettingName> requestedSettings = new(userSettingNames);

            if (string.IsNullOrEmpty(userSmtpAddress))
                {
                throw new ServiceValidationException(Strings.InvalidAutodiscoverSmtpAddress);
                }

            if (requestedSettings.Count == 0)
                {
                throw new ServiceValidationException(Strings.InvalidAutodiscoverSettingsCount);
                }

            if (RequestedServerVersion < MinimumRequestVersionForAutoDiscoverSoapService)
                {
                return InternalGetLegacyUserSettings(userSmtpAddress, requestedSettings);
                }
            else
                {
                return InternalGetSoapUserSettings(userSmtpAddress, requestedSettings);
                }
            }

        /// <summary>
        /// Retrieves the specified settings for a set of users.
        /// </summary>
        /// <param name="userSmtpAddresses">The SMTP addresses of the users.</param>
        /// <param name="userSettingNames">The user setting names.</param>
        /// <returns>A GetUserSettingsResponseCollection object containing the responses for each individual user.</returns>
        public GetUserSettingsResponseCollection GetUsersSettings(
            IEnumerable<string> userSmtpAddresses,
            params UserSettingName[] userSettingNames)
            {
            if (RequestedServerVersion < MinimumRequestVersionForAutoDiscoverSoapService)
                {
                throw new ServiceVersionException(
                    string.Format(Strings.AutodiscoverServiceIncompatibleWithRequestVersion, MinimumRequestVersionForAutoDiscoverSoapService));
                }

            List<string> smtpAddresses = new(userSmtpAddresses);
            List<UserSettingName> settings = new(userSettingNames);

            return GetUserSettings(smtpAddresses, settings);
            }

        /// <summary>
        /// Retrieves the specified settings for a domain.
        /// </summary>
        /// <param name="domain">The domain.</param>
        /// <param name="requestedVersion">Requested version of the Exchange service.</param>
        /// <param name="domainSettingNames">The domain setting names.</param>
        /// <returns>A DomainResponse object containing the requested settings for the specified domain.</returns>
        public GetDomainSettingsResponse GetDomainSettings(
            string domain,
            ExchangeVersion? requestedVersion,
            params DomainSettingName[] domainSettingNames)
            {
            List<string> domains = new(1);
            domains.Add(domain);
            List<DomainSettingName> settings = new(domainSettingNames);
            return GetDomainSettings(domains, settings, requestedVersion)[0];
            }

        /// <summary>
        /// Retrieves the specified settings for a set of domains.
        /// </summary>
        /// <param name="domains">The SMTP addresses of the domains.</param>
        /// <param name="requestedVersion">Requested version of the Exchange service.</param>
        /// <param name="domainSettingNames">The domain setting names.</param>
        /// <returns>A GetDomainSettingsResponseCollection object containing the responses for each individual domain.</returns>
        public GetDomainSettingsResponseCollection GetDomainSettings(
            IEnumerable<string> domains,
            ExchangeVersion? requestedVersion,
            params DomainSettingName[] domainSettingNames)
            {
            List<DomainSettingName> settings = new(domainSettingNames);

            return GetDomainSettings(new List<string>(domains), settings, requestedVersion);
            }

        /// <summary>
        /// Try to get the partner access information for the given target tenant.
        /// </summary>
        /// <param name="targetTenantDomain">The target domain or user email address.</param>
        /// <param name="partnerAccessCredentials">The partner access credentials.</param>
        /// <param name="targetTenantAutodiscoverUrl">The autodiscover url for the given tenant.</param>
        /// <returns>True if the partner access information was retrieved, false otherwise.</returns>
        public bool TryGetPartnerAccess(
            string targetTenantDomain,
            out ExchangeCredentials partnerAccessCredentials,
            out Uri targetTenantAutodiscoverUrl)
            {
            EwsUtilities.ValidateNonBlankStringParam(targetTenantDomain, "targetTenantDomain");

            // the user should set the url to its own tenant's autodiscover url.
            // 
            if (Url == null)
                {
                throw new ServiceValidationException(Strings.PartnerTokenRequestRequiresUrl);
                }

            if (RequestedServerVersion < ExchangeVersion.Exchange2010_SP1)
                {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.PartnerTokenIncompatibleWithRequestVersion,
                        ExchangeVersion.Exchange2010_SP1));
                }

            partnerAccessCredentials = null;
            targetTenantAutodiscoverUrl = null;

            string smtpAddress = targetTenantDomain;
            if (!smtpAddress.Contains("@"))
                {
                smtpAddress = "SystemMailbox{e0dc1c29-89c3-4034-b678-e6c29d823ed9}@" + targetTenantDomain;
                }

            GetUserSettingsRequest request = new(this, Url, true /* expectPartnerToken */);
            request.SmtpAddresses = new List<string>(new[] { smtpAddress });
            request.Settings = new List<UserSettingName>(new[] { UserSettingName.ExternalEwsUrl });

            GetUserSettingsResponseCollection response = null;
            try
                {
                response = request.Execute();
                }
            catch (ServiceRequestException)
                {
                return false;
                }
            catch (ServiceRemoteException)
                {
                return false;
                }

            if (string.IsNullOrEmpty(request.PartnerToken)
                || string.IsNullOrEmpty(request.PartnerTokenReference))
                {
                return false;
                }

            if (response.ErrorCode == AutodiscoverErrorCode.NoError)
                {
                GetUserSettingsResponse firstResponse = response.Responses[0];
                if (firstResponse.ErrorCode == AutodiscoverErrorCode.NoError)
                    {
                    targetTenantAutodiscoverUrl = Url;
                    }
                else if (firstResponse.ErrorCode == AutodiscoverErrorCode.RedirectUrl)
                    {
                    targetTenantAutodiscoverUrl = new Uri(firstResponse.RedirectTarget);
                    }
                else
                    {
                    return false;
                    }
                }
            else
                {
                return false;
                }

            partnerAccessCredentials = new PartnerTokenCredentials(
                request.PartnerToken,
                request.PartnerTokenReference);

            targetTenantAutodiscoverUrl = partnerAccessCredentials.AdjustUrl(
                targetTenantAutodiscoverUrl);

            return true;
            }
        #endregion

        #region Properties
        /// <summary>
        /// Gets or sets the domain this service is bound to. When this property is set, the domain
        /// name is used to automatically determine the Autodiscover service URL.
        /// </summary>
        public string Domain
            {
            get { return domain; }
            set
                {
                EwsUtilities.ValidateDomainNameAllowNull(value, "Domain");

                // If Domain property is set to non-null value, Url property is nulled.
                if (value != null)
                    {
                    url = null;
                    }
                domain = value;
                }
            }

        /// <summary>
        /// Gets or sets the URL this service is bound to.
        /// </summary>
        public Uri Url
            {
            get { return url; }
            set
                {
                // If Url property is set to non-null value, Domain property is set to host portion of Url.
                if (value != null)
                    {
                    domain = value.Host;
                    }
                url = value;
                }
            }

        /// <summary>
        /// Gets a value indicating whether the Autodiscover service that URL points to is internal (inside the corporate network)
        /// or external (outside the corporate network).
        /// </summary>
        /// <remarks>
        /// IsExternal is null in the following cases:
        /// - This instance has been created with a domain name and no method has been called,
        /// - This instance has been created with a URL.
        /// </remarks>
        public bool? IsExternal
            {
            get { return isExternal; }
            internal set { isExternal = value; }
            }

        /// <summary>
        /// Gets or sets the redirection URL validation callback.
        /// </summary>
        /// <value>The redirection URL validation callback.</value>
        public AutodiscoverRedirectionUrlValidationCallback RedirectionUrlValidationCallback
            {
            get { return redirectionUrlValidationCallback; }
            set { redirectionUrlValidationCallback = value; }
            }

        /// <summary>
        /// Gets or sets the DNS server address.
        /// </summary>
        /// <value>The DNS server address.</value>
        internal IPAddress DnsServerAddress
            {
            get { return dnsServerAddress; }
            set { dnsServerAddress = value; }
            }

        /// <summary>
        /// Gets or sets a value indicating whether the AutodiscoverService should perform SCP (ServiceConnectionPoint) record lookup when determining
        /// the Autodiscover service URL.
        /// </summary>
        public bool EnableScpLookup
            {
            get { return enableScpLookup; }
            set { enableScpLookup = value; }
            }

        /// <summary>
        /// Gets or sets the delegate used to resolve Autodiscover SCP urls for a specified domain.
        /// </summary>
        public Func<string, ICollection<string>> GetScpUrlsForDomainCallback
            {
            get;
            set;
            }

        #endregion
        }
    }