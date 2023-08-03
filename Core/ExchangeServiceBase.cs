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
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Net;
    using System.Security.Cryptography;
    using System.Xml;

    /// <summary>
    /// Represents an abstract binding to an Exchange Service.
    /// </summary>
    public abstract class ExchangeServiceBase
        {
        #region Const members
        private static readonly object lockObj = new();

        private readonly ExchangeVersion requestedServerVersion = ExchangeVersion.Exchange2013_SP1;

        /// <summary>
        /// Special HTTP status code that indicates that the account is locked.
        /// </summary>
        internal const HttpStatusCode AccountIsLocked = (HttpStatusCode)456;

        /// <summary>
        /// The binary secret.
        /// </summary>
        private static byte[] binarySecret;
        #endregion

        #region Static members

        /// <summary>
        /// Default UserAgent
        /// </summary>
        private static string defaultUserAgent = "ExchangeServicesClient/" + EwsUtilities.BuildVersion;

        #endregion

        #region Fields        

        /// <summary>
        /// Occurs when the http response headers of a server call is captured.
        /// </summary>
        public event ResponseHeadersCapturedHandler OnResponseHeadersCaptured;

        private ExchangeCredentials credentials;
        private bool useDefaultCredentials;
        private int timeout = 100000;
        private bool traceEnabled;
        private bool sendClientLatencies = true;
        private TraceFlags traceFlags = TraceFlags.All;
        private ITraceListener traceListener = new EwsTraceListener();
        private bool preAuthenticate;
        private string userAgent = ExchangeService.defaultUserAgent;
        private bool acceptGzipEncoding = true;
        private bool keepAlive = true;
        private string connectionGroupName;
        private string clientRequestId;
        private bool returnClientRequestId;
        private CookieContainer cookieContainer = new();
        private TimeZoneInfo timeZone;
        private TimeZoneDefinition timeZoneDefinition;
        private ExchangeServerInfo serverInfo;
        private IWebProxy webProxy;
        private IDictionary<string, string> httpHeaders = new Dictionary<string, string>();
        private IDictionary<string, string> httpResponseHeaders = new Dictionary<string, string>();
        private IEwsHttpWebRequestFactory ewsHttpWebRequestFactory = new EwsHttpWebRequestFactory();
        #endregion

        #region Event handlers

        /// <summary>
        /// Calls the custom SOAP header serialization event handlers, if defined.
        /// </summary>
        /// <param name="writer">The XmlWriter to which to write the custom SOAP headers.</param>
        internal void DoOnSerializeCustomSoapHeaders(XmlWriter writer)
            {
            EwsUtilities.Assert(
                writer != null,
                "ExchangeService.DoOnSerializeCustomSoapHeaders",
                "writer is null");

            if (OnSerializeCustomSoapHeaders != null)
                {
                OnSerializeCustomSoapHeaders(writer);
                }
            }

        #endregion

        #region Utilities

        /// <summary>
        /// Creates an HttpWebRequest instance and initializes it with the appropriate parameters,
        /// based on the configuration of this service object.
        /// </summary>
        /// <param name="url">The URL that the HttpWebRequest should target.</param>
        /// <param name="acceptGzipEncoding">If true, ask server for GZip compressed content.</param>
        /// <param name="allowAutoRedirect">If true, redirection responses will be automatically followed.</param>
        /// <returns>A initialized instance of HttpWebRequest.</returns>
        internal IEwsHttpWebRequest PrepareHttpWebRequestForUrl(
            Uri url,
            bool acceptGzipEncoding,
            bool allowAutoRedirect)
            {
            // Verify that the protocol is something that we can handle
            if ((url.Scheme != Uri.UriSchemeHttp) && (url.Scheme != Uri.UriSchemeHttps))
                {
                throw new ServiceLocalException(string.Format(Strings.UnsupportedWebProtocol, url.Scheme));
                }

            IEwsHttpWebRequest request = HttpWebRequestFactory.CreateRequest(url);

            request.PreAuthenticate = PreAuthenticate;
            request.Timeout = Timeout;
            SetContentType(request);
            request.Method = "POST";
            request.UserAgent = UserAgent;
            request.AllowAutoRedirect = allowAutoRedirect;
            request.CookieContainer = CookieContainer;
            request.KeepAlive = keepAlive;
            request.ConnectionGroupName = connectionGroupName;

            if (acceptGzipEncoding)
                {
                request.Headers.Add(HttpRequestHeader.AcceptEncoding, "gzip,deflate");
                }

            if (!string.IsNullOrEmpty(clientRequestId))
                {
                request.Headers.Add("client-request-id", clientRequestId);
                if (returnClientRequestId)
                    {
                    request.Headers.Add("return-client-request-id", "true");
                    }
                }

            if (webProxy != null)
                {
                request.Proxy = webProxy;
                }

            if (HttpHeaders.Count > 0)
                {
                HttpHeaders.ForEach((kv) => request.Headers.Add(kv.Key, kv.Value));
                }

            request.UseDefaultCredentials = UseDefaultCredentials;
            if (!request.UseDefaultCredentials)
                {
                ExchangeCredentials serviceCredentials = Credentials;
                if (serviceCredentials == null)
                    {
                    throw new ServiceLocalException(Strings.CredentialsRequired);
                    }

                // Make sure that credentials have been authenticated if required
                serviceCredentials.PreAuthenticate();

                // Apply credentials to the request
                serviceCredentials.PrepareWebRequest(request);
                }

            httpResponseHeaders.Clear();

            return request;
            }

        internal virtual void SetContentType(IEwsHttpWebRequest request)
            {
            request.ContentType = "text/xml; charset=utf-8";
            request.Accept = "text/xml";
            }

        /// <summary>
        /// Processes an HTTP error response
        /// </summary>
        /// <param name="httpWebResponse">The HTTP web response.</param>
        /// <param name="webException">The web exception.</param>
        /// <param name="responseHeadersTraceFlag">The trace flag for response headers.</param>
        /// <param name="responseTraceFlag">The trace flag for responses.</param>
        /// <remarks>
        /// This method doesn't handle 500 ISE errors. This is handled by the caller since
        /// 500 ISE typically indicates that a SOAP fault has occurred and the handling of
        /// a SOAP fault is currently service specific.
        /// </remarks>
        internal void InternalProcessHttpErrorResponse(
                            IEwsHttpWebResponse httpWebResponse,
                            WebException webException,
                            TraceFlags responseHeadersTraceFlag,
                            TraceFlags responseTraceFlag)
            {
            EwsUtilities.Assert(
                httpWebResponse.StatusCode != HttpStatusCode.InternalServerError,
                "ExchangeServiceBase.InternalProcessHttpErrorResponse",
                "InternalProcessHttpErrorResponse does not handle 500 ISE errors, the caller is supposed to handle this.");

            ProcessHttpResponseHeaders(responseHeadersTraceFlag, httpWebResponse);

            // Deal with new HTTP error code indicating that account is locked.
            // The "unlock" URL is returned as the status description in the response.
            if (httpWebResponse.StatusCode == ExchangeServiceBase.AccountIsLocked)
                {
                string location = httpWebResponse.StatusDescription;

                Uri accountUnlockUrl = null;
                if (Uri.IsWellFormedUriString(location, UriKind.Absolute))
                    {
                    accountUnlockUrl = new Uri(location);
                    }

                TraceMessage(responseTraceFlag, string.Format("Account is locked. Unlock URL is {0}", accountUnlockUrl));

                throw new AccountIsLockedException(
                    string.Format(Strings.AccountIsLocked, accountUnlockUrl),
                    accountUnlockUrl,
                    webException);
                }
            }

        /// <summary>
        /// Processes an HTTP error response.
        /// </summary>
        /// <param name="httpWebResponse">The HTTP web response.</param>
        /// <param name="webException">The web exception.</param>
        internal abstract void ProcessHttpErrorResponse(IEwsHttpWebResponse httpWebResponse, WebException webException);

        /// <summary>
        /// Determines whether tracing is enabled for specified trace flag(s).
        /// </summary>
        /// <param name="traceFlags">The trace flags.</param>
        /// <returns>True if tracing is enabled for specified trace flag(s).
        /// </returns>
        internal bool IsTraceEnabledFor(TraceFlags traceFlags)
            {
            return TraceEnabled && ((TraceFlags & traceFlags) != 0);
            }

        /// <summary>
        /// Logs the specified string to the TraceListener if tracing is enabled.
        /// </summary>
        /// <param name="traceType">Kind of trace entry.</param>
        /// <param name="logEntry">The entry to log.</param>
        internal void TraceMessage(TraceFlags traceType, string logEntry)
            {
            if (IsTraceEnabledFor(traceType))
                {
                string traceTypeStr = traceType.ToString();
                string logMessage = EwsUtilities.FormatLogMessage(traceTypeStr, logEntry);
                TraceListener.Trace(traceTypeStr, logMessage);
                }
            }

        /// <summary>
        /// Logs the specified XML to the TraceListener if tracing is enabled.
        /// </summary>
        /// <param name="traceType">Kind of trace entry.</param>
        /// <param name="stream">The stream containing XML.</param>
        internal void TraceXml(TraceFlags traceType, MemoryStream stream)
            {
            if (IsTraceEnabledFor(traceType))
                {
                string traceTypeStr = traceType.ToString();
                string logMessage = EwsUtilities.FormatLogMessageWithXmlContent(traceTypeStr, stream);
                TraceListener.Trace(traceTypeStr, logMessage);
                }
            }

        /// <summary>
        /// Traces the HTTP request headers.
        /// </summary>
        /// <param name="traceType">Kind of trace entry.</param>
        /// <param name="request">The request.</param>
        internal void TraceHttpRequestHeaders(TraceFlags traceType, IEwsHttpWebRequest request)
            {
            if (IsTraceEnabledFor(traceType))
                {
                string traceTypeStr = traceType.ToString();
                string headersAsString = EwsUtilities.FormatHttpRequestHeaders(request);
                string logMessage = EwsUtilities.FormatLogMessage(traceTypeStr, headersAsString);
                TraceListener.Trace(traceTypeStr, logMessage);
                }
            }

        /// <summary>
        /// Traces the HTTP response headers.
        /// </summary>
        /// <param name="traceType">Kind of trace entry.</param>
        /// <param name="response">The response.</param>
        internal void ProcessHttpResponseHeaders(TraceFlags traceType, IEwsHttpWebResponse response)
            {
            TraceHttpResponseHeaders(traceType, response);

            SaveHttpResponseHeaders(response.Headers);
            }

        /// <summary>
        /// Traces the HTTP response headers.
        /// </summary>
        /// <param name="traceType">Kind of trace entry.</param>
        /// <param name="response">The response.</param>
        private void TraceHttpResponseHeaders(TraceFlags traceType, IEwsHttpWebResponse response)
            {
            if (IsTraceEnabledFor(traceType))
                {
                string traceTypeStr = traceType.ToString();
                string headersAsString = EwsUtilities.FormatHttpResponseHeaders(response);
                string logMessage = EwsUtilities.FormatLogMessage(traceTypeStr, headersAsString);
                TraceListener.Trace(traceTypeStr, logMessage);
                }
            }

        /// <summary>
        /// Save the HTTP response headers.
        /// </summary>
        /// <param name="headers">The response headers</param>
        private void SaveHttpResponseHeaders(WebHeaderCollection headers)
            {
            httpResponseHeaders.Clear();

            foreach (string key in headers.AllKeys)
                {
                string existingValue;

                if (httpResponseHeaders.TryGetValue(key, out existingValue))
                    {
                    httpResponseHeaders[key] = existingValue + "," + headers[key];
                    }
                else
                    {
                    httpResponseHeaders.Add(key, headers[key]);
                    }
                }

            if (OnResponseHeadersCaptured != null)
                {
                OnResponseHeadersCaptured(headers);
                }
            }

        /// <summary>
        /// Converts the universal date time string to local date time.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>DateTime</returns>
        internal DateTime? ConvertUniversalDateTimeStringToLocalDateTime(string value)
            {
            if (string.IsNullOrEmpty(value))
                {
                return null;
                }
            else
                {
                // Assume an unbiased date/time is in UTC. Convert to UTC otherwise.
                DateTime dateTime = DateTime.Parse(
                    value,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal);

                if (TimeZone == TimeZoneInfo.Utc)
                    {
                    // This returns a DateTime with Kind.Utc
                    return dateTime;
                    }
                else
                    {
                    DateTime localTime = EwsUtilities.ConvertTime(
                        dateTime,
                        TimeZoneInfo.Utc,
                        TimeZone);

                    if (EwsUtilities.IsLocalTimeZone(TimeZone))
                        {
                        // This returns a DateTime with Kind.Local
                        return new DateTime(localTime.Ticks, DateTimeKind.Local);
                        }
                    else
                        {
                        // This returns a DateTime with Kind.Unspecified
                        return localTime;
                        }
                    }
                }
            }

        /// <summary>
        /// Converts xs:dateTime string with either "Z", "-00:00" bias, or "" suffixes to 
        /// unspecified StartDate value ignoring the suffix.
        /// </summary>
        /// <param name="value">The string value to parse.</param>
        /// <returns>The parsed DateTime value.</returns>
        internal DateTime? ConvertStartDateToUnspecifiedDateTime(string value)
            {
            if (string.IsNullOrEmpty(value))
                {
                return null;
                }
            else
                {
                DateTimeOffset dateTimeOffset = DateTimeOffset.Parse(value, CultureInfo.InvariantCulture);

                // Return only the date part with the kind==Unspecified.
                return dateTimeOffset.Date;
                }
            }

        /// <summary>
        /// Converts the date time to universal date time string.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>String representation of DateTime.</returns>
        internal string ConvertDateTimeToUniversalDateTimeString(DateTime value)
            {
            DateTime dateTime;

            switch (value.Kind)
                {
                case DateTimeKind.Unspecified:
                    dateTime = EwsUtilities.ConvertTime(
                        value,
                        TimeZone,
                        TimeZoneInfo.Utc);

                    break;
                case DateTimeKind.Local:
                    dateTime = EwsUtilities.ConvertTime(
                        value,
                        TimeZoneInfo.Local,
                        TimeZoneInfo.Utc);

                    break;
                default:
                    // The date is already in UTC, no need to convert it.
                    dateTime = value;

                    break;
                }
            return dateTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
            }

        /// <summary>
        /// Register the custom auth module to support non-ascii upn authentication if the server supports that 
        /// </summary>
        internal void RegisterCustomBasicAuthModule()
            {
            if (RequestedServerVersion >= ExchangeVersion.Exchange2013_SP1)
                {
                BasicAuthModuleForUTF8.InstantiateIfNeeded();
                }
            }

        /// <summary>
        /// Sets the user agent to a custom value
        /// </summary>
        /// <param name="userAgent">User agent string to set on the service</param>
        internal void SetCustomUserAgent(string userAgent)
            {
            this.userAgent = userAgent;
            }

        #endregion

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeServiceBase"/> class.
        /// </summary>
        internal ExchangeServiceBase()
            : this(TimeZoneInfo.Local)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeServiceBase"/> class.
        /// </summary>
        /// <param name="timeZone">The time zone to which the service is scoped.</param>
        internal ExchangeServiceBase(TimeZoneInfo timeZone)
            {
            this.timeZone = timeZone;
            UseDefaultCredentials = true;
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeServiceBase"/> class.
        /// </summary>
        /// <param name="requestedServerVersion">The requested server version.</param>
        internal ExchangeServiceBase(ExchangeVersion requestedServerVersion)
            : this(requestedServerVersion, TimeZoneInfo.Local)
            {
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeServiceBase"/> class.
        /// </summary>
        /// <param name="requestedServerVersion">The requested server version.</param>
        /// <param name="timeZone">The time zone to which the service is scoped.</param>
        internal ExchangeServiceBase(ExchangeVersion requestedServerVersion, TimeZoneInfo timeZone)
            : this(timeZone)
            {
            this.requestedServerVersion = requestedServerVersion;
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeServiceBase"/> class.
        /// </summary>
        /// <param name="service">The other service.</param>
        /// <param name="requestedServerVersion">The requested server version.</param>
        internal ExchangeServiceBase(ExchangeServiceBase service, ExchangeVersion requestedServerVersion)
            : this(requestedServerVersion)
            {
            useDefaultCredentials = service.useDefaultCredentials;
            credentials = service.credentials;
            traceEnabled = service.traceEnabled;
            traceListener = service.traceListener;
            traceFlags = service.traceFlags;
            timeout = service.timeout;
            preAuthenticate = service.preAuthenticate;
            userAgent = service.userAgent;
            acceptGzipEncoding = service.acceptGzipEncoding;
            keepAlive = service.keepAlive;
            connectionGroupName = service.connectionGroupName;
            timeZone = service.timeZone;
            httpHeaders = service.httpHeaders;
            ewsHttpWebRequestFactory = service.ewsHttpWebRequestFactory;
            webProxy = service.webProxy;
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="ExchangeServiceBase"/> class from existing one.
        /// </summary>
        /// <param name="service">The other service.</param>
        internal ExchangeServiceBase(ExchangeServiceBase service)
            : this(service, service.RequestedServerVersion)
            {
            }

        #endregion

        #region Validation

        /// <summary>
        /// Validates this instance.
        /// </summary>
        internal virtual void Validate()
            {
            }

        #endregion

        #region Properties

        /// <summary>
        /// Gets or sets the cookie container.
        /// </summary>
        /// <value>The cookie container.</value>
        public CookieContainer CookieContainer
            {
            get { return cookieContainer; }
            set { cookieContainer = value; }
            }

        /// <summary>
        /// Gets the time zone this service is scoped to.
        /// </summary>
        internal TimeZoneInfo TimeZone
            {
            get { return timeZone; }
            }

        /// <summary>
        /// Gets a time zone definition generated from the time zone info to which this service is scoped.
        /// </summary>
        internal TimeZoneDefinition TimeZoneDefinition
            {
            get
                {
                if (timeZoneDefinition == null)
                    {
                    timeZoneDefinition = new TimeZoneDefinition(TimeZone);
                    }

                return timeZoneDefinition;
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether client latency info is push to server.
        /// </summary>
        public bool SendClientLatencies
            {
            get
                {
                return sendClientLatencies;
                }

            set
                {
                sendClientLatencies = value;
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether tracing is enabled.
        /// </summary>
        public bool TraceEnabled
            {
            get
                {
                return traceEnabled;
                }

            set
                {
                traceEnabled = value;
                if (traceEnabled && (traceListener == null))
                    {
                    traceListener = new EwsTraceListener();
                    }
                }
            }

        /// <summary>
        /// Gets or sets the trace flags.
        /// </summary>
        /// <value>The trace flags.</value>
        public TraceFlags TraceFlags
            {
            get
                {
                return traceFlags;
                }

            set
                {
                traceFlags = value;
                }
            }

        /// <summary>
        /// Gets or sets the trace listener.
        /// </summary>
        /// <value>The trace listener.</value>
        public ITraceListener TraceListener
            {
            get
                {
                return traceListener;
                }

            set
                {
                traceListener = value;
                traceEnabled = value != null;
                }
            }

        /// <summary>
        /// Gets or sets the credentials used to authenticate with the Exchange Web Services. Setting the Credentials property
        /// automatically sets the UseDefaultCredentials to false.
        /// </summary>
        public ExchangeCredentials Credentials
            {
            get
                {
                return credentials;
                }

            set
                {
                credentials = value;
                useDefaultCredentials = false;
                cookieContainer = new CookieContainer();       // Changing credentials resets the Cookie container
                }
            }

        /// <summary>
        /// Gets or sets a value indicating whether the credentials of the user currently logged into Windows should be used to
        /// authenticate with the Exchange Web Services. Setting UseDefaultCredentials to true automatically sets the Credentials
        /// property to null.
        /// </summary>
        public bool UseDefaultCredentials
            {
            get
                {
                return useDefaultCredentials;
                }

            set
                {
                useDefaultCredentials = value;

                if (value)
                    {
                    credentials = null;
                    cookieContainer = new CookieContainer();   // Changing credentials resets the Cookie container
                    }
                }
            }

        /// <summary>
        /// Gets or sets the timeout used when sending HTTP requests and when receiving HTTP responses, in milliseconds.
        /// Defaults to 100000.
        /// </summary>
        public int Timeout
            {
            get
                {
                return timeout;
                }

            set
                {
                if (value < 1)
                    {
                    throw new ArgumentException(Strings.TimeoutMustBeGreaterThanZero);
                    }

                timeout = value;
                }
            }

        /// <summary>
        /// Gets or sets a value that indicates whether HTTP pre-authentication should be performed.
        /// </summary>
        public bool PreAuthenticate
            {
            get { return preAuthenticate; }
            set { preAuthenticate = value; }
            }

        /// <summary>
        /// Gets or sets a value indicating whether GZip compression encoding should be accepted.
        /// </summary>
        /// <remarks>
        /// This value will tell the server that the client is able to handle GZip compression encoding. The server
        /// will only send Gzip compressed content if it has been configured to do so.
        /// </remarks>
        public bool AcceptGzipEncoding
            {
            get { return acceptGzipEncoding; }
            set { acceptGzipEncoding = value; }
            }

        /// <summary>
        /// Gets the requested server version.
        /// </summary>
        /// <value>The requested server version.</value>
        public ExchangeVersion RequestedServerVersion
            {
            get { return requestedServerVersion; }
            }

        /// <summary>
        /// Gets or sets the user agent.
        /// </summary>
        /// <value>The user agent.</value>
        public string UserAgent
            {
            get { return userAgent; }
            set { userAgent = value + " (" + ExchangeService.defaultUserAgent + ")"; }
            }

        /// <summary>
        /// Gets information associated with the server that processed the last request.
        /// Will be null if no requests have been processed.
        /// </summary>
        public ExchangeServerInfo ServerInfo
            {
            get { return serverInfo; }
            internal set { serverInfo = value; }
            }

        /// <summary>
        /// Gets or sets the web proxy that should be used when sending requests to EWS.
        /// Set this property to null to use the default web proxy.
        /// </summary>
        public IWebProxy WebProxy
            {
            get { return webProxy; }
            set { webProxy = value; }
            }

        /// <summary>
        /// Gets or sets if the request to the internet resource should contain a Connection HTTP header with the value Keep-alive
        /// </summary>
        public bool KeepAlive
            {
            get
                {
                return keepAlive;
                }

            set
                {
                keepAlive = value;
                }
            }

        /// <summary>
        /// Gets or sets the name of the connection group for the request. 
        /// </summary>
        public string ConnectionGroupName
            {
            get
                {
                return connectionGroupName;
                }

            set
                {
                connectionGroupName = value;
                }
            }

        /// <summary>
        /// Gets or sets the request id for the request.
        /// </summary>
        public string ClientRequestId
            {
            get { return clientRequestId; }
            set { clientRequestId = value; }
            }

        /// <summary>
        /// Gets or sets a flag to indicate whether the client requires the server side to return the  request id.
        /// </summary>
        public bool ReturnClientRequestId
            {
            get { return returnClientRequestId; }
            set { returnClientRequestId = value; }
            }

        /// <summary>
        /// Gets a collection of HTTP headers that will be sent with each request to EWS.
        /// </summary>
        public IDictionary<string, string> HttpHeaders
            {
            get { return httpHeaders; }
            }

        /// <summary>
        /// Gets a collection of HTTP headers from the last response.
        /// </summary>
        public IDictionary<string, string> HttpResponseHeaders
            {
            get { return httpResponseHeaders; }
            }

        /// <summary>
        /// Gets the session key.
        /// </summary>
        internal static byte[] SessionKey
            {
            get
                {
                // this has to be computed only once.
                lock (ExchangeServiceBase.lockObj)
                    {
                    if (ExchangeServiceBase.binarySecret == null)
                        {
                        RandomNumberGenerator randomNumberGenerator = RandomNumberGenerator.Create();
                        ExchangeServiceBase.binarySecret = new byte[256 / 8];
                        randomNumberGenerator.GetNonZeroBytes(binarySecret);
                        }

                    return ExchangeServiceBase.binarySecret;
                    }
                }
            }

        /// <summary>
        /// Gets or sets the HTTP web request factory.
        /// </summary>
        internal IEwsHttpWebRequestFactory HttpWebRequestFactory
            {
            get { return ewsHttpWebRequestFactory; }

            set
                {
                // If new value is null, reset to default factory.
                ewsHttpWebRequestFactory = (value == null) ? new EwsHttpWebRequestFactory() : value;
                }
            }

        /// <summary>
        /// For testing: suppresses generation of the SOAP version header.
        /// </summary>
        internal bool SuppressXmlVersionHeader { get; set; }

        #endregion

        #region Events

        /// <summary>
        /// Provides an event that applications can implement to emit custom SOAP headers in requests that are sent to Exchange.
        /// </summary>
        public event CustomXmlSerializationDelegate OnSerializeCustomSoapHeaders;

        #endregion
        }
    }