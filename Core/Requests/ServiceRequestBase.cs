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
    using System.IO;
    using System.IO.Compression;
    using System.Net;
    using System.Text;
    using System.Xml;

    /// <summary>
    /// Represents an abstract service request.
    /// </summary>
    internal abstract class ServiceRequestBase
        {
        #region Private Constants
        /// <summary>
        /// The two contants below are used to set the AnchorMailbox and ExplicitLogonUser values
        /// in the request header.
        /// </summary>
        /// <remarks>
        /// Note: Setting this values will route the request directly to the backend hosting the 
        /// AnchorMailbox. These headers should be used primarily for UnifiedGroup scenario where
        /// a request needs to be routed directly to the group mailbox versus the user mailbox.
        /// </remarks>
        private const string AnchorMailboxHeaderName = "X-AnchorMailbox";
        private const string ExplicitLogonUserHeaderName = "X-OWA-ExplicitLogonUser";

        private static readonly string[] RequestIdResponseHeaders = new[] { "RequestId", "request-id", };
        private const string XMLSchemaNamespace = "http://www.w3.org/2001/XMLSchema";
        private const string XMLSchemaInstanceNamespace = "http://www.w3.org/2001/XMLSchema-instance";
        private const string ClientStatisticsRequestHeader = "X-ClientStatistics";

        #endregion

        /// <summary>
        /// Gets or sets the anchor mailbox associated with the request
        /// </summary>
        /// <remarks>
        /// Setting this value will add special headers to the request which in turn
        /// will route the request directly to the mailbox server against which the request
        /// is to be executed.
        /// </remarks>
        internal string AnchorMailbox
            {
            get;
            set;
            }

        /// <summary>
        /// Maintains the collection of client side statistics for requests already completed
        /// </summary>
        private static List<string> clientStatisticsCache = new();

        private ExchangeService service;

        /// <summary>
        /// Gets the response stream (may be wrapped with GZip/Deflate stream to decompress content)
        /// </summary>
        /// <param name="response">HttpWebResponse.</param>
        /// <returns>ResponseStream</returns>
        protected static Stream GetResponseStream(IEwsHttpWebResponse response)
            {
            string contentEncoding = response.ContentEncoding;
            Stream responseStream = response.GetResponseStream();

            return WrapStream(responseStream, response.ContentEncoding);
            }

        /// <summary>
        /// Gets the response stream (may be wrapped with GZip/Deflate stream to decompress content)
        /// </summary>
        /// <param name="response">HttpWebResponse.</param>
        /// <param name="readTimeout">read timeout in milliseconds</param>
        /// <returns>ResponseStream</returns>
        protected static Stream GetResponseStream(IEwsHttpWebResponse response, int readTimeout)
            {
            Stream responseStream = response.GetResponseStream();

            responseStream.ReadTimeout = readTimeout;
            return WrapStream(responseStream, response.ContentEncoding);
            }

        private static Stream WrapStream(Stream responseStream, string contentEncoding)
            {
            if (contentEncoding.ToLowerInvariant().Contains("gzip"))
                {
                return new GZipStream(responseStream, CompressionMode.Decompress);
                }
            else if (contentEncoding.ToLowerInvariant().Contains("deflate"))
                {
                return new DeflateStream(responseStream, CompressionMode.Decompress);
                }
            else
                {
                return responseStream;
                }
            }

        #region Methods for subclasses to override

        /// <summary>
        /// Gets the name of the XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal abstract string GetXmlElementName();

        /// <summary>
        /// Gets the name of the response XML element.
        /// </summary>
        /// <returns>XML element name,</returns>
        internal abstract string GetResponseXmlElementName();

        /// <summary>
        /// Gets the minimum server version required to process this request.
        /// </summary>
        /// <returns>Exchange server version.</returns>
        internal abstract ExchangeVersion GetMinimumRequiredServerVersion();

        /// <summary>
        /// Writes XML elements.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal abstract void WriteElementsToXml(EwsServiceXmlWriter writer);

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>Response object.</returns>
        internal virtual object ParseResponse(EwsServiceXmlReader reader)
            {
            throw new NotImplementedException("you must override either this or the 2-parameter version");
            }

        /// <summary>
        /// Parses the response.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <param name="responseHeaders">Response headers</param>
        /// <returns>Response object.</returns>
        /// <remarks>If this is overriden instead of the 1-parameter version, you can read response headers</remarks>
        internal virtual object ParseResponse(EwsServiceXmlReader reader, WebHeaderCollection responseHeaders)
            {
            return ParseResponse(reader);
            }

        /// <summary>
        /// Gets a value indicating whether the TimeZoneContext SOAP header should be eimitted.
        /// </summary>
        /// <value><c>true</c> if the time zone should be emitted; otherwise, <c>false</c>.</value>
        internal virtual bool EmitTimeZoneHeader
            {
            get { return false; }
            }

        #endregion

        /// <summary>
        /// Validate request.
        /// </summary>
        internal virtual void Validate()
            {
            Service.Validate();
            }

        /// <summary>
        /// Writes XML body.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteBodyToXml(EwsServiceXmlWriter writer)
            {
            writer.WriteStartElement(XmlNamespace.Messages, GetXmlElementName());

            WriteAttributesToXml(writer);
            WriteElementsToXml(writer);

            writer.WriteEndElement(); // m:this.GetXmlElementName()
            }

        /// <summary>
        /// Writes XML attributes.
        /// </summary>
        /// <remarks>
        /// Subclass will override if it has XML attributes.
        /// </remarks>
        /// <param name="writer">The writer.</param>
        internal virtual void WriteAttributesToXml(EwsServiceXmlWriter writer)
            {
            }

        /// <summary>
        /// Allows the subclasses to add their own header information
        /// </summary>
        /// <param name="webHeaderCollection">The HTTP request headers</param>
        internal virtual void AddHeaders(WebHeaderCollection webHeaderCollection)
            {
            if (!string.IsNullOrEmpty(AnchorMailbox))
                {
                webHeaderCollection.Set(AnchorMailboxHeaderName, AnchorMailbox);
                webHeaderCollection.Set(ExplicitLogonUserHeaderName, AnchorMailbox);
                }
            }

        /// <summary>
        /// Initializes a new instance of the <see cref="ServiceRequestBase"/> class.
        /// </summary>
        /// <param name="service">The service.</param>
        internal ServiceRequestBase(ExchangeService service)
            {
            if (service == null)
                {
                throw new ArgumentNullException("service");
                }

            this.service = service;
            ThrowIfNotSupportedByRequestedServerVersion();
            }

        /// <summary>
        /// Gets the service.
        /// </summary>
        /// <value>The service.</value>
        internal ExchangeService Service
            {
            get { return service; }
            }

        /// <summary>
        /// Throw exception if request is not supported in requested server version.
        /// </summary>
        /// <exception cref="ServiceVersionException">Raised if request requires a later version of Exchange.</exception>
        internal void ThrowIfNotSupportedByRequestedServerVersion()
            {
            if (Service.RequestedServerVersion < GetMinimumRequiredServerVersion())
                {
                throw new ServiceVersionException(
                    string.Format(
                        Strings.RequestIncompatibleWithRequestVersion,
                        GetXmlElementName(),
                        GetMinimumRequiredServerVersion()));
                }
            }

        #region HttpWebRequest-based implementation

        /// <summary>
        /// Writes XML.
        /// </summary>
        /// <param name="writer">The writer.</param>
        internal void WriteToXml(EwsServiceXmlWriter writer)
            {
            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
            writer.WriteAttributeValue("xmlns", EwsUtilities.EwsXmlSchemaInstanceNamespacePrefix, EwsUtilities.EwsXmlSchemaInstanceNamespace);
            writer.WriteAttributeValue("xmlns", EwsUtilities.EwsMessagesNamespacePrefix, EwsUtilities.EwsMessagesNamespace);
            writer.WriteAttributeValue("xmlns", EwsUtilities.EwsTypesNamespacePrefix, EwsUtilities.EwsTypesNamespace);
            if (writer.RequireWSSecurityUtilityNamespace)
                {
                writer.WriteAttributeValue("xmlns", EwsUtilities.WSSecurityUtilityNamespacePrefix, EwsUtilities.WSSecurityUtilityNamespace);
                }

            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName);

            if (Service.Credentials != null)
                {
                Service.Credentials.EmitExtraSoapHeaderNamespaceAliases(writer.InternalWriter);
                }

            // Emit the RequestServerVersion header
            if (!Service.SuppressXmlVersionHeader)
                {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.RequestServerVersion);
                writer.WriteAttributeValue(XmlAttributeNames.Version, GetRequestedServiceVersionString());
                writer.WriteEndElement(); // RequestServerVersion
                }

            // Against Exchange 2007 SP1, we always emit the simplified time zone header. It adds very little to
            // the request, so bandwidth consumption is not an issue. Against Exchange 2010 and above, we emit
            // the full time zone header but only when the request actually needs it.
            //
            // The exception to this is if we are in Exchange2007 Compat Mode, in which case we should never emit 
            // the header.  (Note: Exchange2007 Compat Mode is enabled for testability purposes only.)
            //
            if ((Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1 || EmitTimeZoneHeader) &&
                (!Service.Exchange2007CompatibilityMode))
                {
                writer.WriteStartElement(XmlNamespace.Types, XmlElementNames.TimeZoneContext);

                Service.TimeZoneDefinition.WriteToXml(writer);

                writer.WriteEndElement(); // TimeZoneContext

                writer.IsTimeZoneHeaderEmitted = true;
                }

            // Emit the MailboxCulture header
            if (Service.PreferredCulture != null)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.MailboxCulture,
                    Service.PreferredCulture.Name);
                }

            // Emit the DateTimePrecision header
            if (Service.DateTimePrecision != DateTimePrecision.Default)
                {
                writer.WriteElementValue(
                    XmlNamespace.Types,
                    XmlElementNames.DateTimePrecision,
                    Service.DateTimePrecision.ToString());
                }

            // Emit the ExchangeImpersonation header
            if (Service.ImpersonatedUserId != null)
                {
                Service.ImpersonatedUserId.WriteToXml(writer);
                }
            else if (Service.PrivilegedUserId != null)
                {
                Service.PrivilegedUserId.WriteToXml(writer, Service.RequestedServerVersion);
                }
            else if (Service.ManagementRoles != null)
                {
                Service.ManagementRoles.WriteToXml(writer);
                }

            if (Service.Credentials != null)
                {
                Service.Credentials.SerializeExtraSoapHeaders(writer.InternalWriter, GetXmlElementName());
                }

            Service.DoOnSerializeCustomSoapHeaders(writer.InternalWriter);

            writer.WriteEndElement(); // soap:Header

            writer.WriteStartElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);

            WriteBodyToXml(writer);

            writer.WriteEndElement(); // soap:Body
            writer.WriteEndElement(); // soap:Envelope
            }

        /// <summary>
        /// Gets string representation of requested server version.
        /// </summary>
        /// <remarks>
        /// In order to support E12 RTM servers, ExchangeService has another flag indicating that
        /// we should use "Exchange2007" as the server version string rather than Exchange2007_SP1.
        /// </remarks>
        /// <returns>String representation of requested server version.</returns>
        private string GetRequestedServiceVersionString()
            {
            if (Service.Exchange2007CompatibilityMode && Service.RequestedServerVersion == ExchangeVersion.Exchange2007_SP1)
                {
                return "Exchange2007";
                }
            else
                {
                return Service.RequestedServerVersion.ToString();
                }
            }

        /// <summary>
        /// Emits the request.
        /// </summary>
        /// <param name="request">The request.</param>
        private void EmitRequest(IEwsHttpWebRequest request)
            {
            using (Stream requestStream = GetWebRequestStream(request))
                {
                using (EwsServiceXmlWriter writer = new(Service, requestStream))
                    {
                    WriteToXml(writer);
                    }
                }
            }

        /// <summary>
        /// Traces the and emits the request.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <param name="needSignature"></param>
        /// <param name="needTrace"></param>
        private void TraceAndEmitRequest(IEwsHttpWebRequest request, bool needSignature, bool needTrace)
            {
            using (MemoryStream memoryStream = new())
                {
                using (EwsServiceXmlWriter writer = new(Service, memoryStream))
                    {
                    writer.RequireWSSecurityUtilityNamespace = needSignature;
                    WriteToXml(writer);
                    }

                if (needSignature)
                    {
                    service.Credentials.Sign(memoryStream);
                    }

                if (needTrace)
                    {
                    TraceXmlRequest(memoryStream);
                    }

                using (Stream serviceRequestStream = GetWebRequestStream(request))
                    {
                    EwsUtilities.CopyStream(memoryStream, serviceRequestStream);
                    }
                }
            }

        /// <summary>
        /// Get the request stream
        /// </summary>
        /// <param name="request">The request</param>
        /// <returns>The Request stream</returns>
        private Stream GetWebRequestStream(IEwsHttpWebRequest request)
            {
            // In the async case, although we can use async callback to make the entire worflow completely async, 
            // there is little perf gain with this approach because of EWS's message nature.
            // The overall latency of BeginGetRequestStream() is same as GetRequestStream() in this case.
            // The overhead to implement a two-step async operation includes wait handle synchronization, exception handling and wrapping.
            // Therefore, we only leverage BeginGetResponse() and EndGetResponse() to provide the async functionality.
            // Reference: http://www.wintellect.com/CS/blogs/jeffreyr/archive/2009/02/08/httpwebrequest-its-request-stream-and-sending-data-in-chunks.aspx
            return request.EndGetRequestStream(request.BeginGetRequestStream(null, null));
            }

        /// <summary>
        /// Reads the response.
        /// </summary>
        /// <param name="ewsXmlReader">The XML reader.</param>
        /// <param name="responseHeaders">HTTP response headers</param>
        /// <returns>Service response.</returns>
        protected object ReadResponse(EwsServiceXmlReader ewsXmlReader, WebHeaderCollection responseHeaders)
            {
            object serviceResponse;

            ReadPreamble(ewsXmlReader);
            ewsXmlReader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
            ReadSoapHeader(ewsXmlReader);
            ewsXmlReader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);

            ewsXmlReader.ReadStartElement(XmlNamespace.Messages, GetResponseXmlElementName());

            if (responseHeaders != null)
                {
                serviceResponse = ParseResponse(ewsXmlReader, responseHeaders);
                }
            else
                {
                serviceResponse = ParseResponse(ewsXmlReader);
                }

            ewsXmlReader.ReadEndElementIfNecessary(XmlNamespace.Messages, GetResponseXmlElementName());

            ewsXmlReader.ReadEndElement(XmlNamespace.Soap, XmlElementNames.SOAPBodyElementName);
            ewsXmlReader.ReadEndElement(XmlNamespace.Soap, XmlElementNames.SOAPEnvelopeElementName);
            return serviceResponse;
            }

        /// <summary>
        /// Reads any preamble data not part of the core response.
        /// </summary>
        /// <param name="ewsXmlReader">The EwsServiceXmlReader.</param>
        protected virtual void ReadPreamble(EwsServiceXmlReader ewsXmlReader)
            {
            ReadXmlDeclaration(ewsXmlReader);
            }

        /// <summary>
        /// Read SOAP header and extract server version
        /// </summary>
        /// <param name="reader">EwsServiceXmlReader</param>
        private void ReadSoapHeader(EwsServiceXmlReader reader)
            {
            reader.ReadStartElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName);
            do
                {
                reader.Read();

                // Is this the ServerVersionInfo?
                if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ServerVersionInfo))
                    {
                    Service.ServerInfo = ExchangeServerInfo.Parse(reader);
                    }

                // Ignore anything else inside the SOAP header
                }
            while (!reader.IsEndElement(XmlNamespace.Soap, XmlElementNames.SOAPHeaderElementName));
            }

        /// <summary>
        /// Reads the SOAP fault.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns>SOAP fault details.</returns>
        protected SoapFaultDetails ReadSoapFault(EwsServiceXmlReader reader)
            {
            SoapFaultDetails soapFaultDetails = null;

            try
                {
                ReadXmlDeclaration(reader);

                reader.Read();
                if (!reader.IsStartElement() || (reader.LocalName != XmlElementNames.SOAPEnvelopeElementName))
                    {
                    return soapFaultDetails;
                    }

                // EWS can sometimes return SOAP faults using the SOAP 1.2 namespace. Get the
                // namespace URI from the envelope element and use it for the rest of the parsing.
                // If it's not 1.1 or 1.2, we can't continue.
                XmlNamespace soapNamespace = EwsUtilities.GetNamespaceFromUri(reader.NamespaceUri);
                if (soapNamespace == XmlNamespace.NotSpecified)
                    {
                    return soapFaultDetails;
                    }

                reader.Read();

                // EWS doesn't always return a SOAP header. If this response contains a header element, 
                // read the server version information contained in the header.
                if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPHeaderElementName))
                    {
                    do
                        {
                        reader.Read();

                        if (reader.IsStartElement(XmlNamespace.Types, XmlElementNames.ServerVersionInfo))
                            {
                            Service.ServerInfo = ExchangeServerInfo.Parse(reader);
                            }
                        }
                    while (!reader.IsEndElement(soapNamespace, XmlElementNames.SOAPHeaderElementName));

                    // Queue up the next read
                    reader.Read();
                    }

                // Parse the fault element contained within the SOAP body.
                if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPBodyElementName))
                    {
                    do
                        {
                        reader.Read();

                        // Parse Fault element
                        if (reader.IsStartElement(soapNamespace, XmlElementNames.SOAPFaultElementName))
                            {
                            soapFaultDetails = SoapFaultDetails.Parse(reader, soapNamespace);
                            }
                        }
                    while (!reader.IsEndElement(soapNamespace, XmlElementNames.SOAPBodyElementName));
                    }

                reader.ReadEndElement(soapNamespace, XmlElementNames.SOAPEnvelopeElementName);
                }
            catch (XmlException)
                {
                // If response doesn't contain a valid SOAP fault, just ignore exception and
                // return null for SOAP fault details.
                }

            return soapFaultDetails;
            }

        /// <summary>
        /// Validates request parameters, and emits the request to the server.
        /// </summary>
        /// <param name="request">The request.</param>
        /// <returns>The response returned by the server.</returns>
        protected IEwsHttpWebResponse ValidateAndEmitRequest(out IEwsHttpWebRequest request)
            {
            Validate();

            request = BuildEwsHttpWebRequest();

            if (service.SendClientLatencies)
                {
                string clientStatisticsToAdd = null;

                lock (clientStatisticsCache)
                    {
                    if (clientStatisticsCache.Count > 0)
                        {
                        clientStatisticsToAdd = clientStatisticsCache[0];
                        clientStatisticsCache.RemoveAt(0);
                        }
                    }

                if (!string.IsNullOrEmpty(clientStatisticsToAdd))
                    {
                    if (request.Headers[ClientStatisticsRequestHeader] != null)
                        {
                        request.Headers[ClientStatisticsRequestHeader] =
                            request.Headers[ClientStatisticsRequestHeader]
                            + clientStatisticsToAdd;
                        }
                    else
                        {
                        request.Headers.Add(
                            ClientStatisticsRequestHeader,
                            clientStatisticsToAdd);
                        }
                    }
                }

            DateTime startTime = DateTime.UtcNow;
            IEwsHttpWebResponse response = null;

            try
                {
                response = GetEwsHttpWebResponse(request);
                }
            finally
                {
                if (service.SendClientLatencies)
                    {
                    int clientSideLatency = (int)(DateTime.UtcNow - startTime).TotalMilliseconds;
                    string requestId = string.Empty;
                    string soapAction = GetType().Name.Replace("Request", string.Empty);

                    if (response != null && response.Headers != null)
                        {
                        foreach (string requestIdHeader in ServiceRequestBase.RequestIdResponseHeaders)
                            {
                            string requestIdValue = response.Headers.Get(requestIdHeader);
                            if (!string.IsNullOrEmpty(requestIdValue))
                                {
                                requestId = requestIdValue;
                                break;
                                }
                            }
                        }

                    StringBuilder sb = new();
                    sb.Append("MessageId=");
                    sb.Append(requestId);
                    sb.Append(",ResponseTime=");
                    sb.Append(clientSideLatency);
                    sb.Append(",SoapAction=");
                    sb.Append(soapAction);
                    sb.Append(";");

                    lock (clientStatisticsCache)
                        {
                        clientStatisticsCache.Add(sb.ToString());
                        }
                    }
                }

            return response;
            }

        /// <summary>
        /// Builds the IEwsHttpWebRequest object for current service request with exception handling.
        /// </summary>
        /// <returns>An IEwsHttpWebRequest instance</returns>
        protected IEwsHttpWebRequest BuildEwsHttpWebRequest()
            {
            IEwsHttpWebRequest request = null;
            try
                {
                request = Service.PrepareHttpWebRequest(GetXmlElementName());

                Service.TraceHttpRequestHeaders(TraceFlags.EwsRequestHttpHeaders, request);

                bool needSignature = Service.Credentials != null && Service.Credentials.NeedSignature;
                bool needTrace = Service.IsTraceEnabledFor(TraceFlags.EwsRequest);

                // The request might need to add additional headers
                AddHeaders(request.Headers);

                // If tracing is enabled, we generate the request in-memory so that we
                // can pass it along to the ITraceListener. Then we copy the stream to
                // the request stream.
                if (needSignature || needTrace)
                    {
                    TraceAndEmitRequest(request, needSignature, needTrace);
                    }
                else
                    {
                    EmitRequest(request);
                    }

                return request;
                }
            catch (WebException ex)
                {
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                    {
                    ProcessWebException(ex);
                    }

                // Wrap exception if the above code block didn't throw
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
                }
            catch (IOException e)
                {
                if (request != null)
                    {
                    request.Abort();
                    }
                // Wrap exception.
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
                }
            }

        /// <summary>
        ///  Gets the IEwsHttpWebRequest object from the specified IEwsHttpWebRequest object with exception handling
        /// </summary>
        /// <param name="request">The specified IEwsHttpWebRequest</param>
        /// <returns>An IEwsHttpWebResponse instance</returns>
        protected IEwsHttpWebResponse GetEwsHttpWebResponse(IEwsHttpWebRequest request)
            {
            try
                {
                return request.GetResponse();
                }
            catch (WebException ex)
                {
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                    {
                    ProcessWebException(ex);
                    }

                // Wrap exception if the above code block didn't throw
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
                }
            catch (IOException e)
                {
                // Wrap exception.
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
                }
            }

        /// <summary>
        /// Ends getting the specified async IEwsHttpWebRequest object from the specified IEwsHttpWebRequest object with exception handling.
        /// </summary>
        /// <param name="request">The specified IEwsHttpWebRequest</param>
        /// <param name="asyncResult">An IAsyncResult that references the asynchronous request.</param>
        /// <returns>An IEwsHttpWebResponse instance</returns>
        protected IEwsHttpWebResponse EndGetEwsHttpWebResponse(IEwsHttpWebRequest request, IAsyncResult asyncResult)
            {
            try
                {
                // Note that this call may throw ArgumentException if the HttpWebRequest instance is not the original one,
                // and we just let it out
                return request.EndGetResponse(asyncResult);
                }
            catch (WebException ex)
                {
                if (ex.Status == WebExceptionStatus.ProtocolError && ex.Response != null)
                    {
                    ProcessWebException(ex);
                    }

                // Wrap exception if the above code block didn't throw
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, ex.Message), ex);
                }
            catch (IOException e)
                {
                // Wrap exception.
                throw new ServiceRequestException(string.Format(Strings.ServiceRequestFailed, e.Message), e);
                }
            }

        /// <summary>
        /// Processes the web exception.
        /// </summary>
        /// <param name="webException">The web exception.</param>
        private void ProcessWebException(WebException webException)
            {
            if (webException.Response != null)
                {
                IEwsHttpWebResponse httpWebResponse = Service.HttpWebRequestFactory.CreateExceptionResponse(webException);
                SoapFaultDetails soapFaultDetails = null;

                if (httpWebResponse.StatusCode == HttpStatusCode.InternalServerError)
                    {
                    Service.ProcessHttpResponseHeaders(TraceFlags.EwsResponseHttpHeaders, httpWebResponse);

                    // If tracing is enabled, we read the entire response into a MemoryStream so that we
                    // can pass it along to the ITraceListener. Then we parse the response from the 
                    // MemoryStream.
                    if (Service.IsTraceEnabledFor(TraceFlags.EwsResponse))
                        {
                        using (MemoryStream memoryStream = new())
                            {
                            using (Stream serviceResponseStream = ServiceRequestBase.GetResponseStream(httpWebResponse))
                                {
                                // Copy response to in-memory stream and reset position to start.
                                EwsUtilities.CopyStream(serviceResponseStream, memoryStream);
                                memoryStream.Position = 0;
                                }

                            TraceResponseXml(httpWebResponse, memoryStream);

                            EwsServiceXmlReader reader = new(memoryStream, Service);
                            soapFaultDetails = ReadSoapFault(reader);
                            }
                        }
                    else
                        {
                        using (Stream stream = ServiceRequestBase.GetResponseStream(httpWebResponse))
                            {
                            EwsServiceXmlReader reader = new(stream, Service);
                            soapFaultDetails = ReadSoapFault(reader);
                            }
                        }

                    if (soapFaultDetails != null)
                        {
                        switch (soapFaultDetails.ResponseCode)
                            {
                            case ServiceError.ErrorInvalidServerVersion:
                                throw new ServiceVersionException(Strings.ServerVersionNotSupported);

                            case ServiceError.ErrorSchemaValidation:
                                // If we're talking to an E12 server (8.00.xxxx.xxx), a schema validation error is the same as a version mismatch error.
                                // (Which only will happen if we send a request that's not valid for E12).
                                if ((Service.ServerInfo != null) &&
                                    (Service.ServerInfo.MajorVersion == 8) && (Service.ServerInfo.MinorVersion == 0))
                                    {
                                    throw new ServiceVersionException(Strings.ServerVersionNotSupported);
                                    }

                                break;

                            case ServiceError.ErrorIncorrectSchemaVersion:
                                // This shouldn't happen. It indicates that a request wasn't valid for the version that was specified.
                                EwsUtilities.Assert(
                                    false,
                                    "ServiceRequestBase.ProcessWebException",
                                    "Exchange server supports requested version but request was invalid for that version");
                                break;

                            case ServiceError.ErrorServerBusy:
                                throw new ServerBusyException(new ServiceResponse(soapFaultDetails));

                            default:
                                // Other error codes will be reported as remote error
                                break;
                            }

                        // General fall-through case: throw a ServiceResponseException
                        throw new ServiceResponseException(new ServiceResponse(soapFaultDetails));
                        }
                    }
                else
                    {
                    Service.ProcessHttpErrorResponse(httpWebResponse, webException);
                    }
                }
            }

        /// <summary>
        /// Traces an XML request.  This should only be used for synchronous requests, or synchronous situations
        /// (such as a WebException on an asynchrounous request).
        /// </summary>
        /// <param name="memoryStream">The request content in a MemoryStream.</param>
        protected void TraceXmlRequest(MemoryStream memoryStream)
            {
            Service.TraceXml(TraceFlags.EwsRequest, memoryStream);
            }

        /// <summary>
        /// Traces the response.  This should only be used for synchronous requests, or synchronous situations
        /// (such as a WebException on an asynchrounous request).
        /// </summary>
        /// <param name="response">The response.</param>
        /// <param name="memoryStream">The response content in a MemoryStream.</param>
        protected void TraceResponseXml(IEwsHttpWebResponse response, MemoryStream memoryStream)
            {
            if (!string.IsNullOrEmpty(response.ContentType) &&
                (response.ContentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) ||
                 response.ContentType.StartsWith("application/soap", StringComparison.OrdinalIgnoreCase)))
                {
                Service.TraceXml(TraceFlags.EwsResponse, memoryStream);
                }
            else
                {
                Service.TraceMessage(TraceFlags.EwsResponse, "Non-textual response");
                }
            }

        /// <summary>
        /// Try to read the XML declaration. If it's not there, the server didn't return XML.
        /// </summary>
        /// <param name="reader">The reader.</param>
        private void ReadXmlDeclaration(EwsServiceXmlReader reader)
            {
            try
                {
                reader.Read(XmlNodeType.XmlDeclaration);
                }
            catch (XmlException ex)
                {
                throw new ServiceRequestException(Strings.ServiceResponseDoesNotContainXml, ex);
                }
            catch (ServiceXmlDeserializationException ex)
                {
                throw new ServiceRequestException(Strings.ServiceResponseDoesNotContainXml, ex);
                }
            }

        #endregion
        }
    }