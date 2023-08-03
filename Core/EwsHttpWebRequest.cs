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
    using System.IO;
    using System.Net;
    using System.Security.Cryptography.X509Certificates;

    /// <summary>
    /// Represents an implementation of the IEwsHttpWebRequest interface that uses HttpWebRequest.
    /// </summary>
    internal class EwsHttpWebRequest : IEwsHttpWebRequest
        {
        /// <summary>
        /// Underlying HttpWebRequest.
        /// </summary>
        private HttpWebRequest request;

        /// <summary>
        /// Initializes a new instance of the <see cref="EwsHttpWebRequest"/> class.
        /// </summary>
        /// <param name="uri">The URI.</param>
        internal EwsHttpWebRequest(Uri uri)
            {
            request = (HttpWebRequest)WebRequest.Create(uri);
            }

        #region IEwsHttpWebRequest Members

        /// <summary>
        /// Aborts this instance.
        /// </summary>
        void IEwsHttpWebRequest.Abort()
            {
            request.Abort();
            }

        /// <summary>
        /// Begins an asynchronous request for a <see cref="T:System.IO.Stream"/> object to use to write data.
        /// </summary>
        /// <param name="callback">The <see cref="T:System.AsyncCallback"/> delegate.</param>
        /// <param name="state">The state object for this request.</param>
        /// <returns>
        /// An <see cref="T:System.IAsyncResult"/> that references the asynchronous request.
        /// </returns>
        IAsyncResult IEwsHttpWebRequest.BeginGetRequestStream(AsyncCallback callback, object state)
            {
            return request.BeginGetRequestStream(callback, state);
            }

        /// <summary>
        /// Begins an asynchronous request to an Internet resource.
        /// </summary>
        /// <param name="callback">The <see cref="T:System.AsyncCallback"/> delegate</param>
        /// <param name="state">The state object for this request.</param>
        /// <returns>
        /// An <see cref="T:System.IAsyncResult"/> that references the asynchronous request for a response.
        /// </returns>
        IAsyncResult IEwsHttpWebRequest.BeginGetResponse(AsyncCallback callback, object state)
            {
            return request.BeginGetResponse(callback, state);
            }

        /// <summary>
        /// Ends an asynchronous request for a <see cref="T:System.IO.Stream"/> object to use to write data.
        /// </summary>
        /// <param name="asyncResult">The pending request for a stream.</param>
        /// <returns>
        /// A <see cref="T:System.IO.Stream"/> to use to write request data.
        /// </returns>
        Stream IEwsHttpWebRequest.EndGetRequestStream(IAsyncResult asyncResult)
            {
            return request.EndGetRequestStream(asyncResult);
            }

        /// <summary>
        /// Ends an asynchronous request to an Internet resource.
        /// </summary>
        /// <param name="asyncResult">The pending request for a response.</param>
        /// <returns>
        /// A <see cref="IEwsHttpWebResponse"/> that contains the response from the Internet resource.
        /// </returns>
        IEwsHttpWebResponse IEwsHttpWebRequest.EndGetResponse(IAsyncResult asyncResult)
            {
            return new EwsHttpWebResponse((HttpWebResponse)request.EndGetResponse(asyncResult));
            }

        /// <summary>
        /// Gets a <see cref="T:System.IO.Stream"/> object to use to write request data.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.IO.Stream"/> to use to write request data.
        /// </returns>
        Stream IEwsHttpWebRequest.GetRequestStream()
            {
            return request.GetRequestStream();
            }

        /// <summary>
        /// Returns a response from an Internet resource.
        /// </summary>
        /// <returns>
        /// A <see cref="T:System.Net.HttpWebResponse"/> that contains the response from the Internet resource.
        /// </returns>
        IEwsHttpWebResponse IEwsHttpWebRequest.GetResponse()
            {
            return new EwsHttpWebResponse(request.GetResponse() as HttpWebResponse);
            }

        /// <summary>
        /// Gets or sets the value of the Accept HTTP header.
        /// </summary>
        /// <returns>The value of the Accept HTTP header. The default value is null.</returns>
        string IEwsHttpWebRequest.Accept
            {
            get { return request.Accept; }
            set { request.Accept = value; }
            }

        /// <summary>
        /// Gets or sets a value that indicates whether the request should follow redirection responses.
        /// </summary>
        /// <returns>
        /// True if the request should automatically follow redirection responses from the Internet resource; otherwise, false.
        /// The default value is true.
        /// </returns>
        bool IEwsHttpWebRequest.AllowAutoRedirect
            {
            get { return request.AllowAutoRedirect; }
            set { request.AllowAutoRedirect = value; }
            }

        /// <summary>
        /// Gets or sets the client certificates.
        /// </summary>
        /// <value></value>
        /// <returns>The collection of X509 client certificates.</returns>
        X509CertificateCollection IEwsHttpWebRequest.ClientCertificates
            {
            get { return request.ClientCertificates; }
            set { request.ClientCertificates = value; }
            }

        /// <summary>
        /// Gets or sets the value of the Content-type HTTP header.
        /// </summary>
        /// <returns>The value of the Content-type HTTP header. The default value is null.</returns>
        string IEwsHttpWebRequest.ContentType
            {
            get { return request.ContentType; }
            set { request.ContentType = value; }
            }

        /// <summary>
        /// Gets or sets the cookie container.
        /// </summary>
        /// <value>The cookie container.</value>
        CookieContainer IEwsHttpWebRequest.CookieContainer
            {
            get { return request.CookieContainer; }
            set { request.CookieContainer = value; }
            }

        /// <summary>
        /// Gets or sets authentication information for the request.
        /// </summary>
        /// <returns>An <see cref="T:System.Net.ICredentials"/> that contains the authentication credentials associated with the request. The default is null.</returns>
        ICredentials IEwsHttpWebRequest.Credentials
            {
            get { return request.Credentials; }
            set { request.Credentials = value; }
            }

        /// <summary>
        /// Specifies a collection of the name/value pairs that make up the HTTP headers.
        /// </summary>
        /// <returns>A <see cref="T:System.Net.WebHeaderCollection"/> that contains the name/value pairs that make up the headers for the HTTP request.</returns>
        WebHeaderCollection IEwsHttpWebRequest.Headers
            {
            get { return request.Headers; }
            set { request.Headers = value; }
            }

        /// <summary>
        /// Gets or sets the method for the request.
        /// </summary>
        /// <returns>The request method to use to contact the Internet resource. The default value is GET.</returns>
        /// <exception cref="T:System.ArgumentException">No method is supplied.-or- The method string contains invalid characters. </exception>
        string IEwsHttpWebRequest.Method
            {
            get { return request.Method; }
            set { request.Method = value; }
            }

        /// <summary>
        /// Gets or sets proxy information for the request.
        /// </summary>
        IWebProxy IEwsHttpWebRequest.Proxy
            {
            get { return request.Proxy; }
            set { request.Proxy = value; }
            }

        /// <summary>
        /// Gets or sets a value that indicates whether to send an authenticate header with the request.
        /// </summary>
        /// <returns>true to send a WWW-authenticate HTTP header with requests after authentication has taken place; otherwise, false. The default is false.</returns>
        bool IEwsHttpWebRequest.PreAuthenticate
            {
            get { return request.PreAuthenticate; }
            set { request.PreAuthenticate = value; }
            }

        /// <summary>
        /// Gets the original Uniform Resource Identifier (URI) of the request.
        /// </summary>
        /// <returns>A <see cref="T:System.Uri"/> that contains the URI of the Internet resource passed to the <see cref="M:System.Net.WebRequest.Create(System.String)"/> method.</returns>
        Uri IEwsHttpWebRequest.RequestUri
            {
            get { return request.RequestUri; }
            }

        /// <summary>
        /// Gets or sets the time-out value in milliseconds for the <see cref="M:System.Net.HttpWebRequest.GetResponse"/> and <see cref="M:System.Net.HttpWebRequest.GetRequestStream"/> methods.
        /// </summary>
        /// <returns>The number of milliseconds to wait before the request times out. The default is 100,000 milliseconds (100 seconds).</returns>
        int IEwsHttpWebRequest.Timeout
            {
            get { return request.Timeout; }
            set { request.Timeout = value; }
            }

        /// <summary>
        /// Gets or sets a <see cref="T:System.Boolean"/> value that controls whether default credentials are sent with requests.
        /// </summary>
        /// <returns>true if the default credentials are used; otherwise false. The default value is false.</returns>
        bool IEwsHttpWebRequest.UseDefaultCredentials
            {
            get { return request.UseDefaultCredentials; }
            set { request.UseDefaultCredentials = value; }
            }

        /// <summary>
        /// Gets or sets the value of the User-agent HTTP header.
        /// </summary>
        /// <returns>The value of the User-agent HTTP header. The default value is null.The value for this property is stored in <see cref="T:System.Net.WebHeaderCollection"/>. If WebHeaderCollection is set, the property value is lost.</returns>
        string IEwsHttpWebRequest.UserAgent
            {
            get { return request.UserAgent; }
            set { request.UserAgent = value; }
            }

        /// <summary>
        /// Gets or sets if the request to the internet resource should contain a Connection HTTP header with the value Keep-alive
        /// </summary>
        public bool KeepAlive
            {
            get { return request.KeepAlive; }
            set { request.KeepAlive = value; }
            }

        /// <summary>
        /// Gets or sets the name of the connection group for the request. 
        /// </summary>
        public string ConnectionGroupName
            {
            get { return request.ConnectionGroupName; }
            set { request.ConnectionGroupName = value; }
            }

        #endregion
        }
    }