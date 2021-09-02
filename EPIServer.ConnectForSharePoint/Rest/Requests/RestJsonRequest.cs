using System.Text;

namespace EPiServer.ConnectForSharePoint.Rest.Requests
{
    /// <summary>
    /// Common class for SharePoint REST requests with content type "application/json;odata=verbose".
    /// </summary>
    public class RestJsonRequest : IRestRequest
    {
        private string _restApiPath;
        private string[] _queryParts;
        private string _queryClause;
        private string _method; // defines GET or POST

        /// <summary>
        /// Initial a RestJsonRequest with GET method
        /// </summary>
        /// <param name="apiPath">The SP REST endpoint</param>
        public RestJsonRequest(string apiPath)
            : this(apiPath, "GET", null)
        { }

        /// <summary>
        /// Initial a RestJsonRequest 
        /// </summary>
        /// <param name="restApiPath">The SP REST endpoint, e.g /_api/web/lists.</param>
        /// <param name="requestMethod">The methhod for WebRequest</param>
        /// <param name="queryParts">Query parts must include the key word, e.g. select, expand.</param>
        public RestJsonRequest(string restApiPath, string requestMethod, params string[] queryParts)
        {
            _restApiPath = restApiPath;
            _queryParts = queryParts;
            _method = requestMethod;
        }

        /// <summary>
        /// Return query part of the request, e.g. 
        /// &$select=id,title,LastItemModifiedDate&$expand=Views/Title&$filter=BaseType eq 1 and Hidden eq false.
        /// </summary>
        public virtual string QueryClause
        {
            get
            {
                if (string.IsNullOrEmpty(_queryClause))
                {
                    if (_queryParts == null || (_queryParts != null && _queryParts.Length < 1))
                    {
                        return string.Empty;
                    }

                    _queryClause = string.Join("&$", _queryParts);
                }

                return _queryClause;
            }
        }

        /// <summary>
        /// Gets the SP REST path.
        /// </summary>
        public virtual string RestApiPath
        {
            get
            {
                return _restApiPath;
            }
        }

        public override string ToString()
        {
            var reqString = new StringBuilder(RestApiPath);

            if (!string.IsNullOrEmpty(QueryClause))
            {
                reqString.AppendFormat("?${0}", QueryClause);
            }

            return reqString.ToString();
        }

        /// <summary>
        /// Return the encoded string.
        /// </summary>
        public virtual string EncodedUrl
        {
            get
            {
                var requestUrl = this.ToString();

                // TECHNOTE: We just wanted to replace a space character only.
                return requestUrl.Replace(" ", "%20");
            }
        }

        /// <summary>
        /// Return request methhod
        /// </summary>
        public virtual string RequestMethod
        {
            get
            {
                return _method;
            }
        }
    }
}