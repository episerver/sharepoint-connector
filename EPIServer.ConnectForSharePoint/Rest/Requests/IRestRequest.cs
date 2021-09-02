using System;

namespace EPiServer.ConnectForSharePoint.Rest.Requests
{
    /// <summary>
    /// Interface for request to SharePoint REST API request.
    /// </summary>
    public interface IRestRequest
    {
        /// <summary>
        /// The query part for the request, e.g. 
        /// &$select=id,title,LastItemModifiedDate&$expand=Views/Title&$filter=BaseType eq 1 and Hidden eq false.
        /// </summary>
        string QueryClause { get; }

        /// <summary>
        /// The SP REST request path, e.g. /_api/web/lists.
        /// </summary>
        string RestApiPath { get; }

        /// <summary>
        /// The encoded string matchs SP REST requirement
        /// </summary>
        string EncodedUrl { get; }

        /// <summary>
        /// The method string for WebRequest
        /// </summary>
        string RequestMethod { get; }
    }
}
