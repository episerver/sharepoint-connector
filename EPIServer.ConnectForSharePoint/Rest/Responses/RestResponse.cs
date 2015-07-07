using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPiServer.ConnectForSharePoint.Rest.Responses
{
    /// <summary>
    /// Base class for response received when invoking REST.
    /// </summary>
    public class RestResponse<TResult>
    {  
        public D<TResult> d { get; set; }
    }

    public class D<TResult>
    {
        public List<TResult> results { get; set; }
    }
}
