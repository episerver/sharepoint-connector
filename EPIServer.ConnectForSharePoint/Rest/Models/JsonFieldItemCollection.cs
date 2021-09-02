using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    public class JsonFieldItemCollection
    {
        public D d { get; set; }
        public class D
        {
            public JsonFieldItem.D[] results { get; set; }
        }
    }
}