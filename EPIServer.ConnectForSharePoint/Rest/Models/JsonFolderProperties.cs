using System;

namespace EPiServer.ConnectForSharePoint.Rest.Models
{
    /// <summary>
    /// SharePoint - REST API
    ///     Folder resource
    ///         Represents the properties for a folder on a SharePoint Web site.
    /// Endpoint URI
    ///     http://<site url>/_api/web/getfolderbyserverrelativeurl('/<folder name>')/properties
    /// </summary>
    public class JsonFolderProperties
    {
        public D d { get; set; }

        public class D
        {
            public __Metadata __metadata { get; set; }
            public string vti_x005f_hassubdirs { get; set; }
            public int vti_x005f_listbasetype { get; set; }
            public string vti_x005f_listrequirecheckout { get; set; }
            public DateTime vti_x005f_dirlateststamp { get; set; }
            public string vti_x005f_replid { get; set; }
            public string ContentTypeId { get; set; }
            public string vti_x005f_rtag { get; set; }
            public string vti_x005f_isbrowsable { get; set; }
            public int vti_x005f_listservertemplate { get; set; }
            public string vti_x005f_modifiedby { get; set; }
            public string vti_x005f_parentid { get; set; }
            public string vti_x005f_title { get; set; }
            public int vti_x005f_level { get; set; }
            public string vti_x005f_isscriptable { get; set; }
            public string vti_x005f_listenablemoderation { get; set; }
            public int vti_x005f_foldersubfolderitemcount { get; set; }
            public string vti_x005f_isexecutable { get; set; }
            public int vti_x005f_metainfoversion { get; set; }
            public string vti_x005f_listenableminorversions { get; set; }
            public int vti_x005f_docstoretype { get; set; }
            public int vti_x005f_folderitemcount { get; set; }
            public DateTime vti_x005f_timelastmodified { get; set; }
            public DateTime vti_x005f_nexttolasttimemodified { get; set; }
            public string vti_x005f_candeleteversion { get; set; }
            public string vti_x005f_listenableversioning { get; set; }
            public string vti_x005f_listname { get; set; }
            public string vti_x005f_etag { get; set; }
            public DateTime vti_x005f_timecreated { get; set; }
        }
    }

}