using System;
using System.IO;

namespace EPiServer.ConnectForSharePoint.Models
{
    public class SPFile
    {
        public string CheckInComment { get; set; }
        public int CheckOutType { get; set; }
        public string ContentTag { get; set; }
        public int CustomizedPageStatus { get; set; }
        public string ETag { get; set; }
        public bool Exists { get; set; }
        public string Length { get; set; }
        public int Level { get; set; }
        public int MajorVersion { get; set; }
        public int MinorVersion { get; set; }
        public string Name { get; set; }
        public string ServerRelativeUrl { get; set; }
        public DateTime TimeCreated { get; set; }
        public DateTime TimeLastModified { get; set; }
        public string Title { get; set; }
        public int UIVersion { get; set; }
        public string UIVersionLabel { get; set; }
        public SPListItem ListItem { get; set; }
        public SPUser Author { get; set; }
        public SPUser ModifiedBy { get; set; }
    }
}