using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    public class SPFolderMapper : IMapper<JsonFolder.D, SPFolder>
    {
        public SPFolder Map(JsonFolder.D source)
        {
            if (source == null || string.IsNullOrEmpty(source.ServerRelativeUrl))
            {
                return null;
            }

            var folder = new SPFolder()
            {
                Name = source.Name,
                ItemCount = source.ItemCount,
                ServerRelativeUrl = source.ServerRelativeUrl,
                TimeCreated = source.TimeCreated,
                TimeLastModified = source.TimeLastModified,
            };

            return folder;
        }

        public IEnumerable<SPFolder> MapList(JsonFolder.D[] source)
        {
            return source.Select<JsonFolder.D, SPFolder>(jsData => Map(jsData));
        }
    }
}