using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    public class SPListItemMapper : IMapper<JsonListItem.D, SPListItem>
    {
        public SPListItem Map(JsonListItem.D source)
        {
            return new SPListItem()
            {
                Created = source.Created,
                FileSystemObjectType = source.FileSystemObjectType,
                GUID = Guid.Parse(source.GUID),
                Id = source.Id,
                Modified = source.Modified,
                Title = source.Title,
                FileUrl = source.__metadata.uri + "/File",
                FolderUrl = source.__metadata.uri + "/Folder",
                EncodedAbsUrl = source.EncodedAbsUrl,
                VideoSetShowEmbedLink = source.VideoSetShowEmbedLink,
                VideoSetShowDownloadLink = source.VideoSetShowDownloadLink 
            };
        }

        public IEnumerable<SPListItem> MapList(JsonListItem.D[] source)
        {
            return source.Select<JsonListItem.D, SPListItem>(jsData => Map(jsData));
        }
    }
}