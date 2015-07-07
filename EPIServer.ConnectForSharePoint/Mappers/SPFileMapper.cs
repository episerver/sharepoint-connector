using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    public class SPFileMapper : IMapper<JsonFile.D, SPFile>
    {
        public SPFile Map(JsonFile.D source)
        {
            if (source == null || string.IsNullOrEmpty(source.ServerRelativeUrl))
            {
                return null;
            }

            return new SPFile()
            {
                Exists = source.Exists,
                Length = source.Length,
                Level = source.Level,
                Name = source.Name,
                TimeCreated = source.TimeCreated,
                TimeLastModified = source.TimeLastModified,
                Title = source.Title,
                ServerRelativeUrl = source.ServerRelativeUrl
            };
        }

        public IEnumerable<SPFile> MapList(JsonFile.D[] source)
        {
            return source.Select<JsonFile.D, SPFile>(jsData => Map(jsData));
        }
    }
}