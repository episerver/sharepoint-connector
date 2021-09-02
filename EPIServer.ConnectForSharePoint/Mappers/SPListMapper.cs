using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    public class SPListMapper : IMapper<JsonList.D, SPList>
    {
        public SPList Map(JsonList.D source)
        {
            return new SPList()
            {
                BaseTemplate = source.BaseTemplate,
                BaseType = source.BaseType,
                Created = source.Created,
                Description = source.Description,
                Hidden = source.Hidden,
                Id = Guid.Parse(source.Id),
                ItemCount = source.ItemCount,
                LastItemDeletedDate = source.LastItemDeletedDate,
                LastItemModifiedDate = source.LastItemModifiedDate,
                Title = source.Title
            };
        }

        public IEnumerable<SPList> MapList(JsonList.D[] source)
        {
            return source.Select<JsonList.D, SPList>(jsData => Map(jsData));
        }
    }
}