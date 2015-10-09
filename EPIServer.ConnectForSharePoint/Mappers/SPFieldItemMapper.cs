using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;
using System.Collections.Generic;
using System.Linq;
using System;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    public class SPFieldItemMapper : IMapper<JsonFieldItem.D, SPFieldItem>
    {
        public SPFieldItem Map(JsonFieldItem.D source)
        {
            return new SPFieldItem()
            {
                DefaultValue = source.DefaultValue,
                Filterable = source.Filterable,
                Hidden = source.Hidden,
                InternalName = source.InternalName,
                StaticName = source.StaticName,
                Title = source.Title,
                Group = source.Group
            };
        }

        public IEnumerable<SPFieldItem> MapList(JsonFieldItem.D[] source)
        {
            return source.Select<JsonFieldItem.D, SPFieldItem>(jsData => Map(jsData));
        }
    }
}