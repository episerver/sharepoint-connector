using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    public class SPViewMapper : IMapper<JsonView.D, SPListView>
    {
        public SPListView Map(JsonView.D source)
        {
            return new SPListView()
            {
                Hidden = source.Hidden,
                Id = Guid.Parse(source.Id),
                Title = source.Title,
                ListViewXml = source.ListViewXml,
                ViewQuery = source.ViewQuery,
                IncludeRootFolder = source.IncludeRootFolder,
                Scope = source.Scope
            };
        }

        public IEnumerable<SPListView> MapList(JsonView.D[] source)
        {
            return source.Select<JsonView.D, SPListView>(jsData => Map(jsData));
        }
    }
}