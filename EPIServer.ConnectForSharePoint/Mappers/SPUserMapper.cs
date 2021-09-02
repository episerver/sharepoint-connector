using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest.Models;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    public class SPUserMapper : IMapper<JsonUser.D, SPUser>
    {
        public SPUser Map(JsonUser.D source)
        {
            if (source == null)
            {
                return null;
            }

            return new SPUser()
            {
                Title = source.Title,
                Id = source.Id,
                Email = source.Email
            };
        }

        public IEnumerable<SPUser> MapList(JsonUser.D[] source)
        {
            return source.Select<JsonUser.D, SPUser>(jsData => Map(jsData));
        }
    }
}