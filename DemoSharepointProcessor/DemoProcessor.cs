using EPiServer.ConnectForSharePoint.Messaging;
using System;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ServiceLocation;

namespace DemoSharepointProcessor
{
    [ServiceConfiguration(ServiceType = typeof(ISharePointProcessor))]
    public class DemoProcessor : ISharePointProcessor
    {
        public void AddDocument(SPListItem item)
        {
            item.Id = 42;
        }

        public void AddFolder(SPListItem item)
        {
            item.Id = 43;
        }

        public void DeleteDocument(SPListItem item)
        {
            item.Id = 44;
        }

        public void DeleteFolder(SPListItem item)
        {
            item.Id = 45;
        }

        public void UpdateDocument(SPListItem item)
        {
            item.Id = 46;
        }

        public void UpdateFolder(SPListItem item)
        {
            item.Id = 47;
        }
    }
}
