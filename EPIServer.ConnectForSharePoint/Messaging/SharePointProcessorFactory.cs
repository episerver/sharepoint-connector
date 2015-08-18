using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ServiceLocation;
using System;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public class SharePointProcessorFactory
    {
        public SharePointProcessorFactory()
        {

        }

        public ISharePointProcessor GetSharePointProcessor()
        {
            return new SharePointProcessor();
        }

        internal ISharePointProcessor GetSharePointProcessor(IServiceLocator serviceLocator, IContentProviderHelper contentHelper, IFileStreamService fileStreamService)
        {
            return new SharePointProcessor(serviceLocator, contentHelper, fileStreamService);
        }
    }
}