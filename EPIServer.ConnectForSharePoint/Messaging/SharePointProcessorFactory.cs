using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ServiceLocation;
using System;
using System.Threading;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public class SharePointProcessorFactory
    {
        IServiceLocator _serviceLocator;

        public SharePointProcessorFactory()
        {
            _serviceLocator = ServiceLocator.Current;
        }

        internal SharePointProcessorFactory(IServiceLocator serviceLocator)
        {
            _serviceLocator = serviceLocator;
        }

        public ISharePointProcessor GetSharePointProcessor()
        {
            try
            {
                // Check for a plugin, if we dont have one we will return our default
                return _serviceLocator.GetInstance<ISharePointProcessor>();
            }
            catch
            {
                // Intentionally empty, if nothing is registered we were getting an 
                // exception thrown.
            }
            return new SharePointProcessor();
        }

        internal ISharePointProcessor GetSharePointProcessor(IContentProviderHelper contentHelper, 
            IFileStreamService fileStreamService, 
            IListsViewsService listViewService, IAccessRightsHelper accessRightsHelper)
        {
            var processor = _serviceLocator.GetInstance<ISharePointProcessor>();
            if (processor == null)
            {
                return new SharePointProcessor(_serviceLocator, contentHelper, fileStreamService, listViewService, accessRightsHelper);
            }
            else
            {
                return processor;
            }
        }
    }
}