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

        internal ISharePointProcessor GetSharePointProcessor(IContentRepository repository)
        {
            return new SharePointProcessor(repository);
        }
    }
}