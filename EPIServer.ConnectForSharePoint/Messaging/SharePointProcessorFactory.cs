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
    }
}