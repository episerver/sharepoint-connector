using EPiServer.ConnectForSharePoint.Messaging;
using System;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ServiceLocation;

namespace DemoSharepointProcessor
{
    [ServiceConfiguration(ServiceType = typeof(ISharePointProcessor))]
    public class DemoProcessor : ISharePointProcessor
    {
        private SharePointProcessor dp;

        // Default processor that adds documents and folders to EPiServer (optional)
        private SharePointProcessor defaultProcessor {
            get
            {
                if(dp == null)
                    dp = new SharePointProcessor();
                return dp;
            }
        }


        public void AddDocument(SPListItem item)
        {
            // Do some custom processing

            // The splistitem does not contain the data stream associated with the file.
            // The following line is an example of how to get the data associated with the SPListItem
            var data = defaultProcessor.getFileStream(item);

            // Optionally call the default processor - note that the default processor 
            // also gets the data stream so that the data can be inserted into EPiServer
            defaultProcessor.AddDocument(item);
        }

        public void AddFolder(SPListItem item)
        {
            // Do some custom processing

            // Optionally call the default processor
            defaultProcessor.AddFolder(item);
        }

        public void DeleteDocument(SPListItem item)
        {
            // Do some custom processing

            // Optionally call the default processor
            defaultProcessor.DeleteDocument(item);
        }

        public void DeleteFolder(SPListItem item)
        {
            // Do some custom processing

            // Optionally call the default processor
            defaultProcessor.DeleteFolder(item);
        }

        public void UpdateDocument(SPListItem item)
        {
            // Do some custom processing

            // Optionally call the default processor
            defaultProcessor.UpdateDocument(item);
        }

        public void UpdateFolder(SPListItem item)
        {
            // Do some custom processing

            // Optionally call the default processor
            defaultProcessor.UpdateFolder(item);
        }
    }
}
