using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.Framework.Blobs;
using EPiServer.ServiceLocation;
using System;
using System.IO;

namespace EPiServer.ConnectForSharePoint.Providers.Constructions
{
    public abstract class SharePointDataBuilder<T> : ISharePointDataBuilder where T : IContent, new()
    {
        public EPiServer.DataAbstraction.ContentType RequestedType { get; set; }

        /// <summary>
        /// Determine whether the builder can handle requrested type.
        /// </summary>
        /// <param name="requestedType">Type of the requested.</param>
        /// <returns>
        /// true if the builder can handle requested type, otherwise false.
        /// </returns>
        public virtual bool HandleType(Type requestedType)
        {
            if (typeof(T).IsAssignableFrom(requestedType))
            {
                RequestedType = ServiceLocator.Current.GetInstance<IContentTypeRepository>().Load(requestedType);
                return true;
            }
            return false;
        }

        /// <summary>
        /// Creates the specified content link.
        /// </summary>
        /// <param name="content">The icontent.</param>
        /// <param name="contentLink">The content link.</param>
        /// <param name="parentLink">The parent link.</param>
        /// <param name="modelWrapper">The model wrapper item.</param>
        public abstract IContent Build(IContent content, ContentReference contentLink, ContentReference parentLink, IModelWrapper modelWrapper);

        /// <summary>
        /// Create SharePoint blob for media/image file
        /// </summary>
        /// <param name="mediaData">The content as media data type </param>
        /// <param name="modelWrapper">The SharePoint model wrapper</param>
        /// <returns></returns>
        public virtual SharePointBlob CreateSharePointBlob(MediaData mediaData, IModelWrapper modelWrapper)
        {
            // workaround to create sharepoint blob since CreateBlob always return default blob provider
            var file = modelWrapper.ListItem.File;
            var uriString = string.Format("{0}/{1}", mediaData.BinaryDataContainer, file.Name);
            uriString = uriString.Replace("default", "sharepoint");
            // do the blob
            var blob = BlobFactory.Instance.GetProvider(new Uri("epi.fx.blob://sharepoint/"))
                .CreateBlob(new Uri(uriString), new FileInfo(file.Name).Extension) as SharePointBlob;
            blob.FileItem = modelWrapper.ListItem;
            blob.ContentLink = modelWrapper.ContentLink;
            modelWrapper.BlobUri = blob.ID;

            return blob;
        }
    }
}