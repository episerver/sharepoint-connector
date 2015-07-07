using EPiServer.Core;
using EPiServer.Framework;
using EPiServer.Framework.Blobs;
using EPiServer.ImageLibrary;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EPiServer.ConnectForSharePoint.Helpers;

namespace EPiServer.ConnectForSharePoint.Providers.Constructions
{
    /// <summary>
    /// Builder for creating SharePoint image.
    /// </summary>
    [ServiceConfiguration(ServiceType = typeof(ISharePointDataBuilder))]
    public class SharePointImageBuilder : SharePointDataBuilder<ImageData>
    {
        protected static readonly BlobFactory BlobFactory = ServiceLocator.Current.GetInstance<BlobFactory>();

        /// <summary>
        /// Creates the specified content link.
        /// </summary>
        /// <param name="content">The icontent.</param>
        /// <param name="contentLink">The content link.</param>
        /// <param name="parent">The parent.</param>
        /// <param name="modelWrapper">The model wrapper item.</param>
        public override IContent Build(IContent content, ContentReference contentLink, ContentReference parentLink, IModelWrapper modelWrapper)
        {
            Validator.ThrowIfNull("content", content);
            Validator.ThrowIfNull("modelWrapper", modelWrapper);
            Validator.ThrowIfNull("ListItem", modelWrapper.ListItem);

            var mediaData = content as MediaData;
            if (mediaData == null)
            {
                return content;
            }
            
            var blob = CreateSharePointBlob(mediaData, modelWrapper);
            blob.IsThumbnail = true;
            mediaData.Thumbnail = blob;

            return content;
        }

    }
}