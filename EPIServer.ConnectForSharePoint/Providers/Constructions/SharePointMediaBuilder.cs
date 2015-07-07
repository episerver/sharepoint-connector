using System;
using System.IO;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.Core;
using EPiServer.Framework;
using EPiServer.Framework.Blobs;
using EPiServer.ServiceLocation;

namespace EPiServer.ConnectForSharePoint.Providers.Constructions
{
    /// <summary>
    /// Base builder class for creating SharePoint media.
    /// </summary>
    [ServiceConfiguration(ServiceType = typeof(ISharePointDataBuilder))]
    public class SharePointMediaBuilder : SharePointDataBuilder<MediaData>
    {
        /// <summary>
        /// Create content from SharePoint list item.
        /// </summary>
        /// <param name="content">The icontent.</param>
        /// <param name="contentLink">The content link.</param>
        /// <param name="parentLink">The parent.</param>z
        /// <param name="modelWrapper">The model wrapper item.</param>
        public override IContent Build(IContent content, ContentReference contentLink, ContentReference parentLink, IModelWrapper modelWrapper)
        {
            Validator.ThrowIfNull("content", content);
            Validator.ThrowIfNull("modelWrapper", modelWrapper);
            Validator.ThrowIfNull("ListItem", modelWrapper.ListItem);

            if (modelWrapper.ListItem.File == null)
            {
                return content;
            }

            var file = modelWrapper.ListItem.File;
            content.Name = file.Name;

            var categorizable = content as ICategorizable;
            if (categorizable != null)
            {
                // TECHNOTE: Set the Category to null to indicate that it has not been initialized yet.
                // It will be initialized when this content is actually loaded in Edit View.
                categorizable.Category = null;
            }

            var basicContent = content as BasicContent;
            if (basicContent != null)
            {
                basicContent.Created = file.TimeCreated;
                basicContent.Changed = basicContent.Saved = file.TimeLastModified;
                basicContent.ChangedBy = file.ModifiedBy != null ? file.ModifiedBy.Title : string.Empty;
            }

            var versionable = content as IVersionable;
            if (versionable != null)
            {
                versionable.Status = VersionStatus.Published;
                versionable.StartPublish = file.TimeLastModified;
            }

            var mediaData = content as MediaData;
            if (mediaData == null)
            {
                return content;
            }
            
            var blob = CreateSharePointBlob(mediaData, modelWrapper);
            mediaData.BinaryData = blob;
            mediaData.RouteSegment = WebHelper.StripUrlSpecialCharacters(content.Name);

            return content;
        }

    }
}