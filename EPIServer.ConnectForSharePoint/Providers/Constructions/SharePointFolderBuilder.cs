using EPiServer.Core;
using EPiServer.Framework;
using EPiServer.ServiceLocation;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using System;

namespace EPiServer.ConnectForSharePoint.Providers.Constructions
{
    /// <summary>
    /// Builder for create SharePoint folder.
    /// </summary>
    [ServiceConfiguration(ServiceType = typeof(ISharePointDataBuilder))]
    public class SharePointFolderBuilder : SharePointDataBuilder<ContentFolder>
    {
        /// <summary>
        /// Create content from SharePoint list item.
        /// </summary>
        /// <param name="content">The icontent.</param>
        /// <param name="contentLink">The content link.</param>
        /// <param name="parentLink">The parent.</param>
        /// <param name="modelWrapper">The model wrapper item.</param>
        public override IContent Build(IContent content, ContentReference contentLink, ContentReference parentLink, IModelWrapper modelWrapper)
        {
            Validator.ThrowIfNull("content", content);
            Validator.ThrowIfNull("modelWrapper", modelWrapper);

            var basicContent = content as BasicContent;
            if (basicContent != null)
            {
                // Sharepoint views does not contain properties about date time created, changed, saved,
                // we need to assign fake data to support import/export content in CMS.
                basicContent.Created = basicContent.Changed = basicContent.Saved = DateTime.Now;
            }

            if (modelWrapper != null)
            {
                content.Name = modelWrapper.ListItem != null
                    ? modelWrapper.ListItem.DisplayName
                    : string.Format("{0} - {1}", modelWrapper.List.Title, modelWrapper.View.Title);
                
                var contentFolder = content as ContentFolder;
                if (contentFolder != null)
                {
                    contentFolder.RouteSegment = WebHelper.StripUrlSpecialCharacters(content.Name);
                }
            }

            return content;
        }

    }
}