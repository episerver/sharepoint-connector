using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.Core;
using EPiServer.Logging;
using EPiServer.Shell.Services.Rest;
using System.Web.Mvc;

namespace EPiServer.ConnectForSharePoint.Rest
{
    [RestStore("sharepointcontent")]
    public class SharePointContentStore : RestControllerBase
    {
        private static readonly ILogger _logger = LogManager.GetLogger();

        /// <summary>
        /// Initializes the Category for a content if neccessary.
        /// </summary>
        /// <param name="contentId">The content identifier.</param>
        /// <returns></returns>
        [HttpGet]
        public RestResult Get(int contentId)
        {
            SharePointMediaProvider spMediaProvider = ContentProviderHelper.SharePointMediaProvider;
            var contentLink = new ContentReference(contentId, ContentProviderHelper.ContentProviderKey);

            IModelWrapper modelWrapper;
            var content = spMediaProvider.LoadContent(contentLink, out modelWrapper);

            if (content != null)
            {
                if (modelWrapper == null)
                {
                    modelWrapper = spMediaProvider.GetModelWrapper(contentLink);
                }

                if (modelWrapper != null)
                {
                    var categorizable = content as ICategorizable;
                    if (categorizable != null && categorizable.Category == null) // has not been initialized yet
                    {
                        _logger.Debug("Get tags for content id {0}", contentId);
                        categorizable.Category = SharePointHelper.SPServicesHub.SPTagsService.GetTagsForListItem(modelWrapper.ListItem.EncodedAbsUrl);
                        CacheManager.Insert(DataFactoryCache.PageMasterLanguageCacheKey(contentLink), content);
                    }
                }
            }

            return null;
        }
    }
}