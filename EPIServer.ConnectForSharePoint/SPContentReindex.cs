using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.Framework;
using EPiServer.Logging;
using EPiServer.Search;
using EPiServer.ServiceLocation;
using System.Collections.Generic;
using System.Linq;

namespace EPiServer.ConnectForSharePoint
{
    [ServiceConfiguration(ServiceType = typeof(SPContentReindex), FactoryMember = "CreateInstance")]
    public class SPContentReindex : ContentSearchHandler
    {
        private ILogger _logger = LogManager.GetLogger();
        private IContentRepository _contentRepository;

        public SPContentReindex(SearchHandler searchHandler, IContentRepository contentRepository, 
            IContentTypeRepository contentTypeRepository,  SearchIndexConfig searchIndexConfig)  
            : base (searchHandler, contentRepository, contentTypeRepository, searchIndexConfig)
        {
            Validator.ThrowIfNull("contentRepository", contentRepository);

            _contentRepository = contentRepository;
        }

        public new static SPContentReindex CreateInstance(SearchHandler searchHandler, IContentRepository contentRepository, 
            IContentTypeRepository contentTypeRepository, SearchIndexConfig searchIndexConfig)
        {
            return new SPContentReindex(searchHandler, contentRepository, contentTypeRepository, searchIndexConfig)
            {
                ServiceActive = SearchSettings.Config.Active
            };
        }

        /// <summary>
        /// Index the media content from SharePoint
        /// </summary>
        public override void IndexPublishedContent()
        {
            if (!ServiceActive)
            {
                _logger.Information("SharePoint contents are not indexed since the search service is configured to be not active.");
                return;
            }

            var settings = SharePointSettings.Current;
            if (ContentReference.IsNullOrEmpty(settings.EntryPoint))
            {
                _logger.Information("SharePoint contents are not indexed since there is no SharePoint entry point set.");
                return;
            }

            var queue = GetQueueItems(settings.EntryPoint);

            while (queue.Count() > 0)
            {
                var content = queue.Dequeue();
                if (content is ContentFolder)
                {
                    continue;
                }

                IVersionable versionStatus = content as IVersionable;
                // If the content supports version status, we check that we don't index any non published content.
                if (versionStatus == null || (versionStatus.Status == VersionStatus.Published))
                {
                    UpdateItem(content);
                }
            }
        }

        /// <summary>
        /// Queue children from a root content
        /// </summary>
        /// <param name="rootLink">Root content reference</param>
        /// <returns>a queue of items</returns>
        private Queue<IContent> GetQueueItems(ContentReference rootLink)
        { 
            var queue = new Queue<IContent>();
            var subLinks = _contentRepository.GetDescendents(rootLink);

            subLinks.Select<ContentReference, IContent>(clink => _contentRepository.Get<IContent>(clink)).ToList().ForEach(c => queue.Enqueue(c));

            return queue;
        }
    }
}