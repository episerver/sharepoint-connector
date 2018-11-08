using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.SPServices.REST;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPiServer.ConnectForSharePoint.SPServices
{
    [ServiceConfiguration(ServiceType = typeof(ITagsService))]
    public class TagsService : RESTServiceBase, ITagsService
    {
        private static readonly ILogger _logger = LogManager.GetLogger();
        private CategoryRepository _categoryRepository;

        private SocialDataService _socialDataService;

        public TagsService()
        {
            _categoryRepository = ServiceLocator.Current.GetInstance<CategoryRepository>();
            InitSocialDataService();
        }

        /// <summary>
        /// Gets tags for list item.
        /// </summary>
        /// <param name="encodedAbsUrl">The SharePoint's content endcoded absolute url</param>
        /// <returns></returns>
        public CategoryList GetTagsForListItem(string encodedAbsUrl)
        {
            try
            {
                // TECHNOTE:
                //  GetTagTermsOnUrl() method required input parameter that is an URL string (not in encoded format).
                //  So that, we should decode the given encoded URL before pass it.
                //  For example:
                //      The given encoded URL is: "http://MySharePointServer/MyAssets/My%27s%20Image.jpg"
                //      Should decode to: "http://MySharePointServer/MyAssets/My's Image.jpg"
                var tagsOnUrl = _socialDataService.GetTagTermsOnUrl(Url.Decode(encodedAbsUrl), 999);
                if (!tagsOnUrl.Any())
                {
                    return new CategoryList();
                }

                var tagNames = tagsOnUrl.Select(term => term.Term.Name);
                if (!tagNames.Any())
                {
                    return new CategoryList();
                }

                var categoryIds = new List<int>();
                foreach (var tag in tagNames)
                {
                    var category = _categoryRepository.Get(tag);
                    if (category != null)
                    {
                        categoryIds.Add(category.ID);
                    }
                }

                // Sort the category Ids before initializing the CategoryList. This should be done inside the CategoryList itself
                // but somehow this is missing.
                categoryIds.Sort();

                return new CategoryList(categoryIds.ToArray());
            }
            catch (Exception ex)
            {
                _logger.Error("Error when getting SharePoint Tags and Notes", ex);

                return new CategoryList();
            }
        }

        /// <summary>
        /// Initializes the SocialDataService for current SharePoint site.
        /// </summary>
        private void InitSocialDataService()
        {
            if (_socialDataService == null)
            {
                _socialDataService = new SocialDataService();
            }

            if (SharePointHelper.IsConnectToSharePointOnline(SPSettings.Username))
            {
                _socialDataService.UseDefaultCredentials = false;
                _socialDataService.CookieContainer = SharePointHelper.GetCookies(SPSettings);
            }
            else
            {
                _socialDataService.Credentials = SharePointHelper.GetCredentialCache(SPSettings);
            }

            _socialDataService.Url = SPSettings.SiteUrl + "/_vti_bin/SocialDataService.asmx";
        }
    }
}