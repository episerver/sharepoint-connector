using EPiServer.Core;

namespace EPiServer.ConnectForSharePoint.SPServices
{
    public interface ITagsService
    {
        /// <summary>
        /// Gets tags for list item.
        /// </summary>
        /// <param name="encodedAbsUrl">The SharePoint's content endcoded absolute url</param>
        /// <returns></returns>
        CategoryList GetTagsForListItem(string encodedAbsUrl);
    }
}