using EPiServer.ConnectForSharePoint.Rest;

namespace EPiServer.ConnectForSharePoint.SPServices
{
    public interface IServicesHub
    {
        IEndpointContextService<SPRestClient> SPEndpointContextService { get; }

        IListsViewsService SPListsViewsService { get; }

        IFilesFoldersService SPFilesFoldersService { get; }

        IFileStreamService SPFileStreamService { get; }

        ITagsService SPTagsService { get; }
    }
}