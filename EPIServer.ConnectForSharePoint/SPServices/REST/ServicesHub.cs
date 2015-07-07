using EPiServer.ConnectForSharePoint.Mappers;
using EPiServer.ConnectForSharePoint.Rest;
using EPiServer.ServiceLocation;

namespace EPiServer.ConnectForSharePoint.SPServices.REST
{
    [ServiceConfiguration(ServiceType = typeof(IServicesHub))]
    public class ServicesHub : RESTServiceBase, IServicesHub
    {
        /// <summary>
        /// SharePoint's endpoint context service
        /// </summary>
        public IEndpointContextService<SPRestClient> SPEndpointContextService
        {
            get
            {
                return EndpointContextService;
            }
        }

        private IListsViewsService _spListsViewsService;
        /// <summary>
        /// SharePoint's lists and views service
        /// </summary>
        public IListsViewsService SPListsViewsService
        {
            get
            {
                if (_spListsViewsService == null)
                {
                    try
                    {
                        ServiceLocator.Current.TryGetExistingInstance<IListsViewsService>(out _spListsViewsService);
                    }
                    catch { }

                    if (_spListsViewsService == null)
                    {
                        _spListsViewsService = new ListsViewsService();
                        ServiceLocator.Current.Buildup(_spListsViewsService);
                    }
                }


                return _spListsViewsService;
            }
        }

        private IFilesFoldersService _spFilesFoldersService;
        /// <summary>
        /// SharePoint's files and folders service
        /// </summary>
        public IFilesFoldersService SPFilesFoldersService
        {
            get
            {
                if (_spFilesFoldersService == null)
                {
                    try
                    {
                        ServiceLocator.Current.TryGetExistingInstance<IFilesFoldersService>(out _spFilesFoldersService);
                    }
                    catch { }

                    if (_spFilesFoldersService == null)
                    {
                        _spFilesFoldersService = new FilesFoldersService();
                        ServiceLocator.Current.Buildup(_spFilesFoldersService);
                    }
                }


                return _spFilesFoldersService;
            }
        }

        private IFileStreamService _spFileStreamService;
        /// <summary>
        /// The SharePoint's file stream service
        /// </summary>
        public IFileStreamService SPFileStreamService
        {
            get
            {
                if (_spFileStreamService == null)
                {
                    try
                    {
                        ServiceLocator.Current.TryGetExistingInstance<IFileStreamService>(out _spFileStreamService);
                    }
                    catch { }

                    if (_spFileStreamService == null)
                    {
                        _spFileStreamService = new FileStreamService();
                        ServiceLocator.Current.Buildup(_spFileStreamService);
                    }
                }

                return _spFileStreamService;
            }
        }

        private ITagsService _spTagsService;
        /// <summary>
        /// SharePoint's tags service
        /// </summary>
        public ITagsService SPTagsService
        {
            get
            {
                if (_spTagsService == null)
                {
                    try
                    {
                        ServiceLocator.Current.TryGetExistingInstance<ITagsService>(out _spTagsService);
                    }
                    catch { }

                    if (_spTagsService == null)
                    {
                        _spTagsService = new TagsService();
                        ServiceLocator.Current.Buildup(_spTagsService);
                    }
                }


                return _spTagsService;
            }
        }
    }
}