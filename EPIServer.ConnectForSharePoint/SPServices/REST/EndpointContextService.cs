using EPiServer.ConnectForSharePoint.Rest;
using EPiServer.ServiceLocation;

namespace EPiServer.ConnectForSharePoint.SPServices.REST
{
    [ServiceConfiguration(ServiceType = typeof(IEndpointContextService<SPRestClient>))]
    public class EndpointContextService : IEndpointContextService<SPRestClient>
    {
        private SPRestClient _endpointContext;
        public SPRestClient EndpointContext
        {
            get
            {
                if (_endpointContext == null)
                {
                    _endpointContext = new SPRestClient();
                }

                return _endpointContext;
            }
            set {
                _endpointContext = value;
            }
        }

 
    }
}