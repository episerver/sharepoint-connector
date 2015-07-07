namespace EPiServer.ConnectForSharePoint.SPServices
{
    public interface IEndpointContextService<T>
    {
        /// <summary>
        /// Get default SharePoint's endpoint context object from current SharePointSettings
        /// </summary>
        T EndpointContext { get; }
    }
}