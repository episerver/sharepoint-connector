using EPiServer.ConnectForSharePoint.Models.FileTypes;
using EPiServer.Core;
using EPiServer.ServiceLocation;
using System;

namespace EPiServer.ConnectForSharePoint
{
    [ServiceConfiguration(ServiceType = typeof(SPContentMediaResolver), Lifecycle = ServiceInstanceScope.HttpContext)]
    public class SPContentMediaResolver : ContentMediaResolver
    {
        public override Type GetFirstMatching(string extension)
        {
            var selectedAsDefault = base.GetFirstMatching(extension);
            // For all image data, I want to use class SharepointImageFile
            if (typeof(ImageData).IsAssignableFrom(selectedAsDefault))
                return typeof(SharepointImageFile);

            // For all video data, I want to use class SharepointVideoFile
            if (typeof(VideoData).IsAssignableFrom(selectedAsDefault))
                return typeof(SharepointVideoFile);

            // Is this ok to just always return our generic media?
            return typeof(SharepointGenericMedia);
        }
    }
}