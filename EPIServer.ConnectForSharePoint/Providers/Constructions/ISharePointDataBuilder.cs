using EPiServer.Core;
using EPiServer.ConnectForSharePoint.Models;
using System;

namespace EPiServer.ConnectForSharePoint.Providers.Constructions
{
    /// <summary>
    /// Interface for implement SharePointDataBuilder.
    /// </summary>
    public interface ISharePointDataBuilder
    {
        /// <summary>
        /// Determine whether the builder can handle requrested type.
        /// </summary>
        /// <param name="requestedType">Type of the requested.</param>
        /// <returns>true if the builder can handle requested type, otherwise false.</returns>
        bool HandleType(Type requestedType);

        /// <summary>
        /// Create content from SharePoint list item.
        /// </summary>
        /// <param name="content">The icontent.</param>
        /// <param name="contentLink">The content link.</param>
        /// <param name="parent">The parent.</param>
        /// <param name="modelWrapper">The model wrapper item.</param>
        IContent Build(IContent content, ContentReference contentLink, ContentReference parentLink, IModelWrapper modelWrapper);
    }
}