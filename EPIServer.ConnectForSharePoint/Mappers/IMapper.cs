using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Mappers
{
    /// <summary>
    /// Mapper interface
    /// </summary>
    /// <typeparam name="TSource">The source type to map from.</typeparam>
    /// <typeparam name="TTarget">The target type to map to.</typeparam>
    public interface IMapper<TSource, TTarget>
    {
        /// <summary>
        /// Maps the source type to the target type.
        /// </summary>
        /// <param name="source">The object to map from.</param>
        /// <returns>The mapped object.</returns>
        TTarget Map(TSource source);

        /// <summary>
        /// Maps the source type to a list of the target type.
        /// </summary>
        /// <param name="source">The array of object to map from.</param>
        /// <returns>An enumearble of mapped objects.</returns>
        IEnumerable<TTarget> MapList(TSource[] source);
    }
}