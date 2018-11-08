using EPiServer.ConnectForSharePoint.Models;
using System.Collections.Generic;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface IMetaDataProcessor
    {
        /// <summary>
        /// This is where the main processing begins. 
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        void BeginProcessing(string listids, IEnumerable<SPList> splists);
    }
}
