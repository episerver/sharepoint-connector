using EPiServer.ConnectForSharePoint.Models;
using System.Collections.Generic;

namespace EPiServer.ConnectForSharePoint.Helpers
{
    public interface ISharePointMessageHelper
    {
        /// <summary>
        /// This is where the main processing begins. 
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        void BeginProcessing(string listids, IEnumerable<SPList> splists);

        /// <summary>
        /// Called by the UI when lists are removed and the 
        /// configuration is saved. 
        /// </summary>
        /// <param name="splists"></param>
        /// <param name="listids"></param>
        void QueueRootFoldersForDelete(IEnumerable<SPList> splists, IEnumerable<string> listids = null);
    }
}
