using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPiServer.ConnectForSharePoint.Models
{
    /// <summary>
    /// Stores SharePoint files/folders. 
    /// </summary>
    public class DataMapping
    {
        /// <summary>
        /// The SharePoint files are stored in a dictionary with the key is the content link of content corresponding to the parent folder,
        /// the value is a list of <see cref="ModelWrapper"/> mapped with all files which are the direct children of the parent folder
        /// </summary>
        public Dictionary<string, List<IModelWrapper>> Files;
        
        /// <summary>
        /// The SharePoint folders are stored in a dictionary with the key is the content link of content corresponding to the parent folder,
        /// the value is a list of <see cref="ModelWrapper" /> mapped with all folder which are the direct children of the parent folder
        /// </summary>
        public Dictionary<string, List<IModelWrapper>> Folders;

        public DataMapping()
            : this(new Dictionary<string, List<IModelWrapper>>(), new Dictionary<string, List<IModelWrapper>>())
        {
        }

        public DataMapping(Dictionary<string, List<IModelWrapper>> files, Dictionary<string, List<IModelWrapper>> folders)
        {
            Files = files;
            Folders = folders;
        }
    }
}