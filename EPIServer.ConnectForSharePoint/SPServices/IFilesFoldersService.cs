using System.Collections.Generic;
using EPiServer.ConnectForSharePoint.Models;

namespace EPiServer.ConnectForSharePoint.SPServices
{
    public interface IFilesFoldersService
    {
        /// <summary>
        /// Load and assign the <see cref="SPFolder"/> pointed by a <see cref="SPListItem"/>
        /// </summary>
        /// <param name="listItem">The associate list item.</param>
        SPFolder LoadFolder(SPListItem listItem);

        /// <summary>
        /// Load and assign the <see cref="SPFile"/> pointed by a <see cref="SPListItem"/>
        /// </summary>
        /// <param name="listItem">The associate list item.</param>
        SPFile LoadFile(SPListItem listItem);

        /// <summary>
        /// Load and assign ParentFolder to current <see cref="SPFolder"/>
        /// </summary>
        /// <param name="currentFolder">The folder to load its parent folder.</param>
        SPFolder LoadParentFolder(SPFolder currentFolder);

        /// <summary>
        /// Load and assign RootFolder to current <see cref="SPList"/>
        /// </summary>
        /// <param name="currentList">The list to load its root folder.</param>
        SPFolder LoadRootFolder(SPList currentList);

        /// <summary>
        /// Load and re-assign properties value to current <see cref="SPFolder"/>
        /// </summary>
        /// <param name="currentFolder">The folder for updating its properties.</param>
        SPFolder LoadFolderProperties(SPFolder currentFolder);

        /// <summary>
        /// Get files in the current <see cref="SPFolder"/>
        /// </summary>
        /// <param name="currentFolder">The folder to get its files.</param>
        IEnumerable<SPFile> GetFiles(SPFolder currentFolder);
    }
}