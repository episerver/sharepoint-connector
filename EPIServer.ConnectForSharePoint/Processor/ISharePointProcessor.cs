using EPiServer.ConnectForSharePoint.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public interface ISharePointProcessor
    {
        void AddDocument(SPListItem item);
        void AddFolder(SPListItem item);
        void UpdateDocument(SPListItem item);
        void UpdateFolder(SPListItem item);
        void DeleteDocument(SPListItem item);
        void DeleteFolder(SPListItem item);
    }
}
