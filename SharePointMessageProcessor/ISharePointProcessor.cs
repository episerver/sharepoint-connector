using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public interface ISharePointProcessor
    {
        void AddDocument(string jsonData);
        void AddFolder(string jsonData);
        void DeleteDocument(string jsonData);
        void DeleteFolder(string jsonData);
    }
}
