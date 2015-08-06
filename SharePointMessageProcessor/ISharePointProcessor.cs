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
        void AddFolder();
        void DeleteDocument();
        void DeleteFolder();
    }
}
