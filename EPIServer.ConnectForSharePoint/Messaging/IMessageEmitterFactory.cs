using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Messaging
{
    public interface IMessageEmitterFactory
    {
        IMessageEmitter Create();
    }
}
