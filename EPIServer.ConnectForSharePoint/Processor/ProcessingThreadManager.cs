using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Handlers;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Threading;

namespace EPiServer.ConnectForSharePoint.Processor
{
    public class ProcessingThreadManager
    {
        internal static int DefaultThreadCount = 4;

        private int maxThreads = DefaultThreadCount;
        private int maxRetries = 3;
        private readonly List<AutoResetEvent> eventList = new List<AutoResetEvent>();
        private readonly ILogger _logger;

        class ThreadData
        {
            public AutoResetEvent resetEvent;
            public SharePointMessage message;
            public ISharePointProcessor processor;
        }

        public ProcessingThreadManager()
        {
            _logger = LogManager.GetLogger();
//            GetOptions();
        }

        internal ProcessingThreadManager(IServiceLocator serviceLocator)
        {
            _logger = serviceLocator.GetInstance<ILogger>();
//            GetOptions();
        }

        //private void GetOptions()
        //{
        //    var setting = ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"]?.ToString();
        //    if (setting != null)
        //    {
        //        _ = int.TryParse(setting, out maxThreads);
        //    }

        //    setting = ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetrys"]?.ToString();
        //    if (setting != null)
        //    {
        //        _ = int.TryParse(setting, out maxRetries);
        //    }

        //    _logger.Debug("Sharepoint thread count : " + maxThreads);
        //    _logger.Debug("Sharepoint max retries : " + maxRetries);
        //}

        object AddThreadLock = new Object();

        public void AddThread(ISharePointProcessor p, SharePointMessage mes)
        {
            lock (AddThreadLock)
            {
                if (eventList.Count == maxThreads)
                {
                    int i = WaitHandle.WaitAny(eventList.ToArray());
                    eventList.Remove(eventList[i]);
                }
                AutoResetEvent rs = new AutoResetEvent(false);
                eventList.Add(rs);
                ThreadPool.QueueUserWorkItem(new WaitCallback(WorkerThread),
                    new ThreadData { resetEvent = rs, message = mes, processor = p });
            }
        }

        // To support unit tests
        internal void WaitAll()
        {
            WaitHandle.WaitAll(eventList.ToArray());
        }

        private void WorkerThread(Object o)
        {
            ThreadData td = o as ThreadData;
            int retry = 0;

            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    if (td.message is SharePointListItemAddMessage)
                    {
                        td.processor.AddListItem(td.message.Item);
                    }
                    else if (td.message is SharePointContentAddMessage)
                    {
                        td.processor.AddDocument(td.message.Item);
                    }
                    retry = 0;
                }
                catch (Exception ex)
                {
                    retry++;
                    if (retry > maxRetries)
                    {
                        td.resetEvent.Set(); // event set, and done
                        _logger.Error("WorkerThread - failed after " + maxRetries + " attempts.", ex);
                        throw;
                    }
                    _logger.Warning("Will retry processing method. Retry count = " + retry + " for item : " + td.message.Item.DisplayName);
                }
            } while (retry > 0 && retry < maxRetries) ;

            td.resetEvent.Set(); // event set, and done
        }
    }
}