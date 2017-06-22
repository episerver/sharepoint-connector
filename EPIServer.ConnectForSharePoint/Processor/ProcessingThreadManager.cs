using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Handlers;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using System;
using System.Collections.Generic;
using System.Threading;

namespace EPiServer.ConnectForSharePoint.Processor
{
    public class ProcessingThreadManager
    {
        private int MAX_THREADS = 8;
        private List<AutoResetEvent> eventList = new List<AutoResetEvent>();
        private ILogger _logger;

        class ThreadData
        {
            public AutoResetEvent resetEvent;
            public SharePointMessage message;
            public ISharePointProcessor processor;
        }

        public ProcessingThreadManager()
        {
            _logger = LogManager.GetLogger();
        }

        ProcessingThreadManager(IServiceLocator serviceLocator)
        {
            _logger = serviceLocator.GetInstance<ILogger>();
        }

        public void AddThread(ISharePointProcessor p, SharePointMessage mes)
        {
            if(eventList.Count == MAX_THREADS)
            {
                // Wait for them all to end. 
                WaitHandle.WaitAll(eventList.ToArray());
                eventList.Clear();
            }
            AutoResetEvent rs = new AutoResetEvent(false);
            eventList.Add( rs );
            ThreadPool.QueueUserWorkItem(new WaitCallback(WorkerThread), 
                new ThreadData { resetEvent = rs, message = mes, processor = p });
        }

        // To support unit tests
        internal void WaitAll()
        {
            WaitHandle.WaitAll(eventList.ToArray());
        }

        private void WorkerThread(Object o)
        {
            ThreadData td = o as ThreadData;
            int MAX_RETRY = 3;
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
                    if (retry > MAX_RETRY)
                    {
                        td.resetEvent.Set(); // event set, and done
                        _logger.Error("WorkerThread - failed after " + MAX_RETRY + " attempts.", ex);
                        throw;
                    }
                    _logger.Warning("Will retry processing method. Retry count = " + retry + " for item : " + td.message.Item.DisplayName);
                }
            } while (retry > 0 && retry < MAX_RETRY);
            td.resetEvent.Set(); // event set, and done
        }
    }
}