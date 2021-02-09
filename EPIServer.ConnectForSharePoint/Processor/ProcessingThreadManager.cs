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
        internal static int DefaultMaxRetries = 3;
        private int maxThreads = DefaultThreadCount;
        private int maxRetries = DefaultMaxRetries;

        private readonly ILogger logger;
        private readonly object addThreadLock = new Object();
        private readonly List<AutoResetEvent> resetEventList = new List<AutoResetEvent>();

        class ThreadData
        {
            public AutoResetEvent resetEvent;
            public SharePointMessage message;
            public ISharePointProcessor processor;
        }

        public ProcessingThreadManager()
        {
            logger = LogManager.GetLogger();
            GetOptions();
        }

        internal ProcessingThreadManager(IServiceLocator serviceLocator)
        {
            logger = serviceLocator.GetInstance<ILogger>();
            GetOptions();
        }

        private void GetOptions()
        {
            var threadSettings = ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"];
            if (threadSettings != null)
            {
                int.TryParse(threadSettings, out maxThreads);
                if(maxThreads <= 0 || maxThreads > 8)
                {
                    throw new ArgumentException("maxThreads must be > 1 and less than 9");
                }
            }

            var retrySettings = ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetries"];
            if (retrySettings != null)
            {
                int.TryParse(retrySettings, out maxRetries);
                if (maxRetries <= 0 || maxRetries > 8)
                {
                    throw new ArgumentException("maxRetries must be > 1 and less than 9");
                }
            }

            logger.Debug("Sharepoint thread count : " + maxThreads);
            logger.Debug("Sharepoint max retries : " + maxRetries);
        }

        public void AddThread(ISharePointProcessor p, SharePointMessage mes)
        {
            lock (addThreadLock)
            {
                if (resetEventList.Count == maxThreads)
                {
                    int i = WaitHandle.WaitAny(resetEventList.ToArray());
                    resetEventList.Remove(resetEventList[i]);
                }
                AutoResetEvent rs = new AutoResetEvent(false);
                resetEventList.Add(rs);
                ThreadPool.QueueUserWorkItem(new WaitCallback(WorkerThread),
                    new ThreadData { resetEvent = rs, message = mes, processor = p });
            }
        }

        // To support unit tests
        internal void WaitAll()
        {
            WaitHandle.WaitAll(resetEventList.ToArray());
        }

        private void WorkerThread(Object o)
        {
            ThreadData td = o as ThreadData;
            int retry = 0;

            using (var pp = new PrincipalPreserver())
            {
                do
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
                        else if (td.message is SharePointContentUpdateMessage)
                        {
                            td.processor.UpdateDocument(td.message.Item);
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        retry++;
                        if (retry > maxRetries)
                        {
                            td.resetEvent.Set(); // event set, and done
                            logger.Error("Failed after " + maxRetries + " attempts.", ex);
                            throw;
                        }
                        logger.Warning("Will retry processing method. Retry count = " + retry + " for item : " + td.message.Item.DisplayName);
                    }
                } while (retry > 0 && retry < maxRetries);
            } 

            td.resetEvent.Set(); // event set, and done
        }
    }
}