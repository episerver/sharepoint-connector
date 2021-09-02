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
        internal static int MaxThreadCount = 16;
        internal static int DefaultThreadCount = 4;
        internal static int DefaultWaitTimeSpanInMinutes = 5;

        private int maxThreads = DefaultThreadCount;

        private readonly ILogger logger;
        private readonly object addThreadLock = new Object();
        private readonly List<AutoResetEvent> resetEventList = new List<AutoResetEvent>();

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

        public void AddThread(ISharePointProcessor processor, SharePointMessage message)
        {
            lock (addThreadLock)
            {
                try
                {
                    if (resetEventList.Count == maxThreads)
                    {
                        var waitHandle = WaitHandle.WaitAny(resetEventList.ToArray(), new TimeSpan(0, DefaultWaitTimeSpanInMinutes, 0));
                        if (waitHandle != System.Threading.WaitHandle.WaitTimeout)
                        {
                            resetEventList.Remove(resetEventList[waitHandle]);

                            AddWorkerToQueue(processor, message);
                        }
                        else
                        {
                            logger.Warning("Too many long running threads. " + message.Item.DisplayName +
                                " will be processed the next time the job runs.");
                        }
                    }
                    else
                    {
                        AddWorkerToQueue(processor, message);
                    }
                }
                catch ( Exception exception )
                {
                    logger.Error("Sharepoint : Error adding message ", exception);
                }
            }
        }

        private void AddWorkerToQueue(ISharePointProcessor processor, SharePointMessage message)
        {
            var threadContext = new ThreadContext() { Message = message, Processor = processor };
            resetEventList.Add(threadContext.ResetEvent);
            ThreadPool.QueueUserWorkItem(new WaitCallback(WorkerThread), threadContext);
        }

        private void WorkerThread(Object context)
        {
            var threadContext = context as ThreadContext;
            using (var pp = new PrincipalPreserver())
            {
                try
                {
                    logger.Debug("Processing : " + threadContext.Message.Item.DisplayName);

                    if (threadContext.Message is SharePointListItemAddMessage)
                    {
                        threadContext.Processor.AddListItem(threadContext.Message.Item);
                    }
                    else if (threadContext.Message is SharePointContentAddMessage)
                    {
                        threadContext.Processor.AddDocument(threadContext.Message.Item);
                    }
                    else if (threadContext.Message is SharePointContentUpdateMessage)
                    {
                        threadContext.Processor.UpdateDocument(threadContext.Message.Item);
                    }
                }
                catch (Exception ex)
                {
                    logger.Error("Sharepoint: Failed to process request", ex);
                }
            }
            threadContext.ResetEvent.Set();
        }

        private void GetOptions()
        {
            var threadSettings = ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"];
            if (threadSettings != null)
            {
                int.TryParse(threadSettings, out maxThreads);
                if (maxThreads < 1 || maxThreads > MaxThreadCount)
                {
                    throw new ArgumentException("maxThreads must be between 1 and " + MaxThreadCount);
                }
            }

            logger.Debug("Sharepoint thread count: " + maxThreads + " processor count: " + Environment.ProcessorCount);
        }

        // To support unit tests
        internal void WaitAll()
        {
            WaitHandle.WaitAll(resetEventList.ToArray(), new TimeSpan(0,0,30));
        }

        /// <summary>
        /// ThreadContext contains data required by the thread processing the request.
        /// </summary>
        internal class ThreadContext
        {
            public AutoResetEvent ResetEvent { get; set; }
            public SharePointMessage Message { get; set; }
            public ISharePointProcessor Processor { get; set; }

            public int Retries { get; set; }

            public ThreadContext()
            {
                Retries = 0;
                ResetEvent = new AutoResetEvent(false);
            }
        }
    }
}