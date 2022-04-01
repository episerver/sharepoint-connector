using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Processor;
using EPiServer.Logging;
using EPiServer.ServiceLocation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics.CodeAnalysis;
using System.Threading;

namespace EPiServer.ConnectForSharePoint.Test
{
    [TestClass]
    [ExcludeFromCodeCoverage]
    public class ProcessingThreadManagerTests
    {
        private ProcessingThreadManager threadManager;
        private readonly Mock<IServiceLocator> serviceLocator = new Mock<IServiceLocator>();
        private readonly Mock<ISharePointProcessor> processor = new Mock<ISharePointProcessor>();
        private readonly Mock<ILogger> logger = new Mock<ILogger>();
        private static readonly object SyncRoot = new object();

        public ProcessingThreadManagerTests()
        {
            serviceLocator = new Mock<IServiceLocator>();
            processor = new Mock<ISharePointProcessor>();
            logger = new Mock<ILogger>();

            serviceLocator.Setup(s => s.GetInstance<ILogger>()).Returns(logger.Object);

            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = null;
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetries"] = null;
        }

        [TestInitialize]
        public void TestInitialize()
        {
            Monitor.Enter(SyncRoot);
        }

        [TestCleanup]
        public void cleanup()
        {
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = null;
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetries"] = null;

            Monitor.Exit(SyncRoot);
        }

        #region Message Thread Tests
        [TestMethod]
        public void Uses_Default_ThreadCount_WhenNotSpecifiedInConfig()
        {
            int expectedNumberOfThreads = ProcessingThreadManager.DefaultThreadCount;
            int actualNumberOfThreads = 0;
            int testSleepTime = 5000;

            threadManager = new ProcessingThreadManager(serviceLocator.Object);

            processor.Setup(p => p.AddDocument(It.IsAny<SPListItem>())).Callback(() => {
                actualNumberOfThreads++;
                Thread.Sleep(testSleepTime);
            });

            var messageList = new List<SharePointContentAddMessage>();
            for (var x = 0; x < expectedNumberOfThreads + 1; x++)
            {
                messageList.Add(new SharePointContentAddMessage() { Item = new SPListItem() { Title = "Test" + x} });
            }

            Thread addmessagesThread = new Thread(() =>
            {
                foreach (var message in messageList)
                {
                    threadManager.AddThread(processor.Object, message);
                }
            });
            addmessagesThread.Start();
            addmessagesThread.Join( new TimeSpan(0,0,1) );

            Assert.AreEqual(expectedNumberOfThreads, actualNumberOfThreads);
        }

        [TestMethod]
        public void Uses_ThreadCount_Specified_InConfig()
        {
            int expectedNumberOfThreads = ProcessingThreadManager.DefaultThreadCount - 2;
            int actualNumberOfThreads = 0;
            int testSleepTime = 5000;

            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = expectedNumberOfThreads.ToString();
            threadManager = new ProcessingThreadManager(serviceLocator.Object);

            processor.Setup(p => p.AddDocument(It.IsAny<SPListItem>())).Callback(() => {
                actualNumberOfThreads++;
                Thread.Sleep(testSleepTime);
            });

            var messageList = new List<SharePointContentAddMessage>();
            for (var x = 0; x < expectedNumberOfThreads + 1; x++)
            {
                messageList.Add(new SharePointContentAddMessage() { Item = new SPListItem() { Title = "Test" + x } });
            }

            Thread addmessagesThread = new Thread(() =>
            {
                foreach (var message in messageList)
                {
                    threadManager.AddThread(processor.Object, message);
                }
            });
            addmessagesThread.Start();
            addmessagesThread.Join(new TimeSpan(0, 0, 1));

            Assert.AreEqual(expectedNumberOfThreads, actualNumberOfThreads);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ThrowsException_If_MessageThreads_ToLarge()
        {
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = (ProcessingThreadManager.MaxThreadCount + 1).ToString();
            threadManager = new ProcessingThreadManager(serviceLocator.Object);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ThrowsException_If_MessageThreads_ToSmall()
        {
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = "0";
            threadManager = new ProcessingThreadManager(serviceLocator.Object);
        }

        #endregion Message Thread Tests
    }
}