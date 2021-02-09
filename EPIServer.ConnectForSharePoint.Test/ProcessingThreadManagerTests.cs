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
using System.Threading;

namespace EPiServer.ConnectForSharePoint.Test
{
    [TestClass]
    public class ProcessingThreadManagerTests
    {
        ProcessingThreadManager threadManager;
        Mock<IServiceLocator> serviceLocator = new Mock<IServiceLocator>();
        Mock<ISharePointProcessor> processor = new Mock<ISharePointProcessor>();
        Mock<ILogger> logger = new Mock<ILogger>();

        public ProcessingThreadManagerTests()
        {
            serviceLocator = new Mock<IServiceLocator>();
            processor = new Mock<ISharePointProcessor>();
            logger = new Mock<ILogger>();

            serviceLocator.Setup(s => s.GetInstance<ILogger>()).Returns(logger.Object);

            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = null;
        }

        [TestCleanup]
        public void cleanup()
        {
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = null;
        }

        #region Message Thread Tests
        [TestMethod]
        public void Uses_Default_MessageThreads_Value_IfNotSpecified_InConfig()
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
                messageList.Add(new SharePointContentAddMessage() { });
            }

            Thread addmessagesThread = new Thread(() =>
            {
                foreach (var message in messageList)
                {
                    threadManager.AddThread(processor.Object, message);
                }
            });
            addmessagesThread.Start();
            while (!addmessagesThread.IsAlive) ;

            Thread.Sleep(1000);

            Assert.AreEqual(expectedNumberOfThreads, actualNumberOfThreads);
        }

        [TestMethod]
        public void Uses_MessageThreads_Config_Value_IfSpecified_InConfig()
        {
            int expectedNumberOfThreads = ProcessingThreadManager.DefaultThreadCount - 1;
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
                messageList.Add(new SharePointContentAddMessage() { });
            }

            Thread addmessagesThread = new Thread(() =>
            {
                foreach (var message in messageList)
                {
                    threadManager.AddThread(processor.Object, message);
                }
            });
            addmessagesThread.Start();
            while (!addmessagesThread.IsAlive) ;

            Thread.Sleep(1000);

            Assert.AreEqual(expectedNumberOfThreads, actualNumberOfThreads);
        }

        [TestMethod]
        public void Uses_MessageThreads_Max_Value_IfValueinConfigGreaterthanMx()
        {
            try
            {
                ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = "9";
                threadManager = new ProcessingThreadManager(serviceLocator.Object);
                Assert.Fail();
            }
            catch (ArgumentException e)
            {

            }
        }

        [TestMethod]
        public void IfMaxThreadValueValue_LessThan1_ThrowsArgumentException()
        {
            try
            {
                ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = "0";
                threadManager = new ProcessingThreadManager(serviceLocator.Object);
                Assert.Fail();
            }
            catch (ArgumentException e)
            {

            }
        }

        #endregion Message Thread Tests
    }
}