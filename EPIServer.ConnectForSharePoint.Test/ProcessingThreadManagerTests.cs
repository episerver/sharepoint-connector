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
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetries"] = null;
        }

        [TestCleanup]
        public void cleanup()
        {
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = null;
            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetries"] = null;
        }

        #region Message Thread Tests
        [TestMethod]
        public void Uses_Default_MessageThreads_Value_WhenNot_InConfig()
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
        public void Uses_MessageThreads_Config_Value()
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

            Thread.Sleep(2000);

            Assert.AreEqual(expectedNumberOfThreads, actualNumberOfThreads);
        }

        [TestMethod]
        public void ThrowsException_If_MessageThreads_ToLarge()
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
        public void ThrowsException_If_MessageThreads_ToSmall()
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

        #region Message Retry Tests

        [TestMethod]
        public void ThrowsException_If_MessageRetry_ToLarge()
        {
            try
            {
                ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetries"] = "9";
                threadManager = new ProcessingThreadManager(serviceLocator.Object);
                Assert.Fail();
            }
            catch (ArgumentException e)
            {

            }
        }

        [TestMethod]
        public void ThrowsException_If_MessageRetrys_ToSmall()
        {
            try
            {
                ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetries"] = "0";
                threadManager = new ProcessingThreadManager(serviceLocator.Object);
                Assert.Fail();
            }
            catch (ArgumentException e)
            {

            }
        }

        [TestMethod]
        public void Retrys_ExpectedNumberOfDefaultTimes()
        {
            int expectedNumberOfRetrys = 3;
            int actualNumberOfRetrys = 0;
            int testSleepTime = 5000;

            threadManager = new ProcessingThreadManager(serviceLocator.Object);

            processor.Setup(p => p.AddDocument(It.IsAny<SPListItem>())).Callback(() => {
                actualNumberOfRetrys++;
                throw new Exception("oh oh");
            });

            threadManager.AddThread(processor.Object, new SharePointContentAddMessage() {  Item = new SPListItem() {  Title = "Title" } });

            Thread.Sleep(1000);

            Assert.AreEqual(expectedNumberOfRetrys, actualNumberOfRetrys);
        }

        [TestMethod]
        public void Retrys_ExpectedNumberOfTimes_withConfigValue()
        {
            int expectedNumberOfRetrys = 6;
            int actualNumberOfRetrys = 0;
            int testSleepTime = 5000;

            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetries"] = expectedNumberOfRetrys.ToString();

            threadManager = new ProcessingThreadManager(serviceLocator.Object);

            processor.Setup(p => p.AddDocument(It.IsAny<SPListItem>())).Callback(() => {
                actualNumberOfRetrys++;
                throw new Exception("oh oh");
            });

            threadManager.AddThread(processor.Object, new SharePointContentAddMessage() { Item = new SPListItem() { Title = "Title" } });

            Thread.Sleep(1000);

            Assert.AreEqual(expectedNumberOfRetrys, actualNumberOfRetrys);
        }
        #endregion Message Retry Tests
    }
}