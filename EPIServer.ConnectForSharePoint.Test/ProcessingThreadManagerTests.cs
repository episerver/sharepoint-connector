using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Processor;
using EPiServer.ServiceLocation;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Test
{
    [TestClass]
    public class ProcessingThreadManagerTests
    {
        ProcessingThreadManager threadManager;
        Mock<IServiceLocator> serviceLocator = new Mock<IServiceLocator>();
        Mock<ISharePointProcessor> processor = new Mock<ISharePointProcessor>();

        public ProcessingThreadManagerTests()
        {
            serviceLocator = new Mock<IServiceLocator>();
            threadManager = new ProcessingThreadManager(serviceLocator.Object);

//            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageThreads"] = "10";
//            ConfigurationManager.AppSettings["EPiServer:Sharepoint:MessageRetrys"] = "10";
        }

        #region Message Thread Tests
        [TestMethod]
        public void Uses_Default_MessageThreads_Value_IfNotSpecified_InConfig()
        {
            int expectedNumberOfThreads = ProcessingThreadManager.DefaultThreadCount;
            int actualNumberOfThreads = 0;
            int testSleepTime = 5000;

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
            while (!addmessagesThread.IsAlive);

            Thread.Sleep(1000);

            Assert.AreEqual(expectedNumberOfThreads,actualNumberOfThreads);                
        }
        #endregion Message Thread Tests
    }
}