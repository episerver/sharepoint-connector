using EPiServer.DataAbstraction;
using EPiServer.Framework.Blobs;
using EPiServer.Framework.Cache;
using EPiServer.ServiceLocation;
using StructureMap;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPiServer.ConnectForSharePoint.Test
{
    public class TestScopeHelper : IDisposable
    {
        public TestScopeHelper(StructureMap.IContainer container)
        {
            SetupServiceLocator(container);
        }

        public TestScopeHelper(Action<ConfigurationExpression> configurator)
            : this(new Container(configurator))
        {
        }

        public TestScopeHelper()
            : this(new Container())
        {
        }


        public IServiceLocator Previous { get; set; }
        public IServiceLocator Current { get; set; }
        public StructureMap.IContainer Container { get; set; }
        private TestScopeHelper SetupServiceLocator(StructureMap.IContainer container)
        {
            Previous = ServiceLocator.Current;
            Current = new StructureMapServiceLocator(container);
            Container = container;
            ServiceLocator.SetLocator(Current);

            return this;
        }

        public TestScopeHelper SetupCommonServices()
        {
            Container.Configure((ce) =>
            {
                ce.For<IContentTypeRepository>().Use<MemoryContentTypeRepository>();
                ce.For<BlobFactory>().Singleton().Use(new BlobFactory("nullprovider", new BlobProvider[] { new NullBlobProvider() }));
            });

            return this;
        }

        public void Dispose()
        {
        }
    }
}
