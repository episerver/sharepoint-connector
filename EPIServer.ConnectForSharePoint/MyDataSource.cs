using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using EPiServer.Core;
using EPiServer.Security;
using EPiServer.UI.WebControls.ContentDataSource;
using EPiServer.ServiceLocation;
using System.Web.UI;
using EPiServer.Web.WebControls;
using EPiServer.Web;

namespace EPiServer.ConnectForSharePoint
{
    public class MyDataSourceView : ContentHierarchicalView
    {
        public MyDataSourceView(IHierarchicalContentDataSource dataSource, string viewPath)
            : base(dataSource, viewPath)
        {

        }

        public override IHierarchicalEnumerable Select()
        {
            throw new NotImplementedException();
        }
    }

    public class MyDataSource : System.Web.UI.Control, 
        IHierarchicalContentDataSource, IContentSource, IDataSource, IDataSourceMethods, IHierarchicalDataSource
    {
        public AccessLevel AccessLevel { get; set; }

        public MyDataSource()
        {
            ServiceLocator.Current.Buildup(this);
        }

        public bool EvaluateHasChildren
        {
            get
            {
                return true;
            }
        }

        public IContent CurrentContent
        {
            get
            {
                throw new NotImplementedException();
            }
        }

        public event EventHandler DataSourceChanged;

        public IList<IContentData> GetFilteredChildren(ContentReference contentLink)
        {
            IList<IContentData> list = new List<IContentData>();
            return list;
        }

        public ContentHierarchicalEnumerable HierarchicalSelect(string viewPath)
        {
            var contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
            ContentReference cf = SiteDefinition.Current.GlobalAssetsRoot;
            if ( !string.IsNullOrEmpty(viewPath) )
            {
                cf = new ContentReference(Int32.Parse(viewPath));
            }
            var folders = contentRepository.GetChildren<ContentFolder>(cf);
            return new ContentHierarchicalEnumerable(folders, this);
        }

        public bool IsRoot(IContentData content)
        {
            return true;
        }

        public T Get<T>(ContentReference contentLink) where T : IContentData
        {
            throw new NotImplementedException();
        }

        public IEnumerable<T> GetChildren<T>(ContentReference contentLink) where T : IContentData
        {
            throw new NotImplementedException();
        }

        public DataSourceView GetView(string viewName)
        {
            throw new NotImplementedException();
        }

        public ICollection GetViewNames()
        {
            throw new NotImplementedException();
        }

        public IEnumerable Select(DataSourceSelectArguments arguments)
        {
            throw new NotImplementedException();
        }

        public int Delete(IDictionary values)
        {
            throw new NotImplementedException();
        }

        public int Insert(IDictionary values)
        {
            throw new NotImplementedException();
        }

        public int Update(IDictionary values)
        {
            throw new NotImplementedException();
        }

        public HierarchicalDataSourceView GetHierarchicalView(string viewPath)
        {
            return new ContentHierarchicalView(this, viewPath);
        }
    }
}