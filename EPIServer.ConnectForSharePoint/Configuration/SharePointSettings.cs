using EPiServer.Core;
using EPiServer.Data;
using EPiServer.Data.Dynamic;
using System.Linq;

namespace EPiServer.ConnectForSharePoint.Configuration
{
    /// <summary>
    /// Settings class to be stored in the DDS
    /// </summary>
    [EPiServerDataStore(AutomaticallyCreateStore = true, AutomaticallyRemapStore = true)]
    public class SharePointSettings : IDynamicData
    {
        #region IDynamicData implementation

        public Identity Id { get; set; }

        #endregion

        /// <summary>
        /// Gets or sets the SharePoint standalone/online URL.
        /// </summary>
        public string SiteUrl { get; set; }

        /// <summary>
        /// Gets or sets the username to connect to SharePoint site.
        /// </summary>
        public string Username { get; set; }

        /// <summary>
        /// Gets or sets the user's password.
        /// </summary>
        public string Password { get; set; }

        /// <summary>
        /// Gets or sets Ids of selected Lists.
        /// </summary>
        public string ListIds { get; set; }

        /// <summary>
        /// Gets or sets the mapping SharePoint's content root assets folder
        /// </summary>
        public ContentReference EntryPoint { get; set; }

        /// <summary>
        /// Gets or sets the selected thumbnail option.
        /// </summary>
        /// <value>
        public ThumbnailOptions ThumbnailOption { get; set; }

        private static SharePointSettings _currentSettings;

        /// <summary>
        /// Gets the current instance of <see cref="EPiServer.ConnectForSharePoint.Configuration.SharePointSettings"/>
        /// </summary>
        public static SharePointSettings Current
        {
            get
            {
                if (_currentSettings == null)
                {
                    var store = DynamicDataStoreFactory.Instance.GetStore(typeof(SharePointSettings));
                    _currentSettings = store.LoadAll<SharePointSettings>().OrderByDescending(x => x.Id.StoreId).FirstOrDefault() ?? new SharePointSettings();
                }

                return _currentSettings;
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SharePointSettings"/> class.
        /// </summary>
        public SharePointSettings()
        {
            SiteUrl = string.Empty;
            Username = string.Empty;
            Password = string.Empty;
            ListIds = string.Empty;
            EntryPoint = ContentReference.EmptyReference;
            ThumbnailOption = ThumbnailOptions.GenericThumbnailForAll;
        }

        /// <summary>
        /// Saves setting values for this <see cref="EPiServer.ConnectForSharePoint.Configuration.SharePointSettings"/> instance.
        /// </summary>
        public void Save()
        {
            var store = DynamicDataStoreFactory.Instance.GetStore(typeof(SharePointSettings));
            store.Save(this);

            _currentSettings = this;
        }

        /// <summary>
        /// Clears setting values that being stored in memory.
        /// </summary>
        public void Reset()
        {
            _currentSettings = null;
        }
   }
}