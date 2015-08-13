using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Providers;
using EPiServer.ConnectForSharePoint.Rest;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ConnectForSharePoint.SPServices.REST;
using EPiServer.Core;
using EPiServer.Logging;
using EPiServer.PlugIn;
using EPiServer.ServiceLocation;
using EPiServer.Shell.WebForms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.UI.WebControls;

namespace EPiServer.ConnectForSharePoint.Web
{
    /// <summary>
    /// Access this page from Edit Views/My Settings of current user, it provides guide how to register blog account from MS Word.
    /// </summary>
    [GuiPlugIn(Area = PlugInArea.AdminConfigMenu,
        UrlFromModuleFolder = "Web/SharePointConfiguration.aspx",
        LanguagePath = "/sharepoint/admin")]
    public partial class SharePointConfiguration : WebFormsBase
    {
        private const string SharePointUserPassword = "sharepoint-user-password";

        private static readonly ILogger _logger = LogManager.GetLogger();

        private static readonly IContentRepository _contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
        private static readonly ISharePointContentMappingRepository _spContentMappingRepository = ServiceLocator.Current.GetInstance<ISharePointContentMappingRepository>();
        private static readonly ISharePointContentCacheProvider _spContentCacheProvider = ServiceLocator.Current.GetInstance<ISharePointContentCacheProvider>();

        /// <summary>
        /// Called by web page during data bind operation
        /// </summary>
        protected SharePointSettings CurrentSettings
        {
            get { return SharePointSettings.Current; }
        }

        private string ConfiguredUrl
        {
            get { return ServerUrl.Text.Trim(); }
        }

        private string ConfiguredUsername
        {
            get { return Username.Text.Trim(); }
        }

        private string ConfiguredPassword
        {
            get { return (string)Session[SharePointUserPassword]; }
        }

        /// <summary>
        /// HelpUrl of current SharePoint Connector Shell Module
        /// </summary>
        protected string ModuleHelpUrl
        {
            get
            {
                var shellModule = ModuleHelper.GetShellModule();

                return null != shellModule ? shellModule.GetHelpUrl() : string.Empty;
            }
        }

        #region Events

        /// <summary>
        /// Check for Edit Access before displaying this page or generating Registration Script
        /// </summary>
        /// <param name="e"></param>
        protected override void OnInit(EventArgs e)
        {
            // Security validation: user should have Edit access to view this page
            if (!EPiServer.Security.PrincipalInfo.HasAdminAccess)
            {
                throw new AccessDeniedException();
            }

            // Get saved SharePoint's user password from setting storage and then load it to the current session for the first time.
            if (!Page.IsPostBack)
            {
                Session[SharePointUserPassword] = SharePointSettings.Current.Password;
            }

            DataBind();
            base.OnInit(e);
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (!Page.IsPostBack)
            {
                if (PreAuthenticationCheck())
                {
                    string errorMessage;
                    var lists = GetMediaLists(false, out errorMessage);

                    if (!string.IsNullOrEmpty(errorMessage))
                    {
                        ActivateControls(false);
                        ShowMessage(errorMessage, true);
                    }
                    BindListData(lists);
                }

                rbThumbnailOption.Items.Add(
                    new ListItem(Translate("/sharepoint/admin/thumbnailoptions/usegenericthumbnailforall"), ThumbnailOptions.GenericThumbnailForAll.ToString()));
                rbThumbnailOption.Items.Add(
                    new ListItem(Translate("/sharepoint/admin/thumbnailoptions/usesharepointthumbnail"), ThumbnailOptions.OnlySharePointThumbnail.ToString()));
                rbThumbnailOption.Items.Add(
                    new ListItem(Translate("/sharepoint/admin/thumbnailoptions/createthumbnailforall"), ThumbnailOptions.CreateThumbnailForAll.ToString()));

                rbThumbnailOption.SelectedValue = SharePointSettings.Current.ThumbnailOption.ToString();
            }
        }

        /// <summary>
        /// Handles the Click event of the Connect to Server control.
        /// Bind list of SharePoint Lists
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void Connect_Click(object sender, EventArgs e)
        {
            Session[SharePointUserPassword] = Password.Text.Trim();

            if (PreAuthenticationCheck())
            {
                string errorMessage;
                var lists = GetMediaLists(true, out errorMessage);

                if (!string.IsNullOrEmpty(errorMessage))
                {
                    ActivateControls(false);
                    ShowMessage(errorMessage, true);
                    return;
                }

                BindListData(lists);
            }
        }

        /// <summary>
        /// Saves sharepoint connection setting(url, username, password) and root list id to config.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void Save_Click(object sender, EventArgs e)
        {
            MessagePanel.Visible = false;

            #region Validate data

            if (!Page.IsValid)
            {
                return;
            }

            if (!PreAuthenticationCheck())
            {
                return;
            }

            var strErrors = new StringBuilder();
            var entryPointId = 0;
            if (!int.TryParse(EntryPoint.Text.Trim(), out entryPointId))
            {
                strErrors.AppendLine(Translate("/sharepoint/admin/message/invalidentrypoint"));
            }
            else
            {
                try
                {
                    var contentLink = new ContentReference(entryPointId);
                    var contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
                    var content = contentRepository.Get<IContent>(contentLink);

                    if (content == null || content.IsDeleted || !(content is ContentFolder))
                    {
                        strErrors.AppendLine(Translate("/sharepoint/admin/message/entrypointmustbecontentfolder"));
                    }
                    else if (entryPointId != SharePointSettings.Current.EntryPoint.ID) // If the ID is different, check to see if there are any childred.
                    {
                        var children = contentRepository.GetChildren<IContent>(contentLink);
                        if (children.Count<IContent>() > 0 )
                            strErrors.AppendLine(Translate("/sharepoint/admin/message/entrypointmustnothavechildren"));
                    }
                }
                catch (Exception exc)
                {
                    strErrors.AppendLine(Translate("/sharepoint/admin/message/entrypointmustbecontentfolder"));
                    _logger.Error("Error when saving entry point", exc);
                }
            }

            if (strErrors.Length > 0)
            {
                ShowMessage(strErrors.ToString().Replace(Environment.NewLine, "<br/>"), true);
                return;
            }

            #endregion

            #region Save configurations

            // Double-check for SharePoint connection before saving configuration data
            // and compare with the checked lists-views again to make sure the checked lists-views
            // are from the same server with the same authentication info as in Connect button
            // because user can modify configuration data after SharePoint connection has already been tested
            string errorMessage;
            var lists = GetMediaLists(true, out errorMessage);

            if (!string.IsNullOrEmpty(errorMessage))
            {
                ActivateControls(false);
                ShowMessage(errorMessage, true);
                return;
            }

            // Get current selected Lits-View Ids
            var selectedListViewIds = SharePointListCheckBox.Items.Cast<ListItem>().Where(li => li.Selected).Select(li => li.Value);
            if (!IsAnyMatchedId(selectedListViewIds, lists))
            {
                ShowMessage(Translate("/sharepoint/admin/message/listviewnotmatched"), true);
                return;
            }

            try
            {
                // Get the previous Lists-View Ids (the one was saved before this time)
                var oldListIds = SharePointSettings.Current.ListIds;
                var oldEntryPoint = SharePointSettings.Current.EntryPoint;

                // Queues up All folders under the entry point for delete.
                QueueFoldersForDelete(oldListIds, lists);

                // Assign new setting values to SharePointSettings.Current 
                SharePointSettings.Current.SiteUrl = ConfiguredUrl.EndsWith("/") ? ConfiguredUrl : ConfiguredUrl + "/";
                SharePointSettings.Current.Username = ConfiguredUsername;
                SharePointSettings.Current.Password = ConfiguredPassword;
                // Assign current selected Lits-View Ids to the Setting before storing in DDS
                SharePointSettings.Current.ListIds = String.Join(";", selectedListViewIds);
                SharePointSettings.Current.EntryPoint = new ContentReference(entryPointId);

                // TODO Thumbnail option really doesnt make sense anymore since episerver creates the thumbnails.
                SharePointSettings.Current.ThumbnailOption = (ThumbnailOptions)Enum.Parse(typeof(ThumbnailOptions), rbThumbnailOption.SelectedValue);

                // Save settings into DDS first
                SharePointSettings.Current.Save();

                // Put all the add folder messages on the queue
                QueueItemsForAdd(SharePointSettings.Current.ListIds, lists);
            }
            catch (Exception ex)
            {
                SharePointSettings.Current.Reset();
                ShowMessage(string.Format(Translate("/sharepoint/admin/message/saveerror"), ex.Message), true);
                return;
            }

            if (selectedListViewIds.Count() == 0)
            {
                ShowMessage(Translate("/sharepoint/admin/message/savesuccesswithnomedia"), true);
                return;
            }

            ShowMessage(Translate("/sharepoint/admin/message/savesuccess"), false);

            #endregion
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        protected void QueueFoldersForDelete(string listids, IEnumerable<SPList> splists)
        {
            var selectedLists = SharePointHelper.GetSelectedListViewIds(splists);
            foreach (var item in selectedLists)
            {
                // The idListAndView is an object that contains two quids, one for the list and one for the view.
                var idListAndView = SharePointHelper.GetListAndViewIds(item.Key);

                // Get the selected list using the first quid
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                // Get the selected view to from the second quid
                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));

                var FolderName = selectedList.Title + " - " + selectedView.Title;

                // The message on the queue will be SPListItem
                // Because we have only a list (like "documents") and a view (like "all documents") we need to create a SPListItem
                // that contains some aspects of both.
                SPListItem tbd = new SPListItem()
                {
                    Title = selectedList.Title + " - " + selectedView.Title, // the title will be a combo of the list and view names
                    GUID = selectedView.Id,
                    entryPointId = SharePointSettings.Current.EntryPoint.ID,
                    Folder = new SPFolder()
                    {
                        Name = FolderName,
                    },
                };

                _logger.Information("Deleting SharePointFolderAddMessage to the queue " + tbd.Title);
                SharePointMessageEmitterFactory fact2 = new SharePointMessageEmitterFactory();
                fact2.Create().Emit<SharePointFolderDeleteMessage>(new SharePointFolderDeleteMessage()
                {
                    Item = tbd,
                });
            }
        }

        /// <summary>
        /// Adds all items to be added (SPFolders and SPItems) to the queue to be added
        /// </summary>
        /// <param name="listids"></param>
        /// <param name="splists"></param>
        protected void QueueItemsForAdd(string listids, IEnumerable<SPList> splists)
        {
            var selectedLists = SharePointHelper.GetSelectedListViewIds(splists);
            foreach (var item in selectedLists) // For each view selected
            {
                // The idListAndView is an object that contains two quids, one for the list and one for the view.
                var idListAndView = SharePointHelper.GetListAndViewIds(item.Key);

                // Get the selected list using the first quid
                var selectedList = splists.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[0])));

                // Get the selected view to from the second quid
                var selectedView = selectedList.Views.FirstOrDefault(f => f.Id.Equals(new Guid(idListAndView[1])));
                var FolderName = selectedList.Title + " - " + selectedView.Title;
                var parentFolder = selectedList.RootFolder;

                // Add the folder if its not already there
                if (ContentProviderHelper.FindRecursiveFolders(FolderName, SharePointSettings.Current.EntryPoint) == null)
                {
                    SPListItem viewListItem = new SPListItem()
                    {
                        entryPointId = SharePointSettings.Current.EntryPoint.ID,
                        Title = selectedView.Title,
                        Folder = new SPFolder()
                        {
                            Name = FolderName,
                        }

                    };

                    // Add the view folder.
                    SharePointMessageEmitterFactory spfact = new SharePointMessageEmitterFactory();
                    spfact.Create().Emit<SharePointFolderAddMessage>(new SharePointFolderAddMessage()
                    {
                        Item = viewListItem,
                    });
                }

                AddSubItems(parentFolder, selectedList, selectedView, FolderName);
            }
        }
        

        /// <summary>
        /// Adds all the subfolders and subitems for the given parent folder
        /// </summary>
        /// <param name="parentFolder">The parent folder to use as selection criterea items and as a source for the cms folder name</param>
        /// <param name="selectedList">The list of items associated with all the selected views</param>
        /// <param name="selectedView">The selected views in the configuration</param>
        /// <param name="ParentFolderName">An optional parent folder name to use, i.e. in the case of views as a parent folder the name is actually different. </param>
        protected void AddSubItems(SPFolder parentFolder, SPList selectedList, SPListView selectedView, string ParentFolderName = null)
        {
            var items = SharePointHelper.SPServicesHub.SPListsViewsService.GetItems(selectedList, selectedView, parentFolder);
            var iterator = items.GetEnumerator();
            while (iterator.MoveNext())
            {
                var spItem = iterator.Current;
                spItem.entryPointId = -1; // we cant use the entrypointid, we have to use the parent folder name 
                if (string.IsNullOrEmpty(ParentFolderName))
                    spItem.ParentFolderName = parentFolder.Name;
                else
                    spItem.ParentFolderName = ParentFolderName;

                if (spItem.Folder != null)
                {
                    // Only add the folder if its not already there.
                    // Add the folder if its not already there
                    if (ContentProviderHelper.FindRecursiveFolders(spItem.Folder.Name, SharePointSettings.Current.EntryPoint) == null)
                    {
                        _logger.Information("Adding SharePointFolderAddMessage to the queue " + spItem.Title);
                        SharePointMessageEmitterFactory fact2 = new SharePointMessageEmitterFactory();
                        fact2.Create().Emit<SharePointFolderAddMessage>(new SharePointFolderAddMessage()
                        {
                            Item = spItem,
                        });
                    }

                    AddSubItems(spItem.Folder, selectedList, selectedView);
                }
                else if (spItem.File != null)
                {
                    // Before we emit the message and have it passed to the processor, we need to populate the data in the 
                    // spItem.
                    var relativeUrl = SharePointHelper.SPServicesHub.SPFileStreamService.GetFileRelativeUrl(spItem, false);
                    spItem.File.DataStream = SharePointHelper.SPServicesHub.SPFileStreamService.GetFileSteam(relativeUrl, false);

                    // now emit the message
                    _logger.Information("Adding SharePointFolderAddMessage to the queue " + spItem.Title);
                    SharePointMessageEmitterFactory fact2 = new SharePointMessageEmitterFactory();
                    fact2.Create().Emit<SharePointContentAddMessage>(new SharePointContentAddMessage()
                    {
                        Item = spItem,
                    });
                }
            }
        }

        /// <summary>
        /// Refresh sharepoint data.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void Refresh_Click(object sender, EventArgs e)
        {
            string errorMessage;
            var lists = GetMediaLists(true, out errorMessage);

            // Put all the add folder messages on the queue
            QueueItemsForAdd(SharePointSettings.Current.ListIds, lists);

            ShowMessage(Translate("/sharepoint/admin/message/savesuccess"), false);
        }

        /// <summary>
        /// Remove CMS cache for old list-view collection
        /// </summary>
        /// <param name="oldListIds">The old list-view ids</param>
        private static void RemoveCMSCache(string oldListIds)
        {
            // Remove CMS content cache for old list-view collection
            if (!string.IsNullOrEmpty(oldListIds))
            {
                var previousSelectedListIds = oldListIds.Split(new string[] { ";" }, StringSplitOptions.None);

                foreach (var mappingKey in previousSelectedListIds)
                {
                    var contentLink = _spContentMappingRepository.GetContentLink(mappingKey);

                    if (!ContentReference.IsNullOrEmpty(contentLink))
                    {
                        try
                        {
                            // Remove common cache for all children of mapping key
                            var descendents = _contentRepository.GetDescendents(contentLink);
                            if (descendents != null)
                            {
                                foreach (var child in descendents)
                                {
                                    // TECHNOTE:
                                    //  In the "SharePointContentStore" class, we added this cache key for each loaded sharepoint's content.
                                    //  So that, we should remove this cache key here in order to get updated sharepoint's tags for this content.
                                    CacheManager.Remove(DataFactoryCache.PageMasterLanguageCacheKey(child));

                                    CacheManager.Remove(DataFactoryCache.PageCommonCacheKey(child));
                                }
                            }

                        }
                        catch (Exception exc)
                        {
                            _logger.Error("Error occurs when remove common cache for all previously mapped SharePoint content", exc);
                            // NOTE: Try to clear cache, that why do nothing if there are any exception
                        }

                        // TECHNOTE:
                        //  In the "SharePointContentStore" class, we added this cache key for each loaded sharepoint's content.
                        //  So that, we should remove this cache key here in order to get updated sharepoint's tags for this content.
                        CacheManager.Remove(DataFactoryCache.PageMasterLanguageCacheKey(contentLink));

                        // Remove common cache of this mapping key
                        CacheManager.Remove(DataFactoryCache.PageCommonCacheKey(contentLink));
                    }
                }
            }
        }

        #endregion

        #region Private

        /// <summary>
        /// Binds data to SharePoint list selector.
        /// </summary>
        private void BindListData(IEnumerable<SPList> mediaLists)
        {
            if (mediaLists == null)
            {
                return;
            }

            AppendListItemCollection(SharePointListCheckBox, SharePointHelper.GetListViewPairs(mediaLists));
            ActivateControls(true);

            if (Page.IsPostBack)
            {
                ShowMessage(Translate("/sharepoint/admin/message/connectsuccess"), false);
            }
        }

        private IEnumerable<SPList> GetMediaLists(bool isTest, out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                return isTest ? GetMediaListsTest() :
                    _spContentCacheProvider.LatestMediaLists;
            }
            catch (Exception e)
            {
                _logger.Error("Cannot connect to SharePoint.", e);

                var webException = e as WebException;
                var webResponse = webException != null ? (HttpWebResponse)webException.Response : null;
                if (webResponse != null)
                {
                    switch (webResponse.StatusCode)
                    {
                        case HttpStatusCode.Unauthorized:
                            errorMessage = string.Format(Translate("/sharepoint/admin/message/connecttosharepointerror"),
                                                                Translate("/sharepoint/admin/message/unauthorizederror"));
                            return null;
                        case HttpStatusCode.NotFound:
                            errorMessage = string.Format(Translate("/sharepoint/admin/message/connecttosharepointerror"),
                                                                Translate("/sharepoint/admin/message/notfounderror"));
                            return null;
                        default:
                            break;
                    }
                }
                errorMessage = string.Format(Translate("/sharepoint/admin/message/connecttosharepointerror"), e.Message);
                return null;
            }
        }

        private IEnumerable<SPList> GetMediaListsTest()
        {
            var newSettings = new SharePointSettings()
            {
                SiteUrl = ConfiguredUrl,
                Username = ConfiguredUsername,
                Password = ConfiguredPassword
            };

            var contextService = new EndpointContextService();
            contextService.EndpointContext = new SPRestClient(newSettings, SharePointHelper.GetLatestCookies);
            var svcLists = new ListsViewsService();
            svcLists.EndpointContextService = contextService;

            return svcLists.GetLists((int)ListBaseType.DocumentLibrary);
        }

        private bool PreAuthenticationCheck()
        {
            if (string.IsNullOrEmpty(ConfiguredUrl) || string.IsNullOrEmpty(ConfiguredUsername) || string.IsNullOrEmpty(ConfiguredPassword))
            {
                // Deactivate the view selector and button controls
                ActivateControls(false);
                // Show error message in case of posting back and no url/username/password specified
                if (Page.IsPostBack)
                {
                    ShowMessage(Translate("/sharepoint/admin/message/nourlusernamepasswordspecified"), true);
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// Compare the current checked lists-views with given lists
        /// </summary>
        /// <param name="checkedIds">List of checked List-View Id.</param>
        /// <param name="lists">List of SPList for comparing.</param>
        private bool IsAnyMatchedId(IEnumerable<string> checkedIds, IEnumerable<SPList> lists)
        {
            // Dont need check if no checked
            if (checkedIds.Count() < 1)
            {
                return true;
            }

            if (lists.Count() < 1)
            {
                return false;
            }

            var idPairs = SharePointHelper.GetListViewPairs(lists);
            return idPairs.Where(lvId => checkedIds.Contains(lvId.Key)).Count() > 0;
        }

        /// <summary>
        /// Appends list item collection to the given list control
        /// Also, set selection for the matched list item
        /// <param name="listControl">The given list control</param>
        /// </summary>
        private void AppendListItemCollection(ListControl listControl, IEnumerable<KeyValuePair<string, string>> idPairs)
        {
            listControl.Items.Clear();

            var enumerator = idPairs.GetEnumerator();
            var previousSelectedListIds = SharePointSettings.Current.ListIds.Split(new string[] { ";" }, StringSplitOptions.None);

            ListItem newListItem;
            KeyValuePair<string, string> pair;
            while (enumerator.MoveNext())
            {
                pair = enumerator.Current;
                if (default(KeyValuePair<string, string>).Equals(pair))
                {
                    continue;
                }

                newListItem = new ListItem(Server.HtmlEncode(pair.Value), pair.Key, true);
                newListItem.Selected = previousSelectedListIds.Contains(pair.Key);

                listControl.Items.Add(newListItem);
            }
        }

        /// <summary>
        /// Activates or deactivates the list selector and button controls
        /// </summary>
        /// <param name="active">if set to <c>true</c> [active].</param>
        private void ActivateControls(bool active)
        {
            ListSelector.Visible = active;
            Refresh.Enabled = Save.Enabled = active;
        }

        /// <summary>
        /// Shows the message information.
        /// </summary>
        /// <param name="msg">The message.</param>
        /// <param name="isWarning">if set to <c>true</c> message is a warning.</param>
        private void ShowMessage(String msg, bool isWarning)
        {
            ShowMessage(msg, null, isWarning);
        }

        /// <summary>
        /// Shows the message information.
        /// </summary>
        /// <param name="msg">The message.</param>
        /// <param name="tooltip">The tooltip.</param>
        /// <param name="isWarning">if set to <c>true</c> message is a warning.</param>
        private void ShowMessage(String msg, String tooltip, bool isWarning)
        {
            Panel message = new Panel();

            if (!String.IsNullOrEmpty(tooltip))
            {
                MessagePanel.ToolTip = tooltip;
            }

            message.CssClass = isWarning ? "EP-systemMessage EP-systemMessage-Warning" : "EP-systemMessage";
            message.Controls.Add(new Literal { Text = msg });
            MessagePanel.Controls.Add(message);
            MessagePanel.Visible = true;
        }

        /// <summary>
        /// Shows the message connection to SharePoint 365 information.
        /// </summary>
        /// <param name="msg">The message.</param>
        private void ShowConnectionNote(String msg)
        {
            Panel message = new Panel();

            message.CssClass = "EP-systemMessage";
            message.Controls.Add(new Literal { Text = msg });
            ConnectionNote.Controls.Add(message);
            ConnectionNote.Visible = true;
        }

        #endregion
    }
}