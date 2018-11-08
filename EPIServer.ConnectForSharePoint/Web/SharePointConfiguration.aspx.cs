using EPiServer.ConnectForSharePoint.Configuration;
using EPiServer.ConnectForSharePoint.Helpers;
using EPiServer.ConnectForSharePoint.Messaging.InMemory;
using EPiServer.ConnectForSharePoint.Models;
using EPiServer.ConnectForSharePoint.Rest;
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
using System.Web.UI;
using System.Web.UI.WebControls;
using EPiServer.DataAccess;
using EPiServer.Security;
using EPiServer.Web;
using EPiServer.ConnectForSharePoint.Messaging;
using EPiServer.ConnectForSharePoint.Messaging.Messages;
using EPiServer.DataAbstraction;
using EPiServer.ConnectForSharePoint.SPServices;
using EPiServer.ConnectForSharePoint.Processor;

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
        private static bool _createNewEntryPointFolder = false;
        private static string _newFolderName;
        
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

        [System.Web.Services.WebMethod()]
        /// Used by jscript in aspx to get the current count asyc - 
        /// the jscript needed a method, not a property.
        public static string getCurrentQueCount()
        {
            return QueueCount;
        }

        public static string QueueCount
        {
            get
            {
                string retval = "0";

                var queueStore = new InMemoryQueueStore(AppDomain.CurrentDomain);
                var queue = queueStore.Get("sharepointmod");
                if (queue.Count() > 0)
                {
                    retval = queue.Count.ToString();
                }
                else
                {
                    if( MetaDataProcessor.IsRunning() )
                    {
                        // its possible that the thread is running, but we 
                        // havent added anything to the queue, this allows the 
                        // ui script to set all the buttons disabled.
                        retval = "1"; 
                    }
                }

                return retval;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="id">selected folder id</param>
        /// <param name="isChecked">Tells us to create a new folder or not</param>
        /// <param name="newFolderName">folder name to create if isChecked is True</param>
        /// <returns></returns>
        [System.Web.Services.WebMethod]
        public static string GetEntryPointFolderName(string id, bool isChecked = false, string newFolderName = "")
        {
            //var numVal = int.Parse(id);
            var selectedFolderRef = _contentRepository.Get<IContent>(new ContentReference(id));
            var fullPath = selectedFolderRef.Name;

            // if true, then the root was selected and there is nowhere to traverse as we are at the top already
            if (id != SiteDefinition.Current.GlobalAssetsRoot.ID.ToString())
            {
                var parent = selectedFolderRef.ParentLink;
                bool atRoot = false;
                // traverse up the tree to display path to selected folder
                while (parent != null && !atRoot)
                {
                    var parentRef = _contentRepository.Get<ContentFolder>(new ContentReference(parent.ID));

                    fullPath = parentRef.Name + "/" + fullPath;

                    if (parent == SiteDefinition.Current.GlobalAssetsRoot)
                    {
                        atRoot = true;
                    }

                    parent = parentRef.ParentLink;
                }
            }

            if (isChecked)
            {
                _newFolderName = newFolderName;
                fullPath += "/" + newFolderName;
            }

            _createNewEntryPointFolder = isChecked;

            return fullPath;
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
                    else
                    {
                        BindListData(lists);
                    }

                    if (Convert.ToInt32(QueueCount) > 0)
                    {
                        Connect.Enabled = false;
                        ActivateControls(false);
                    }

                    // show error if destination folder has been deleted in cms
                    ContentFolder f;
                    var repo = ServiceLocator.Current.GetInstance<IContentRepository>();
                    bool loaded = repo.TryGet<ContentFolder>(SharePointSettings.Current.EntryPoint, out f);
                    if (ContentReference.EmptyReference == SharePointSettings.Current.EntryPoint || 
                        loaded == false || f.IsDeleted)
                    {
                        Refresh.Enabled = false;
                        var message = string.Format(Translate("/sharepoint/admin/message/folderdeleted"), SharePointSettings.Current.PreviousEntryPointFolderName);
                        ShowMessage(message.ToString().Replace(Environment.NewLine, "<br/>"), true);
                    }
                }
            }
        }

        protected void lvLists_ItemDataBound(object sender, ListViewItemEventArgs e)
        {
            if (e.Item.ItemType == ListViewItemType.DataItem)
            {
                ListViewDataItem dataItem = (ListViewDataItem)e.Item;
                CheckBox chkList = (CheckBox)e.Item.FindControl("chkList");
                DropDownList ddViews = (DropDownList)e.Item.FindControl("ddViews");
                if (chkList != null && ddViews != null)
                {
                    var listData = (KeyValuePair<string, List<KeyValuePair<string, string>>>)dataItem.DataItem;

                    var previousSelectedListIds = SharePointSettings.Current.ListIds.Split(new string[] { ";" }, StringSplitOptions.None);

                    // Populate Checkboxes (Text = List Title, Value = ListId)
                    string[] listInfo = listData.Key.Split('~');
                    chkList.Text = listInfo[1];
                    chkList.Checked = Array.Exists(previousSelectedListIds, element => element.StartsWith(listInfo[0]));
                    chkList.InputAttributes.Add("Value", listInfo[0]);

                    // Populate Dropdown for corresponding views (Text = View Title, Value = ListId~View Id)
                    var modViews = listData.Value.ConvertAll(x => new KeyValuePair<string, string>(x.Key.Split('~')[1], x.Value));
                    ddViews.DataSource = listData.Value;
                    ddViews.DataTextField = "Value";
                    ddViews.DataValueField = "Key";
                    ddViews.DataBind();
                    foreach (ListItem li in ddViews.Items)
                    {
                        if (Array.Exists(previousSelectedListIds, element => element == li.Value))

                        {
                            // The value of the current ListItem exists in the previously selected Ids.
                            // So make this the default selected value in the dropdownlist.
                            ddViews.SelectedValue = li.Value;
                            break;
                        }
                    }
                }

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
                    Connect.Enabled = true;
                    ShowMessage(errorMessage, true);
                    return;
                }

                BindListData(lists);
            }
            else
            {
                Connect.Enabled = true;
                ListSelectorLabel.Enabled = false;
                ListSelectorLabel.Visible = false;
            }
        }

        /// <summary>
        /// Saves sharepoint connection setting(url, username, password) and root list id to config.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void Save_Click(object sender, EventArgs e)
        {
            // update entry point with the latest cache value that we want to save from the tree dlg selection
            if (!string.IsNullOrEmpty(EntryPointCache.Text))
            {
                EntryPoint = EntryPointCache;
            }

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

            var entryPointId = 0;
            var success = int.TryParse(EntryPoint.Text.Trim(), out entryPointId);
            var strErrors = new StringBuilder();

            if (success)
            {
                var contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
                var contentLink = new ContentReference(entryPointId);

                if (!_createNewEntryPointFolder) // if we aren't creating a new folder, then it's business as usual
                {
                    try
                    {
                        var content = contentRepository.Get<IContent>(contentLink);

                        if (content == null || content.IsDeleted || !(content is ContentFolder))
                        {
                            strErrors.AppendLine(Translate("/sharepoint/admin/message/entrypointmustbecontentfolder"));
                        }
                        else if (entryPointId != SharePointSettings.Current.EntryPoint.ID)
                        {
                            // If the ID is different, check to see if there are any children.
                            var children = contentRepository.GetChildren<IContent>(contentLink);
                                if (children.Any())
                                    strErrors.AppendLine(
                                        Translate("/sharepoint/admin/message/entrypointmustnothavechildren"));

                            // make sure the new folder is not a sub folder of the old parent.
                            var current = new ContentReference(SharePointSettings.Current.EntryPoint.ID);
                            var parents = contentRepository.GetAncestors(contentLink);
                            foreach (var folder in parents)
                            {
                                if (folder.ContentLink == current)
                                {
                                    strErrors.AppendLine(Translate("/sharepoint/admin/message/entrypointcannotbechildofprevious"));
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception exc)
                    {
                        strErrors.AppendLine(Translate("/sharepoint/admin/message/entrypointmustbecontentfolder"));
                        _logger.Error("Error when saving entry point", exc);
                    }
                }
                else // otherwise, we create a new folder, which by default is empty so we can just carry on
                {
                    // make sure we dont keep trying to create it under success or failure conditions.
                    _createNewEntryPointFolder = false; 

                    // make sure the new folder is not a sub folder of the old parent.
                    var current = new ContentReference(SharePointSettings.Current.EntryPoint.ID);
                    var parents = contentRepository.GetAncestors(contentLink);
                    foreach (var folder in parents)
                    {
                        if (folder.ContentLink == current)
                        {
                            strErrors.AppendLine(Translate("/sharepoint/admin/message/entrypointcannotbechildofprevious"));
                            break;
                        }
                    }

                    if (strErrors.Length == 0)
                    {
                        var newUserFolder = _contentRepository.GetDefault<ContentFolder>(contentLink);

                        newUserFolder.Name = _newFolderName;
                        newUserFolder.ContentLink = ContentReference.EmptyReference;
                        newUserFolder.ParentLink = contentLink;

                        // save new folder to db and get a reference to it so we can the entry point to this new folder
                        var newFolderRef = _contentRepository.Save(newUserFolder, SaveAction.Publish, AccessLevel.FullAccess);
                        entryPointId = newFolderRef.ID;
                    }
                }
            }
            else
            {
                strErrors.AppendLine(Translate("/sharepoint/admin/message/invalidentrypoint"));
            }

            if (strErrors.Length > 0)
            {
                // reset text box values to previous valid settings
                EntryPointFolderName.Text = SharePointSettings.Current.EntryPointFolderName;
                EntryPoint.Text = SharePointSettings.Current.EntryPoint.ID.ToString();
                _newFolderName = string.Empty;

                // something happened that caused an error, we need to still bind teh list view data to the page; otherwise,
                // the user needs to refresh manually.
                string errMessage;
                var mediaList = GetMediaLists(true, out errMessage);
                if (!string.IsNullOrEmpty(errMessage))
                {
                    strErrors.AppendLine(errMessage);
                }
                ShowMessage(strErrors.ToString().Replace(Environment.NewLine, "<br/>"), true);

                BindListViewData(SharePointHelper.GetListViewPairsData(mediaList));
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
            List<string> selectedListViewIds = new List<string>();
            selectedListViewIds = this.hdnSelectedViews.Value.TrimEnd(new char[] { ',' }).Split(',').ToList();
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
                var messageHelper = new MetaDataProcessor();

                // if the entry point is different or the url is differnt or 
                // we need to queue all the current folders for delete
                if (entryPointId != SharePointSettings.Current.EntryPoint.ID ||
                    !ConfiguredUrl.Equals(SharePointSettings.Current.SiteUrl))
                {
                    SharePointMessageEmitterFactory emmiter = new SharePointMessageEmitterFactory();

                    // Delete all the folders we created under the current entry point.
                    FolderData fd = new FolderData();
                    var folders = fd.GetFolders();
                    foreach (var folder in folders)
                    {
                        ContentFolder cmsfolder;
                        if (_contentRepository.TryGet<ContentFolder>(new ContentReference(folder.RefId), out cmsfolder))
                        {
                            SPListItem item = new SPListItem()
                            {
                                entryPointId = SharePointSettings.Current.EntryPoint.ID,
                                Folder = new SPFolder() { Name = cmsfolder.Name },
                                GUID = folder.FolderGuid
                            };
                            emmiter.Create().Emit(new SharePointFolderDeleteMessage()
                            {
                                Item = item
                            });
                        }
                        else
                        {   // Failed to load cms folder, so this dynamic data entry is bad.
                            FolderData.Delete(folder);
                        }
                    }
                }
                else // else we only delete folders that have been deselected
                {
                    var utils = new SharepointToCmsUtilities(ServiceLocator.Current);
                    var emFactory = new SharePointMessageEmitterFactory();
                    utils.QueueRootFoldersForDelete(emFactory, lists, selectedListViewIds);
                }

                // Assign new setting values to SharePointSettings.Current and save it to cms
                SharePointSettings.Current.SiteUrl = ConfiguredUrl.EndsWith("/") ? ConfiguredUrl : ConfiguredUrl + "/";
                SharePointSettings.Current.Username = ConfiguredUsername;
                SharePointSettings.Current.Password = ConfiguredPassword;
                SharePointSettings.Current.ListIds = String.Join(";", selectedListViewIds);
                SharePointSettings.Current.EntryPoint = new ContentReference(entryPointId);
                SharePointSettings.Current.EntryPointFolderName = EntryPointFolderName.Text;
                SharePointSettings.Current.IsSharePointOnline = RBOnline.Checked;
                SharePointSettings.Current.Save();

                CacheManager.Remove(SharePointHelper.FormDigestValueCacheKey);

                // Make sure the selected folder has the CmsAdmin full access rights.
                var repo = ServiceLocator.Current.GetInstance<IContentSecurityRepository>();
                var contentSecurity = repo.Get(SharePointSettings.Current.EntryPoint).CreateWritableClone() as IContentSecurityDescriptor;
                if (contentSecurity.IsInherited)
                {
                    contentSecurity.ToLocal();
                }
                contentSecurity.AddEntry(new AccessControlEntry("CmsAdmins", AccessLevel.FullAccess));
                repo.Save(SharePointSettings.Current.EntryPoint, contentSecurity, SecuritySaveType.Replace);

                // Put all the add folder messages on the queue, this will take care of any new folders selected.
                messageHelper.BeginProcessing(SharePointSettings.Current.ListIds, lists);

                QueueCountTB.Text = QueueCount;
                Connect.Enabled = false;
                ActivateControls(false);
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

            // need to reset the folder create flag upon a successful save so we don't continue to make new nested folders
            _createNewEntryPointFolder = false;

            BindListViewData(SharePointHelper.GetListViewPairsData(lists));

            #endregion
        }

        /// <summary>
        /// Refresh sharepoint data.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        protected void Refresh_Click(object sender, EventArgs e)
        {
            string errorMessage;

            // dont test the connection, that would get the data from the control which could have been changed
            var lists = GetMediaLists(false, out errorMessage);
            
            if (lists != null && lists.Any())
            {
            var messageHelper = new MetaDataProcessor();
            // Put all the add folder messages on the queue
            messageHelper.BeginProcessing(SharePointSettings.Current.ListIds, lists);

            ShowMessage(Translate("/sharepoint/admin/message/refreshsuccess"), false);
            QueueCountTB.Text = QueueCount;
            Connect.Enabled = false;
            ActivateControls(false);
                BindListViewData(SharePointHelper.GetListViewPairsData(lists));
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
                ListSelectorLabel.Enabled = false;
                ListSelectorLabel.Visible = false;
                return;
            }
            else
            {
                ListSelectorLabel.Enabled = true;
                ListSelectorLabel.Visible = true;
            }

            ////AppendListItemCollection(SharePointListCheckBox, SharePointHelper.GetListViewPairs(mediaLists));
            BindListViewData(SharePointHelper.GetListViewPairsData(mediaLists));
            ActivateControls(true);

            if (Page.IsPostBack)
            {
                ShowMessage(Translate("/sharepoint/admin/message/connectsuccess"), false);
            }
        }
        private void BindListViewData(List<KeyValuePair<string, List<KeyValuePair<string, string>>>> listViewPairs)
        {
            lvLists.DataSource = listViewPairs;
            lvLists.DataBind();
        }

        private IEnumerable<SPList> GetMediaLists(bool isTest, out string errorMessage)
        {
            errorMessage = string.Empty;
            try
            {
                if (isTest)
                {
                    // just test the connection and return a new list that may or may not match what is in the UI
                    return TestConnection();
                }
                else
                {
                    // Concatinate the document and custom lists.
                    var listViewsService = ServiceLocator.Current.GetInstance<IListsViewsService>();
                    var documentLists = listViewsService.GetLists((int)ListBaseType.DocumentLibrary);
                    var customLists = listViewsService.GetLists((int)ListBaseType.GenericList);
                    var supportedLists = documentLists.Concat(customLists);

                    return supportedLists;
                }
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

        /// <summary>
        /// Uses the configuredurl, username, and password specified in the UI and attempts to connect and get 
        /// the ListBaseType.DocumentLibrary
        /// </summary>
        /// <returns></returns>
        private IEnumerable<SPList> TestConnection()
        {
            var newSettings = new SharePointSettings()
            {
                SiteUrl = ConfiguredUrl,
                Username = ConfiguredUsername,
                Password = ConfiguredPassword,
                IsSharePointOnline = RBOnline.Checked
            };

            var contextService = new EndpointContextService
            {
                EndpointContext = new SPRestClient(newSettings, SharePointHelper.GetLatestCookies)
            };
            var svcLists = new ListsViewsService { EndpointContextService = contextService };

            // Concatinate the document and custom lists.
            var documentLists = svcLists.GetLists((int)ListBaseType.DocumentLibrary);
            var customLists = svcLists.GetLists((int)ListBaseType.GenericList);
            var supportedLists = documentLists.Concat(customLists);

            return supportedLists;
        }

        private bool PreAuthenticationCheck()
        {
            if (string.IsNullOrEmpty(ConfiguredUrl) || string.IsNullOrEmpty(ConfiguredUsername) || string.IsNullOrEmpty(ConfiguredPassword))
            {
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
            ListSelector.Visible = true;
            ListSelector.Enabled = active;
            Save.Enabled = active;
            Refresh.Enabled = !string.IsNullOrEmpty(CurrentSettings.ListIds);

            if (active)
            {
                Connect.Enabled = true;
            }
        }

        /// <summary>
        /// Shows the message information.
        /// </summary>
        /// <param name="msg">The message.</param>
        /// <param name="isWarning">if set to <c>true</c> message is a warning.</param>
        private void ShowMessage(string msg, bool isWarning)
        {
            ShowMessage(msg, null, isWarning);
        }

        /// <summary>
        /// Shows the message information.
        /// </summary>
        /// <param name="msg">The message.</param>
        /// <param name="tooltip">The tooltip.</param>
        /// <param name="isWarning">if set to <c>true</c> message is a warning.</param>
        private void ShowMessage(string msg, string tooltip, bool isWarning)
        {
            Panel message = new Panel();

            if (!string.IsNullOrEmpty(tooltip))
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
        private void ShowConnectionNote(string msg)
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