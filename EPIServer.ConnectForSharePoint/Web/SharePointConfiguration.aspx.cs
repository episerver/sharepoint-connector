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
        private static bool _createNewEntryPointFolder = false;

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
                var queueStore = new InMemoryQueueStore(AppDomain.CurrentDomain);
                var queue = queueStore.Get("sharepointmod");
                return queue.Count.ToString();
            }
        }

        /// <summary>
        /// Return folder name given an id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        [System.Web.Services.WebMethod]
        public static string GetEntryPointFolderName(string id, bool isChecked, string newFolderName)
        {
            //var numVal = int.Parse(id);
            var selectedFolderRef = _contentRepository.Get<IContent>(new ContentReference(id));
            var fullPath = selectedFolderRef.Name;
            var parent = selectedFolderRef.ParentLink;

            // traverse up the tree to display path to selected folder
            while (parent != null && parent != SiteDefinition.Current.GlobalAssetsRoot)
            {
                var parentRef = _contentRepository.Get<ContentFolder>(new ContentReference(parent.ID));

                fullPath = parentRef.Name + "/" + fullPath;
                parent = parentRef.ParentLink;
            }

            if (isChecked)
            {
                fullPath += "/" + newFolderName;
            }

            _createNewEntryPointFolder = isChecked;

            return fullPath;
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
                    else
                    {
                        BindListData(lists);
                    }

                    if (Convert.ToInt32(QueueCount) > 0)
                    {
                        Connect.Enabled = false;
                        ActivateControls(false);
                    }

                    // check to see if current entry point folder has been deleted, if not return as normal
                    if (ContentReference.EmptyReference != SharePointSettings.Current.EntryPoint)
                    {
                        return;
                    }

                    // if entry point folder has been deleted, pop up error message alerting the user
                    var message = string.Format(Translate("/sharepoint/admin/message/folderdeleted"), SharePointSettings.Current.PreviousEntryPointFolderName);
                    ShowMessage(message.ToString().Replace(Environment.NewLine, "<br/>"), true);
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
                var contentLink = new ContentReference(entryPointId);

                if (!_createNewEntryPointFolder) // if we aren't creating a new folder, then it's business as usual
                {
                    try
                    {
                        var contentRepository = ServiceLocator.Current.GetInstance<IContentRepository>();
                        var content = contentRepository.Get<IContent>(contentLink);

                        if (content == null || content.IsDeleted || !(content is ContentFolder))
                        {
                            strErrors.AppendLine(Translate("/sharepoint/admin/message/entrypointmustbecontentfolder"));
                        }
                        else if (entryPointId != SharePointSettings.Current.EntryPoint.ID)
                            // If the ID is different, check to see if there are any children.
                        {
                            var children = contentRepository.GetChildren<IContent>(contentLink);
                            if (children.Any())
                                strErrors.AppendLine(
                                    Translate("/sharepoint/admin/message/entrypointmustnothavechildren"));
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
                    var folders = EntryPointFolderName.Text.Split('/');
                    var newUserFolder = _contentRepository.GetDefault<ContentFolder>(contentLink);

                    newUserFolder.Name = folders.Last();
                    newUserFolder.ContentLink = ContentReference.EmptyReference;
                    newUserFolder.ParentLink = contentLink;

                    // save new folder to db and get a reference to it so we can the entry point to this new folder
                    var newFolderRef = _contentRepository.Save(newUserFolder, SaveAction.Publish, AccessLevel.FullAccess);
                    entryPointId = newFolderRef.ID;
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
                var messageHelper = new SharePointMessageHelper();
                // if the entry point is different we need to queue all the current folders for delete
                if (entryPointId != SharePointSettings.Current.EntryPoint.ID)
                    messageHelper.QueueRootFoldersForDelete(lists);
                else // else we only delete folders that have been deselected
                    messageHelper.QueueRootFoldersForDelete(lists, selectedListViewIds);

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
            var messageHelper = new SharePointMessageHelper();
            // Put all the add folder messages on the queue
            messageHelper.BeginProcessing(SharePointSettings.Current.ListIds, lists);

            ShowMessage(Translate("/sharepoint/admin/message/refreshsuccess"), false);

            QueueCountTB.Text = QueueCount;
            Connect.Enabled = false;
            ActivateControls(false);
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
                    var documentLists = SharePointHelper.SPServicesHub.SPListsViewsService
                        .GetLists((int)ListBaseType.DocumentLibrary);
                    var customLists = SharePointHelper.SPServicesHub.SPListsViewsService
                        .GetLists((int)ListBaseType.GenericList);
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
            Refresh.Enabled = Save.Enabled = active;

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