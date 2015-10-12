<%@ Page Language="c#" AutoEventWireup="False" CodeBehind="SharePointConfiguration.aspx.cs"
    Inherits="EPiServer.ConnectForSharePoint.Web.SharePointConfiguration, EPiServer.ConnectForSharePoint" Title="EPiServer ConnectForSharePoint Add-on" %>

<%@ Import Namespace="EPiServer" %>
<%@ Register Assembly="EPiServer.UI" Namespace="EPiServer.UI.WebControls" TagPrefix="EPiServerUI" %>
<%@ Register TagPrefix="EPiServerUISharePoint" Namespace="EPiServer.ConnectForSharePoint.Web" Assembly="EPiServer.ConnectForSharePoint" %>
<%@ Register TagPrefix="EPiServerUISharePoint" Namespace="EPiServer.ConnectForSharePoint.Models" Assembly="EPiServer.ConnectForSharePoint" %>

<asp:content contentplaceholderid="HeaderContentRegion" runat="server">
</asp:content>

<asp:content contentplaceholderid="FullRegion" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server" EnablePageMethods="true" LoadScriptsBeforeUI="true"/>
	<link href="//ajax.googleapis.com/ajax/libs/jqueryui/1.11.1/themes/smoothness/jquery-ui.min.css" rel="stylesheet">
	<script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
	<script src="//ajax.googleapis.com/ajax/libs/jqueryui/1.11.1/jquery-ui.min.js"></script>
    <script type="text/javascript">
        function updateQueCounts() {
            PageMethods.getCurrentQueCount(onSucess, onError);
        }

        $(document).ready(function () {
            $('#FullRegion_Save').click(function () {
                processSelectedViews();
            });
        });

        function processSelectedViews() {
            var selectedViewIds = '';
            var $listChkBox = $('span.listCheckBox input:[type="checkbox"]'); // List of all the checkboxes corresponding to lists
            var $viewDdl = $('.viewsDropDown'); // List of all the dropdownlists which contain the views of each list.

            // Loop through each checkbox
            $listChkBox.each(function () {
                if ($(this).is(':checked')) {
                    var viewId = $(this).parent().next($('.viewsDropDown')).val(); // Id ('ListId~ViewId') of the option selected from the dropdownlist.
                    if (viewId.indexOf($(this).val()) === 0) { // Additional check to make sure we are choosing the option from the correct List.
                        selectedViewIds += viewId + ',';
                    }
                }
            });

            $('#hdnSelectedViews').val(selectedViewIds);  // Add the List of Ids (comma-separated string) to the hidden variable.
        }

        var lastValue = "0";
        function onSucess(result) {
            var label = document.getElementById('<%=QueueCountTB.ClientID%>');
            var button = document.getElementById('<%=Refresh.ClientID%>');
            label.innerHTML = result;

            // we are done processing, queue is empty - get buttons into proper state
            if (lastValue != "0" && result == "0") {
                window.location.href = window.location.pathname + window.location.search;
            }

            // (save or refersh triggered by another user in another browser)
            // items in the queue, but the buttons are enabled 
            if (result != "0" && button.disabled == false) {
                window.location.href = window.location.pathname + window.location.search;
            }

            // if there is nothing in the queue and the button is disabled
            // odd case where just sometimes we are in the wrong state
            if (result == "0" && button.disabled == true) {
                window.location.href = window.location.pathname + window.location.search;
            }

            lastValue = result;
        }
        function onError(result) {
        }
        document.body.onload = updateQueCounts;
        window.setInterval(updateQueCounts, 1000);
    </script>
    <script type='text/javascript'>
        $(function () {
            $("#treedialog").dialog({
                autoOpen: false,
                modal: true,
            });
            $("#btnShow").click(function () {
                $("#treedialog").dialog("open");
                $(".ui-dialog-titlebar").hide();
            });

        });
        function setval(newval) {
            //set foldername text box to show selected folder
            window.PageMethods.GetEntryPointFolderName(newval, onNodePicked, onError);

            // set hidden text box to be selected folder id so everything continues to work behind the scenes
            document.getElementById('<%# EntryPointCache.ClientID %>').value = newval;
            return false;
        }

        // called when user clicks a node in the tree
        function onNodePicked(result) {
            var epFolderName = document.getElementById('<%= SelectedFolderName.ClientID %>');
            epFolderName.value = result;
        }

        // called when user clicks Ok in select folder dlg
        function onOkClick() {
            // set hidden text box to be selected folder id so everything continues to work behind the scenes
            document.getElementById('<%# EntryPoint.ClientID %>').value = document.getElementById('<%# EntryPointCache.ClientID %>').value;

            //set foldername text box to show selected folder
            window.PageMethods.GetEntryPointFolderName(document.getElementById('<%# EntryPointCache.ClientID %>').value, onSuccess, onError);
            $("#treedialog").dialog("close");
        }

        // updates the entry point folder text box on the main config screen
        function onSuccess(result) {
            var epFolderName = document.getElementById('<%= EntryPointFolderName.ClientID %>');
            epFolderName.value = result;
        }

        function onCloseClick() {
            document.getElementById('<%# EntryPointCache.ClientID %>').value = null;
            $("#treedialog").dialog("close");
        }
    </script>

    <div class="epi-contentContainer epi-padding">
        <div class="epi-contentArea">
            <h1 class="EP-prefix">
                <%= Translate("/sharepoint/admin/displayname")%>
            </h1>
            <asp:Panel ID="ConnectionNote" runat="server" />
            <asp:ValidationSummary runat="server" CssClass="EP-validationSummary" />
            <asp:Panel ID="MessagePanel" runat="server" Visible="false" />
        </div>
               
        <asp:Panel runat="server" class="epi-formArea">
            <fieldset>
                <legend><%= Translate("/sharepoint/admin/description")%></legend>
            <div class="epi-size15">
                <asp:Label AssociatedControlID="ServerUrl" runat="server"><%# Translate("/sharepoint/admin/serverurl") %></asp:Label>
                <asp:TextBox ID="ServerUrl" runat="server" MaxLength="255" Text="<%# CurrentSettings.SiteUrl %>" />
            </div>
            <div class="epi-size15">
                <asp:Label AssociatedControlID="Username" runat="server"><%# Translate("/sharepoint/admin/username") %></asp:Label>
                <asp:TextBox ID="Username" runat="server" MaxLength="255" Text="<%# CurrentSettings.Username %>" />
            </div>

            <div class="epi-size15">
                <asp:Label AssociatedControlID="Password" runat="server"><%# Translate("/sharepoint/admin/password") %></asp:Label>
                <asp:TextBox ID="Password" TextMode="Password" runat="server" MaxLength="255" />
            </div>

            <div class="epi-size15">
                <br />
                <label><%# Translate("/sharepoint/admin/connectiontype") %></label>
                <asp:RadioButton ID="RBOnPremises" GroupName="ConnectionType" Text=<%# Translate("/sharepoint/admin/onpremises") %> runat="server" checked="<%# !CurrentSettings.IsSharePointOnline %>"/>
                <br />
                <label></label>
                <asp:RadioButton ID="RBOnline" GroupName="ConnectionType" Text=<%# Translate("/sharepoint/admin/online") %> runat="server" checked="<%# CurrentSettings.IsSharePointOnline %>"/>
                <br />
            </div>
            <div class="epi-size15">
                <br />
                <label></label>
                <EPiServerUI:ToolButton runat="server" SkinID="ArrowRight" ID="Connect" OnClick="Connect_Click" 
                    Text="<%$ Resources: EPiServer, sharepoint.admin.connect %>" />
            </div>        
            <asp:Panel ID="ListSelector" runat="server" CssClass="epi-size15" Visible="False">
                 <p class="EP-systemInfo">
                    <%= Translate("/sharepoint/admin/selectviewdescription")%>
                </p>
                 <asp:ListView ID="lvLists" runat="server" EnableViewState="True" ItemPlaceholderID="phLists" OnItemDataBound="lvLists_ItemDataBound">
                        <LayoutTemplate>
                            <div>
                                <asp:PlaceHolder runat="server" ID="phLists" />
                            </div>
                        </LayoutTemplate>
                        <ItemTemplate>
                            <div style="padding:2px">
                            <asp:CheckBox ID="chkList" class="listCheckBox" runat="server"></asp:CheckBox>
                            <asp:DropDownList ID="ddViews" class="viewsDropDown" runat="server"></asp:DropDownList>
                           </div>
                        </ItemTemplate>
                    </asp:ListView>
                    <asp:HiddenField ID="hdnSelectedViews" ClientIDMode="Static" runat="server"></asp:HiddenField>
               <%-- <asp:CheckBoxList ID="SharePointListCheckBox" runat="server" />--%>
                <br />
            </asp:Panel>                  
                
                <!-- Modal Tree Dlg -->
                <asp:Label AssociatedControlID="EntryPointFolderName" runat="server"><%# Translate("/sharepoint/admin/selectfolder") %>&nbsp;</asp:Label>
                <asp:TextBox ID="EntryPointFolderName" runat="server" onkeydown="return false" Text="<%# CurrentSettings.EntryPointFolderName %>"></asp:TextBox>
                <input type="button" ID="btnShow" Value="..." />
                <div id="treedialog" >
                    <EPiServerUISharePoint:SharepointContentDataSource ID="contentDataSource" 
                        AccessLevel="NoAccess" 
                        runat="server" 
                        IncludeRootItem="true" 
                        ContentLink="<%# EPiServer.Web.SiteDefinition.Current.SiteAssetsRoot %>" />
                    <div class="episcroll episerver-pagetree-selfcontained" style="max-height:250px;">
                       <EPiServerUI:PageTreeView ID="pageTreeView" 
                           DataSourceID="contentDataSource" 
                           CssClass="episerver-pagetreeview" 
                           runat="server" 
                           ExpandDepth="1"
                           DataTextField="Name" 
                           ExpandOnSelect="false"
                           DataNavigateUrlField="ContentLink" EnableViewState="false">
                            <TreeNodeTemplate>
                                <a href="#" onclick="return setval(<%# ((EPiServer.Core.IContent)((PageTreeNode)Container.DataItem).DataItem).ContentLink.ID %>)" >
                                    <%# Server.HtmlEncode(((PageTreeNode)Container.DataItem).Text) %>
                                </a>
                            </TreeNodeTemplate>
                        </EPiServerUI:PageTreeView>
                    </div> 
                    <div>
                        <asp:Label AssociatedControlID="SelectedFolderName" runat="server"><%# Translate("/sharepoint/admin/selectedfolder") %> </asp:Label>
                        <asp:TextBox ID="SelectedFolderName" runat="server" onkeydown="return false" Text="<%# CurrentSettings.EntryPointFolderName %>"></asp:TextBox>
                    </div>
                    <br />
                    <div align="right">
                        <EPiServerUI:ToolButton runat="server" Text="<%$ Resources: EPiServer, sharepoint.admin.ok %>" OnClientClick="return onOkClick()" CssClass="epi-cmsButton-text epi-cmsButton-tools" />
                        <EPiServerUI:ToolButton runat="server" Text="<%$ Resources: EPiServer, sharepoint.admin.cancel %>" OnClientClick="onCloseClick()" CssClass="epi-cmsButton-text epi-cmsButton-tools"/>
                    </div>
                    <asp:Label AssociatedControlID="EntryPoint" style="display:none" runat="server"><%# Translate("/sharepoint/admin/entrypoint") %></asp:Label>
                    <asp:TextBox ID="EntryPoint" runat="server" style="display:none" MaxLength="10" Text="<%# CurrentSettings.EntryPoint.ID %>" />                
                   
                <asp:RequiredFieldValidator id="EntryPointRequired" runat="server"
                  ControlToValidate="EntryPoint"
                  ErrorMessage='<%# Translate("/sharepoint/admin/entrypointrequired") %>'
                  ForeColor="Red" 
                  validationgroup="inputValidationGroup">
                </asp:RequiredFieldValidator>
                    </div>
                <!-- Modal Tree Dlg -->
                 <asp:TextBox ID="EntryPointCache" runat="server" style="display:none" MaxLength="10" />
            <br/>

                <EPiServerUI:ToolButton runat="server" SkinID="Save" ID="Save" Enabled="False" OnClick="Save_Click"
                    Text="<%$ Resources: EPiServer, sharepoint.admin.save %>" validationgroup="inputValidationGroup" 
                    style="vertical-align:top; float: right"/>
            </fieldset>
            <fieldset>
                <legend><%= Translate("/sharepoint/admin/queuedescription")%></legend>
            <div class="epi-size15">
                <asp:Label AssociatedControlID="QueueCountTB" runat="server">
                    <%= Translate("/sharepoint/admin/remainingitems")%>
                </asp:Label>
                <asp:Label ID="QueueCountTB" runat="server"><%# QueueCount %></asp:Label>
                <EPiServerUI:ToolButton runat="server" SkinID="Refresh" ID="Refresh" ReadOnly="False" OnClick="Refresh_Click"
                    Text="<%$ Resources: EPiServer, sharepoint.admin.refreshcache %>" style="vertical-align:top; float: right" />
            </div>
            </fieldset>

        </asp:Panel>
    </div>
</asp:content>
