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
    <style type="text/css">
         .listCheckBox input[type="checkbox"] {
             vertical-align: top;
         }

         .listViewSelector {
             padding: 0 0 8px 2px;
         }
     </style>
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

            // Disable 'Save' button when no list checkboxes are selected
            if ($('span.listCheckBox input[type="checkbox"]:checked').length === 0) {
                $('#FullRegion_Save').attr("disabled", "disabled").parent().attr('class', 'epi-cmsButtondisabled');
            }

            $('span.listCheckBox input[type="checkbox"]').change(function () {
                if ($('span.listCheckBox input[type="checkbox"]:checked').length === 0) {
                    $('#FullRegion_Save').attr("disabled", "disabled").parent().attr('class', 'epi-cmsButtondisabled');
                }
                else {
                    $('#FullRegion_Save').removeAttr("disabled").parent().attr('class', 'epi-cmsButton');
                }
            });
        });

        function processSelectedViews() {
            var selectedViewIds = '';
            var $listChkBox = $('span.listCheckBox input[type="checkbox"]'); // List of all the checkboxes corresponding to lists
            var $viewDdl = $('.viewsDropDown'); // List of all the dropdownlists which contain the views of each list.

            // Loop through each checkbox
            $listChkBox.each(function () {
                if ($(this).is(':checked')) {
                    var viewId = $(this).parent().next($('div.viewSelector')).children($('.viewsDropDown')).val(); // Id ('ListId~ViewId') of the option selected from the dropdownlist.
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

            if (result == "0" && $('span.listCheckBox input[type="checkbox"]:checked').length > 0) {
                $('span.listCheckBox input[type="checkbox"]').removeAttr('disabled'); // List of all the checkboxes corresponding to lists
                $('.listViewSelector select').removeAttr('disabled'); // List of all the dropdownlists which contain the views of each list.
                $('#FullRegion_Save').parent().attr('class', 'epi-cmsButton');
                $('#FullRegion_Save').removeAttr('disabled');

                $('#FullRegion_Refresh').parent().attr('class', 'epi-cmsButton');
                $('#FullRegion_Refresh').removeAttr('disabled');
            }

            // (save or refersh triggered by another user in another browser)
            // items in the queue, but the buttons are enabled 
            if (result != "0" && button.disabled == false) {
                $('span.listCheckBox input[type="checkbox"]').attr('disabled', 'disabled'); // List of all the checkboxes corresponding to lists
                $('.listViewSelector select').attr('disabled', 'disabled'); // List of all the dropdownlists which contain the views of each list.
                $('#FullRegion_Save').parent().attr('class', 'epi-cmsButtondisabled');
                $('#FullRegion_Save').attr('disabled', 'disabled');

                $('#FullRegion_Refresh').parent().attr('class', 'epi-cmsButtondisabled');
                $('#FullRegion_Refresh').attr('disabled', 'disabled');
            }

            lastValue = result;
        }
        function onError(result) {
        }
        document.body.onload = updateQueCounts;
        window.setInterval(updateQueCounts, 1000);
    </script>
    <script type='text/javascript'>
        var disabled = true;
        $(document).ready(function() {
           var dlg = $("#treedialog").dialog({
                autoOpen: false,
                modal: true
            });
            
            $("#btnShow").click(function () {
                dlg.dialog("open");
                $(".ui-dialog-titlebar").hide();

                // disable Ok button until a folder is selected
                document.getElementById('<%# FolderOk.ClientID %>').disabled = disabled;
            });
        });
        function setval(newval) {
            var isNotChecked = document.getElementById('<%# CBCreateNewFolder.ClientID %>').checked;
            var newFolderName = document.getElementById('<%# TBNewFolderName.ClientID %>').value;

            // set foldername text box to show selected folder
            window.PageMethods.GetEntryPointFolderName(newval, isNotChecked, newFolderName, onNodePicked, onError);

            // set hidden text box to be selected folder id so everything continues to work behind the scenes
            document.getElementById('<%# EntryPointCache.ClientID %>').value = newval;
            return false;
        }

        // called when user clicks a node in the tree
        function onNodePicked(result) {
            var epFolderName = document.getElementById('<%= SelectedFolderName.ClientID %>');
            epFolderName.value = result;

            // enable Ok button once a folder is selected
            disabled = false;
            document.getElementById('<%# FolderOk.ClientID %>').disabled = false;
        }


        // called when user clicks Ok in select folder dlg
        function onOkClick() {
            var cachedValue = document.getElementById('<%# EntryPointCache.ClientID %>').value;
            // set hidden text box to be selected folder id so everything continues to work behind the scenes
            document.getElementById('<%# EntryPoint.ClientID %>').value = cachedValue;

            var isNotChecked = document.getElementById('<%# CBCreateNewFolder.ClientID %>').checked;
            var newFolderName = document.getElementById('<%# TBNewFolderName.ClientID %>').value;

            //set foldername text box to show selected folder
            window.PageMethods.GetEntryPointFolderName(cachedValue, isNotChecked, newFolderName, onSuccess, onError);
            $("#treedialog").dialog("close");
        }

        // updates the entry point folder text box on the main config screen
        function onSuccess(result) {
            var epFolderName = document.getElementById('<%= EntryPointFolderName.ClientID %>');
            epFolderName.value = result;
        }

        function onCloseClick() {
            document.getElementById('<%# EntryPointCache.ClientID %>').value = null;
            document.getElementById('<%# CBCreateNewFolder.ClientID %>').checked = false;
            document.getElementById('<%# TBNewFolderName.ClientID %>').value = "";
            document.getElementById('<%# TBNewFolderName.ClientID %>').disabled = true;

            $("#treedialog").dialog("close");
        }

        function onCBClick() {
            var isNotChecked = document.getElementById('<%# CBCreateNewFolder.ClientID %>').checked;

            if (isNotChecked) {
                document.getElementById('<%# TBNewFolderName.ClientID %>').disabled = true;
            } else {
                document.getElementById('<%# TBNewFolderName.ClientID %>').disabled = false;
            }

            return false;
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
            <p class="EP-systemInfo">
                <%= Translate("/sharepoint/admin/selectviewdescription")%>
            </p>
            <asp:Panel ID="ListSelector" runat="server" CssClass="epi-size15" Visible="False" Height="300px" ScrollBars="Auto">
                 <asp:ListView ID="lvLists" runat="server" EnableViewState="True" ItemPlaceholderID="phLists" OnItemDataBound="lvLists_ItemDataBound">
                        <LayoutTemplate>
                            <div>
                                <asp:PlaceHolder runat="server" ID="phLists" />
                            </div>
                        </LayoutTemplate>
                        <ItemTemplate>
                           <div class="listViewSelector">
                            <asp:CheckBox ID="chkList" class="listCheckBox" runat="server"></asp:CheckBox>
                                      <div class="viewSelector">
                            <asp:DropDownList ID="ddViews" class="viewsDropDown" runat="server"></asp:DropDownList>
                                          </div>
                           </div>
                        </ItemTemplate>
                    </asp:ListView>
                    <asp:HiddenField ID="hdnSelectedViews" ClientIDMode="Static" runat="server"></asp:HiddenField>
                <br />
            </asp:Panel>                  
                <br />
                <!-- Modal Tree Dlg -->
                <asp:Label AssociatedControlID="EntryPointFolderName" runat="server"><%# Translate("/sharepoint/admin/selectfolder") %>&nbsp;</asp:Label>
                <asp:TextBox ID="EntryPointFolderName" runat="server" onkeydown="return false" Text="<%# CurrentSettings.EntryPointFolderName %>"></asp:TextBox>
                <input type="button" ID="btnShow" Value="..." />
                <div id="treedialog" class="ui-helper-hidden" >
                    <asp:Label runat="server"><%# Translate("/sharepoint/admin/entrypointinstructions") %></asp:Label>
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
                    <div class="epi-size15">
                        <asp:TextBox ID="SelectedFolderName" style="display:none" runat="server" onkeydown="return false" Text="<%# CurrentSettings.EntryPointFolderName %>"></asp:TextBox>
                        <asp:CheckBox ID="CBCreateNewFolder" onmousedown="onCBClick();" runat="server" Text='<%# Translate("/sharepoint/admin/newsubfolder") %>' checked="<%# !CurrentSettings.UseSelectedFolder %>"/>
                        <asp:TextBox ID="TBNewFolderName" runat="server" Enabled="False"></asp:TextBox>
                        <br />
                    </div>
                    <br />
                    <div align="right">
                        <EPiServerUI:ToolButton ID="FolderOk" runat="server" Text="<%$ Resources: EPiServer, sharepoint.admin.ok %>" OnClientClick="return onOkClick()" CssClass="epi-cmsButton-text epi-cmsButton-tools" />
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
                    <EPiServerUI:ToolButton runat="server" SkinID="Refresh" ID="Refresh" Enabled="False" OnClick="Refresh_Click"
                        Text="<%$ Resources: EPiServer, sharepoint.admin.refreshcache %>" style="vertical-align:top; float: right" />
                </div>
            </fieldset>

        </asp:Panel>
    </div>
</asp:content>
