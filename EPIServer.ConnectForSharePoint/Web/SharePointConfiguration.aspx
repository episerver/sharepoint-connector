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
    <script type="text/javascript">
        function updateQueCounts() {
            PageMethods.getCurrentQueCount(onSucess, onError);
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
            if( result != "0" && button.disabled == false) {
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
                <asp:CheckBoxList ID="SharePointListCheckBox" runat="server" />
            </asp:Panel>                  
                
            <fieldset >
                <legend><%# Translate("/sharepoint/admin/selectfolder") %></legend>
                <div class="epi-size15">
                    <script type='text/javascript'>
                        function setval(newval, foldername) {
                            document.getElementById('<%# EntryPointFolderName.ClientID %>').value = foldername;
                            document.getElementById('<%# EntryPoint.ClientID %>').value = newval;
                            return false;
                        }
                    </script>    
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
                                <a href="#" onclick="return setval(<%# ((EPiServer.Core.IContent)((PageTreeNode)Container.DataItem).DataItem).ContentLink.ID %>, <%# ((EPiServer.Core.IContent)((PageTreeNode)Container.DataItem).DataItem).ContentLink.ID %>)" >
                                    <%# Server.HtmlEncode(((PageTreeNode)Container.DataItem).Text) %>
                                </a>
                            </TreeNodeTemplate>
                        </EPiServerUI:PageTreeView>
                    </div> 
                    <div>
                        <asp:Label AssociatedControlID="EntryPointFolderName" runat="server"><%# Translate("/sharepoint/admin/selectedfolder") %> </asp:Label>
                        <asp:TextBox ID="EntryPointFolderName" runat="server" Text="<%# CurrentSettings.EntryPointFolderName %>"></asp:TextBox>
                    </div>
                    <asp:Label AssociatedControlID="EntryPoint" style="display:none" runat="server"><%# Translate("/sharepoint/admin/entrypoint") %></asp:Label>
                    <asp:TextBox ID="EntryPoint" runat="server" style="display:none" MaxLength="10" Text="<%# CurrentSettings.EntryPoint.ID %>" />                
                <asp:RequiredFieldValidator id="EntryPointRequired" runat="server"
                  ControlToValidate="EntryPoint"
                  ErrorMessage='<%# Translate("/sharepoint/admin/entrypointrequired") %>'
                  ForeColor="Red" 
                  validationgroup="inputValidationGroup">
                </asp:RequiredFieldValidator>
            <br/>

                <EPiServerUI:ToolButton runat="server" SkinID="Save" ID="Save" Enabled="False" OnClick="Save_Click"
                    Text="<%$ Resources: EPiServer, sharepoint.admin.save %>" validationgroup="inputValidationGroup" 
                    style="vertical-align:top; float: right"/>
            </fieldset>

            <fieldset>
                <legend><%= Translate("/sharepoint/admin/queuedescription")%></legend>
            <div class="epi-size15">
                <asp:Label AssociatedControlID="QueueCountTB" runat="server">
                    <%= Translate("/sharepoint/admin/remaingitems")%>
                </asp:Label>
                <asp:Label ID="QueueCountTB" runat="server"><%# QueueCount %></asp:Label>
                <EPiServerUI:ToolButton runat="server" SkinID="Refresh" ID="Refresh" ReadOnly="False" OnClick="Refresh_Click"
                    Text="<%$ Resources: EPiServer, sharepoint.admin.refreshcache %>" style="vertical-align:top; float: right" />
            </div>
            </fieldset>

        </asp:Panel>
    </div>
</asp:content>