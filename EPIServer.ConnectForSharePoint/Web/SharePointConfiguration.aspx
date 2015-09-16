<%@ Page Language="c#" AutoEventWireup="False" CodeBehind="SharePointConfiguration.aspx.cs"
    Inherits="EPiServer.ConnectForSharePoint.Web.SharePointConfiguration, EPiServer.ConnectForSharePoint" Title="EPiServer ConnectForSharePoint Add-on" %>

<%@ Import Namespace="EPiServer" %>
<%@ Register Assembly="EPiServer.UI" Namespace="EPiServer.UI.WebControls" TagPrefix="EPiServerUI" %>

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
            var button = document.getElementById('<%=Save.ClientID%>');
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
                <asp:Label AssociatedControlID="OnlineBox" runat="server"><%# Translate("/sharepoint/admin/online") %></asp:Label>
                <asp:CheckBox ID="OnlineBox" runat="server" checked="<%# CurrentSettings.IsSharePointOnline %>" />
            </div>
            <div class="epi-size15">
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
            </fieldset> 
                
            <fieldset>
            <div class="epi-size15">
                <asp:Label AssociatedControlID="EntryPoint" runat="server"><%# Translate("/sharepoint/admin/entrypoint") %></asp:Label>
                <asp:TextBox ID="EntryPoint" runat="server" MaxLength="10" Text="<%# CurrentSettings.EntryPoint.ID %>" />                
                <asp:RequiredFieldValidator id="EntryPointRequired" runat="server"
                  ControlToValidate="EntryPoint"
                  ErrorMessage='<%# Translate("/sharepoint/admin/entrypointrequired") %>'
                  ForeColor="Red" 
                  validationgroup="inputValidationGroup">
                </asp:RequiredFieldValidator>
            </div>
            </fieldset>

            <fieldset>
            <div class="epi-size15">
                <asp:Label AssociatedControlID="QueueCountTB" runat="server">Number of Items in Queue</asp:Label>
                <asp:Label ID="QueueCountTB" runat="server"><%# QueueCount %></asp:Label>
            </div>
            </fieldset>

            <div class="epi-buttonContainer">
                <EPiServerUI:ToolButton runat="server" SkinID="Save" ID="Save" Enabled="False" OnClick="Save_Click"
                     Text="<%$ Resources: EPiServer, sharepoint.admin.save %>" validationgroup="inputValidationGroup" />
                <EPiServerUI:ToolButton runat="server" SkinID="Refresh" ID="Refresh" ReadOnly="False" OnClick="Refresh_Click"
                     Text="<%$ Resources: EPiServer, sharepoint.admin.refreshcache %>" />
            </div>

        </asp:Panel>
    </div>
</asp:content>