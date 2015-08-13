<%@ Page Language="c#" AutoEventWireup="False" CodeBehind="SharePointConfiguration.aspx.cs"
    Inherits="EPiServer.ConnectForSharePoint.Web.SharePointConfiguration, EPiServer.ConnectForSharePoint" Title="EPiServer ConnectForSharePoint Add-on" %>

<%@ Import Namespace="EPiServer" %>
<%@ Register Assembly="EPiServer.UI" Namespace="EPiServer.UI.WebControls" TagPrefix="EPiServerUI" %>

<asp:content contentplaceholderid="HeaderContentRegion" runat="server">
</asp:content>

<asp:content contentplaceholderid="FullRegion" runat="server">
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

            <div class="epi-buttonContainer">
                <EPiServerUI:ToolButton runat="server" SkinID="Save" ID="Save" Enabled="False" OnClick="Save_Click"
                     Text="<%$ Resources: EPiServer, sharepoint.admin.save %>" validationgroup="inputValidationGroup" />
                <EPiServerUI:ToolButton runat="server" SkinID="Refresh" ID="Refresh" Enabled="False" OnClick="Refresh_Click"
                     Text="<%$ Resources: EPiServer, sharepoint.admin.refreshcache %>" />
            </div>
        </asp:Panel>
    </div>
</asp:content>