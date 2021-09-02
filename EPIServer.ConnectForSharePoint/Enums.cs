namespace EPiServer.ConnectForSharePoint
{
    public enum ListBaseType
    {
        None = -1,
        GenericList = 0,
        DocumentLibrary = 1,
        Unused = 2,
        DiscussionBoard = 3,
        Survey = 4,
        Issue = 5
    }

    public enum SupportedBaseTemplates
    {
        // List of server templates supported by default.
        // 101 - Documents library
        // 109 - Pictures library
        // 851 - Asset Library
        // 119 - Wiki Page Library
        // Full templates list: http://msdn.microsoft.com/en-us/library/microsoft.sharepoint.splisttemplatetype(v=office.15).aspx
        // more info: http://techtrainingnotes.blogspot.com/2008/01/sharepoint-registrationid-list-template.html
        GenericList = 100,
        DocumentLibrary = 101,
        Announcements = 104,    // MAI-176
        Contacts = 105,         // MAI-178
        // Events = 106,           // MAI-177 - buggy, not getting re-occuring schedule, needs to be investigated
        PictureLibrary = 109,   // implicitly supports thumbnail
        AssetLibary = 851,      // implicitly supports thumbnail
        WebPageLibrary = 119
    }

    /// <summary>
    /// Represents options for generating image thumbnail
    /// </summary>
    public enum ThumbnailOptions
    {
        GenericThumbnailForAll,
        OnlySharePointThumbnail,
        CreateThumbnailForAll
    }
}