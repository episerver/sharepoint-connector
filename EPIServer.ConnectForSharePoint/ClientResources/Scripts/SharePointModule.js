define([
// dojo
    'dojo/_base/declare',
    'dojo/_base/lang',
    "dojo/_base/array",

    'dojo/aspect',

    'dojo/dom-attr',
    'dojo/dom-style',

    'dojo/query',
    'dojo/ready',
    'dojo/topic',
    'dojo/when',

// dijit
    'dijit/Destroyable',
    'dijit/registry',
// epi
    'epi/_Module',
    'epi/dependency',
    'epi/routes',

// cms
    'epi-cms/command/EditImage'
],
function (
// dojo
    declare,
    lang,
    array,

    aspect,

    attr,
    style,

    query,
    ready,
    topic,
    when,
// dijit
    Destroyable,
    registry,
// epi
    _Module,
    dependency,
    routes,

// cms
    EditImage
) {

    // module:
    //      epi-connectforsharepoint/SharePointModule
    // tags:
    //      public

    return declare([_Module, Destroyable], {

        // _moduleInitializationInterval [Number]
        //      Interval identification for module initialization
        _moduleInitializationInterval: null,

        // _expired [Boolean]
        //      Flag to indicates should get refresh resources or not
        _expired: false,

        // _expirationTime [Number]
        //      Expired time in minutes that used to indicate when needed to refresh the mapped resources
        _expirationTime: null,

        // _expiredInterval [Number]
        //      Interval identification for setting expired
        _expiredInterval: null,

        // _providerKey [String]
        //      SharePoint content provider name
        _providerKey: null,

        // _thumbnailMode [String]
        //      SharePoint mapped photo thumbnail mode
        _thumbnailMode: null,

        // _cachedTime [String]
        //      Cached time (by get current Date)
        _cachedTime: null,

        // _contentListGridScrollerOriginalPosition [String]
        //      Original position style of the content list's dgrid scroller
        _contentListGridScrollerOriginalPosition: null,

        initialize: function () {
            // summary:
            //      Module initialization
            // tags:
            //      extensions

            this.inherited(arguments);

            this._setupDefaultSettings();

            this._setupUpdateSharePointContentTags();

            ready(lang.hitch(this, function () {
                this._moduleInitializationInterval = setInterval(lang.hitch(this, this._moduleInitialization), 500);
            }));
        },

        uninitialize: function () {
            // summary:
            //      Module uninitialization
            // tags:
            //      extensions

            this.inherited(arguments);

            clearInterval(this._moduleInitializationInterval);
            clearInterval(this._expiredInterval);
        },

        _setupUpdateSharePointContentTags: function () {
            // summary:
            //      Send request to server to initialize tags/categories for current SharePoint content
            //       before each ContextService._loadNewContext method call.
            // tags:
            //      private

            var self = this,
                registry = this.resolveDependency('epi.storeregistry'),
                sharepointContentStore = registry.create('epi.connectforsharepoint.sharepointcontent', this._getRestPath('sharepointcontent')),
                contextService = dependency.resolve('epi.shell.ContextService');

            if (contextService) {
                this.own(
                    aspect.around(contextService, '_loadNewContext', function (origMethod) {
                        return function (contextParameters, callerData, callback) {
                            var originalParams = [contextParameters, callerData, callback];
                            if (contextParameters != null && contextParameters.uri != null && contextParameters.uri.indexOf(self._providerKey) !== -1) {
                                var contentId = contextParameters.uri.match(/\d+/)[0];
                                if (contentId != null) {
                                    return when(sharepointContentStore.query({ contentId: contentId }), function () {
                                        return origMethod.apply(contextService, originalParams);
                                    });
                                }
                            }

                            return origMethod.apply(contextService, originalParams);
                        }
                    })
                );
            }
        },

        _setupDefaultSettings: function () {
            // summary:
            //      Setups default settings
            // tags:
            //      private

            this._providerKey = this._settings.providerKey || 'sharepointcontent';
            this._expirationTime = (this._settings.expirationTime && parseInt(this._settings.expirationTime)) || 5;
            this._thumbnailMode = this._settings.thumbnailMode || 'CreateThumbnailForAll';

            this._expiredInterval = setInterval(lang.hitch(this, function () {
                this._expired = true;
                this._cachedTime = new Date();
            }), this._expirationTime * 60 * 1000);
        },

        _moduleInitialization: function () {
            // summary:
            //      Find content list widget that is child of Media component.
            //      If found, listens on the 'renderRow' event,
            //      and then remove timestamp value from image's 'SRC' attribute in order to makes all thumbnails cachable.
            // tags:
            //      private

            var mediaList = query('.epi-mediaList')[0];
            if (mediaList == null) {
                return;
            }

            var mediaListWidget = registry.byNode(mediaList);
            if (mediaListWidget == null) {
                return;
            }

            var folderTree = mediaListWidget.tree,
                contentList = mediaListWidget.list;
            if (folderTree == null || contentList == null) {
                return;
            }

            clearInterval(this._moduleInitializationInterval);

            var contentListGrid = contentList.grid;
            if (contentListGrid == null || typeof contentListGrid.renderRow !== 'function') {
                return;
            }

            var contentListGridScroller = query('.dgrid-scroller', contentListGrid.domNode)[0];
            this._contentListGridScrollerOriginalPosition = style.get(contentListGridScroller, 'position');

            this.own(
                aspect.after(folderTree, 'onClick', lang.hitch(this, function (item, node, event) {
                    this._decorateContentListGrid(contentListGridScroller, item, node, event);
                }), true),
                aspect.after(contentListGrid, 'renderRow', lang.hitch(this, this._removeRedundantData))
            );

            //===============HACK: find the [Open in Image Editor] menu item and disable it when display context menu for sharepoint media==========

            // on context menu of item
            mediaListWidget.contextMenu.connect(mediaListWidget.contextMenu, "onOpen", lang.hitch(this, function (evt) {
                this._disableOpenImageEditorMenuItem(mediaListWidget.contextMenu);
            }));

            // on context menu of the component chrome's toolbar button
            var componentChrome = dijit.getEnclosingWidget(mediaListWidget.domNode.parentNode);
            var toolbarContextMenuButtons = array.filter(componentChrome.toolbar.getChildren(), function (item) {
                return item._commandCategory == "context";
            });
            var contextMenuButton = (toolbarContextMenuButtons.length > 0) && toolbarContextMenuButtons[0];
            if (!contextMenuButton) {
                return;
            }
            contextMenuButton.dropDown.connect(contextMenuButton.dropDown, "onOpen", lang.hitch(this, function (evt) {
                this._disableOpenImageEditorMenuItem(contextMenuButton.dropDown);
            }));

            //=======================================================================================================================================
        },

        _disableOpenImageEditorMenuItem: function (contextMenu) {
            // summary:
            //      Disable the OpenImageEditor menu item for the content which belongs to SharePoint provider.
            // contextMenu: [Widget]
            //      Context menu that contains the menu item.

            var self = this,
                editImageMenuItems = array.filter(contextMenu.getChildren(), function (item) {
                    return item._command.isInstanceOf(EditImage);
                });

            array.forEach(editImageMenuItems, function (item, i) {
                if (item && item._command.model.currentListItem && item._command.model.currentListItem.providerName == self._providerKey) {
                    item.set("disabled", true);
                }
            });
        },

        _decorateContentListGrid: function (/*DOM*/contentListGridScroller, /*Object*/item, /*Object*/node, /*Event*/event) {
            // summary:
            //      Decorates content list grid's styles
            // contentListGridScroller: [DOM]
            //      The dgrid scroller DOM
            // item: [Object]
            //      The tree node item data object
            // node: [Object]
            //      The tree node object
            // event: [Event]
            //      The tree node event
            // tags:
            //      private

            if (!contentListGridScroller || !item || !item.contentLink) {
                return;
            }

            var isSharePointFolder = item.contentLink.indexOf(this._providerKey) !== -1;

            // TECHNOTE:
            //      Restore original position style of the dgrid scoller to 'absolute' in order to enable on-demand load
            style.set(contentListGridScroller, {
                'position': isSharePointFolder
                    ? 'absolute'
                    : this._contentListGridScrollerOriginalPosition
            });
        },

        _removeRedundantData: function (/*DOM*/data) {
            // summary:
            //      Removes redundant data from the given DOM
            // data: [DOM]
            //      DOM that want to filters redundant data
            // returns: [DOM]
            //      DOM with filtered redundant data
            // tags:
            //      private

            if (this._expired) {
                this._expired = false;

                return data;
            }

            var imgElement = query('.epi-thumbnail', data)[0];
            if (imgElement == null) {
                return data;
            }

            var srcAttrValue = attr.get(imgElement, 'src'),
                searchString = '__' + this._providerKey + '/Thumbnail?epieditmode=False';

            if (srcAttrValue.indexOf(searchString) === -1) {
                return data;
            }

            var newSrcValue = srcAttrValue.substring(0, srcAttrValue.indexOf(searchString) + searchString.length);
            newSrcValue += this._getCachedKey();
            attr.set(imgElement, 'src', newSrcValue);

            return data;
        },

        _getCachedKey: function () {
            // summary:
            //      Gets cached key string following format
            //          &[thumbnail-mode]-[date-hours-minutes-seconds-milliseconds]
            // returns: [String]
            //      The cached key string
            // tags:
            //      private

            return '&' + this._thumbnailMode + '-' + this._getCachedTime();
        },

        _getCachedTime: function () {
            // summary:
            //      Gets cached time string following format
            //          date-hours-minutes-seconds-milliseconds
            // returns: [String]
            // tags:
            //      private

            if (!this._cachedTime) {
                this._cachedTime = new Date();
            }

            return this._cachedTime.getDate() + '-' +
                this._cachedTime.getHours() + '-' +
                this._cachedTime.getMinutes() + '-' +
                this._cachedTime.getSeconds() + '-' +
                this._cachedTime.getMilliseconds();
        },

        _getRestPath: function (name) {
            return routes.getRestPath({ moduleArea: 'episerver.connectforsharepoint', storeName: name });
        }

    });

});