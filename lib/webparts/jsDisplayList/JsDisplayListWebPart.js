var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { PropertyPaneDropdown } from "@microsoft/sp-property-pane";
import styles from './JsDisplayList.module.scss';
import * as strings from 'jsDisplayListStrings';
import { Version, Environment, EnvironmentType, Log } from '@microsoft/sp-core-library';
import { sp } from '@pnp/sp';
//======================
// export interface ISPLists {
//   value: ISPList[];
// }
// export interface ISPList {
//   Title: string;
//   Id: string;
// }
// export interface ISPOption {
//   Id: string;
//   Title: string;
// }
var JsDisplayListWebPart = /** @class */ (function (_super) {
    __extends(JsDisplayListWebPart, _super);
    function JsDisplayListWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._dropdownOptions = [];
        return _this;
    }
    JsDisplayListWebPart.prototype.render = function () {
        this.context.statusRenderer.clearError(this.domElement);
        Log.verbose('js-display-List', 'Invoking render');
        this._renderWebPart();
        console.info(this.properties.listTitle);
    };
    Object.defineProperty(JsDisplayListWebPart.prototype, "disableReactivePropertyChanges", {
        get: function () {
            return true;
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(JsDisplayListWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    JsDisplayListWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        {
                            groupName: strings.DataGroupName,
                            groupFields: [
                                PropertyPaneDropdown('listTitle', {
                                    label: 'List Title',
                                    options: this._dropdownOptions
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    JsDisplayListWebPart.prototype.onPropertyPaneConfigurationStart = function () {
        if (this._dropdownOptions.length > 0)
            return;
        this._renderListTitles();
    };
    JsDisplayListWebPart.prototype._getListTitles = function () {
        return sp.web.lists.filter('Hidden eq false and BaseTemplate eq 100').get();
    };
    JsDisplayListWebPart.prototype._getListData = function (listName) {
        var searchInput = document.getElementById("searchInput").value;
        var searchField = document.getElementById(styles.searchFieldSelect).value;
        if (searchInput.length === 0) {
            //console.info("getAll");
            return sp.web.lists.getByTitle(this.properties.listTitle).items
                .select("Title", "Id", "Created", "Date", "Content", "Attachments", "FieldValuesAsText/Created")
                .expand("FieldValuesAsText").getAll();
        }
        if (searchField === 'Date') {
            //console.info("getAll " + searchField + " eq datetime'" + searchInput + "T08:00:00Z'");
            return sp.web.lists.getByTitle(this.properties.listTitle).items.filter(searchField + " eq datetime'" + searchInput + "T08:00:00Z'")
                .select("Title", "Id", "Created", "Date", "Content", "Attachments", "FieldValuesAsText/Created")
                .expand("FieldValuesAsText").getAll();
        }
        else {
            //console.info("getAll startswith(" + searchField + "," + searchInput + ")");
            return sp.web.lists.getByTitle(this.properties.listTitle).items.filter("startswith(" + searchField + ",'" + searchInput + "')")
                .select("Title", "Id", "Created", "Date", "Content", "Attachments", "FieldValuesAsText/Created")
                .expand("FieldValuesAsText").getAll();
        }
    };
    JsDisplayListWebPart.prototype._getListDocuments = function (listName, id) {
        return sp.web.lists.getByTitle(this.properties.listTitle).items.getById(id).attachmentFiles.get();
    };
    JsDisplayListWebPart.prototype._getListComments = function (listName, id) {
        return sp.web.lists.getByTitle(this.properties.listTitle).items.getById(id).comments.orderBy('itemId', true).get();
    };
    JsDisplayListWebPart.prototype._renderList = function (items) {
        var _this = this;
        var html = '';
        if (!items) {
            html = "<br /><p class=\"" + styles.fontMPlus + "\">The selected list does not exist.</p>";
        }
        else if (items.length === 0) {
            html = "<br /><p class=\"" + styles.fontMPlus + "\">The selected list is empty</p>";
        }
        else {
            items = items.sort(function (n1, n2) { return n2.Id - n1.Id; });
            items.forEach(function (item) {
                var title = '';
                if (item.Title === null) {
                    title = "Missing title for item with ID= " + item.Id;
                }
                else {
                    title = item.Title;
                }
                var id = item.Id;
                var created = item.FieldValuesAsText.Created;
                var date = item.Date;
                var content = item.Content === null ? ' ' : item.Content;
                // let author: any = item.author.title;
                html += "\n          <hr />\n          <div class=\"" + styles.row + "\">\n                <div class=\"" + styles.column + "\">\n                  " + date.substring(0, 10) + "\n                </div>\n                <div class=\"" + styles.column + "\">" + title + "</div>\n                <div class=\"" + styles.column + "\">" + created + "</div>\n\n          </div>\n          <br />\n          <div class=\"" + styles.row + "\">\n                <div class=\"" + styles.column + "\">" + content + "</div>\n          </div>\n          <div id=\"spListDocuments-" + id + "\" class=\"" + styles.row + "\">\n          </div>\n          <div id=\"spListComments-" + id + "\" class=\"" + styles.row + "\">\n          </div>\n          <br />";
            });
        }
        var listContainer = this.domElement.querySelector('#spListContainer');
        listContainer.innerHTML = html;
        items.forEach(function (item) {
            var id = item["ID"];
            _this._getListComments(_this.properties.listTitle, id).then(function (response) {
                _this._renderListComments(response, id);
            }).catch(function (err) {
                Log.error('js-display-List', err);
                _this.context.statusRenderer.renderError(_this.domElement, err);
            });
            if (item["Attachments"] === true) {
                console.info(id + " has attachments");
                _this._getListDocuments(_this.properties.listTitle, id).then(function (response) {
                    _this._renderListDocuments(response, id);
                }).catch(function (err) {
                    Log.error('js-display-List', err);
                    _this.context.statusRenderer.renderError(_this.domElement, err);
                });
            }
        });
    };
    JsDisplayListWebPart.prototype._renderListDocuments = function (items, id) {
        var html = '';
        if (!items) {
            html = "<br /><p class=\"" + styles.fontMPlus + "\">The selected list does not exist.</p>";
        }
        else if (items.length === 0) {
            html = "";
        }
        else {
            // items = items.sort((n1,n2)=> parseInt(n1.id) - parseInt(n2.id));
            items.forEach(function (item) {
                html += "\n          <div class=\"" + styles.column + "\">\n            <a id=\"link\" href=\"" + item['ServerRelativeUrl'] + "\" download=\"\">\n            " + item['FileName'] + "\n            </a>\n          </div>";
            });
        }
        var listContainer = this.domElement.querySelector("#spListDocuments-" + id);
        listContainer.innerHTML = html;
    };
    JsDisplayListWebPart.prototype._renderListComments = function (items, id) {
        var html = '';
        if (!items) {
            html = "<br /><p class=\"" + styles.fontMPlus + "\">The selected list does not exist.</p>";
        }
        else if (items.length === 0) {
            html = "<br /><p class=\"" + styles.fontMPlus + "\">No comments</p>";
        }
        else {
            items = items.sort(function (n1, n2) { return parseInt(n1.id) - parseInt(n2.id); });
            items.forEach(function (item) {
                html += "\n          <div class=\"" + styles.row + "\" style=\"padding: 5px; margin: 7px;background: lightgrey;\"}>\n                <div class=\"" + styles.column + "\">" + item['author'].name + ": </div>\n                <div class=\"" + styles.column + "\">" + item['text'] + "</div>\n          </div>";
            });
        }
        html += "\n      <div class=\"" + styles.row + "\"}>\n            <div class=\"" + styles.column + "\"><input id=\"newCommentInput-" + id + "\" /></div>\n            <div class=\"" + styles.column + "\">\n              <button class=\"" + styles.button + " create-Button-" + id + "\">\n                <span class=\"\">Send comment</span>\n              </button>\n      </div>";
        var listContainer = this.domElement.querySelector("#spListComments-" + id);
        listContainer.innerHTML = html;
        var webPart = this;
        this.domElement.querySelector("button.create-Button-" + id).addEventListener('click', function () { webPart._createComment(id); });
    };
    JsDisplayListWebPart.prototype._createComment = function (id) {
        var _this = this;
        var comment = document.getElementById("newCommentInput-" + id).value;
        sp.web.lists.getByTitle(this.properties.listTitle).items.getById(id).comments.add({
            text: comment
        }).then(function (result) {
            // Success
            _this.render();
        }, function (error) {
            // Failure
        });
    };
    JsDisplayListWebPart.prototype._renderWebPart = function () {
        var _this = this;
        this.domElement.innerHTML = "\n        <div class='wrapper " + styles.jsDisplayList + "' style='min-height: 500px; max-height: 1000px; overflow-y: scroll; padding-right: 10px;'>\n          <div class=\"" + styles.grid + " " + styles.jsDisplayList + "\">\n              <select id=\"" + styles.searchFieldSelect + "\" name = \"dropdown\">\n                <option value = \"Date\" selected>Date</option>\n                <option value = \"Title\">Title</option>\n              </select>\n              <input id=\"searchInput\" />\n              <button class=\"" + styles.button + " searchButton\">\n                <span class=\"\">Search</span>\n              </button>\n              <a href=\"" + this.context.pageContext.web.absoluteUrl + "/Lists/" + this.properties.listTitle + "/NewForm.aspx?Source=" + escape(document.location.href) + "\" style=\"float: right;\">\n                <button class=\"" + styles.button + " newDiscussionButton\" >\n                  <span class=\"\">New Discussion</span>\n                </button>\n              </a>\n              <div id=\"spListContainer\"></div>\n          </div>\n        </div>";
        this._renderListAsync();
        this.domElement.querySelector("button.searchButton").addEventListener('click', function () { _this._renderListAsync(); });
    };
    JsDisplayListWebPart.prototype._renderListAsync = function () {
        var _this = this;
        var listContainer = this.domElement.querySelector('#spListContainer');
        listContainer.innerHTML = "";
        console.info("Environment.type " + Environment.type);
        // Local environment
        if (Environment.type === EnvironmentType.Local) {
            var html = '<p> Local test environment [No connection to SharePoint]</p>';
            listContainer.innerHTML = html;
        }
        else {
            this._getListData(this.properties.listTitle).then(function (response) {
                _this._renderList(response);
            }).catch(function (err) {
                Log.error('js-display-List', err);
                _this.context.statusRenderer.renderError(_this.domElement, err);
            });
        }
    };
    JsDisplayListWebPart.prototype._renderListTitles = function () {
        var _this = this;
        this._getListTitles()
            .then(function (response) {
            console.info(response);
            _this._dropdownOptions = response.map(function (list) {
                return {
                    key: list.Title,
                    text: list.Title
                };
            });
            _this.context.propertyPane.refresh();
        });
    };
    return JsDisplayListWebPart;
}(BaseClientSideWebPart));
export default JsDisplayListWebPart;
//# sourceMappingURL=JsDisplayListWebPart.js.map