import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IPropertyPaneConfiguration, PropertyPaneDropdown, IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";

import styles from './JsDisplayList.module.scss';

import * as strings from 'jsDisplayListStrings';
import { IJsDisplayListWebPartProps } from './IJsDisplayListWebPartProps';

import { Version, Environment, EnvironmentType, Log } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { sp, Item, ItemAddResult, ItemUpdateResult, CommentData } from '@pnp/sp';

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


export default class JsDisplayListWebPart extends BaseClientSideWebPart<IJsDisplayListWebPartProps> {

  public render(): void {
    this.context.statusRenderer.clearError(this.domElement);
    Log.verbose('js-display-List', 'Invoking render');
    this._renderWebPart();
    console.info(this.properties.listTitle);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
  }

  protected onPropertyPaneConfigurationStart(): void {
    if (this._dropdownOptions.length > 0) return;
    this._renderListTitles();
  }

  private _dropdownOptions: IPropertyPaneDropdownOption[] = [];

  private _getListTitles(): Promise<any> {
    return sp.web.lists.filter('Hidden eq false and BaseTemplate eq 100').get();
  }

  private _getListData(listName: string): Promise<any[]> {
    let searchInput: string = (<HTMLInputElement>document.getElementById(`searchInput`)).value;
    let searchField: string = (<HTMLInputElement>document.getElementById(styles.searchFieldSelect)).value;
    if (searchInput.length === 0) {
      //console.info("getAll");
      return sp.web.lists.getByTitle(this.properties.listTitle).items
        .select("Title", "Id", "Created", "Date", "Content", "Attachments", "FieldValuesAsText/Created")
        .expand("FieldValuesAsText").getAll();
    } if (searchField === 'Date') {
      //console.info("getAll " + searchField + " eq datetime'" + searchInput + "T08:00:00Z'");
      return sp.web.lists.getByTitle(this.properties.listTitle).items.filter(searchField + " eq datetime'" + searchInput + "T08:00:00Z'")
      .select("Title", "Id", "Created", "Date", "Content", "Attachments", "FieldValuesAsText/Created")
      .expand("FieldValuesAsText").getAll();
    } else {
      //console.info("getAll startswith(" + searchField + "," + searchInput + ")");
      return sp.web.lists.getByTitle(this.properties.listTitle).items.filter("startswith(" + searchField + ",'" + searchInput + "')")
      .select("Title", "Id", "Created", "Date", "Content", "Attachments", "FieldValuesAsText/Created")
      .expand("FieldValuesAsText").getAll();
    }
  }

  private _getListDocuments(listName: string, id: any): Promise<any> {
    return sp.web.lists.getByTitle(this.properties.listTitle).items.getById(id).attachmentFiles.get();
  }

  private _getListComments(listName: string, id: any): Promise<any> {
    return sp.web.lists.getByTitle(this.properties.listTitle).items.getById(id).comments.orderBy('itemId', true).get();
  }

  private _renderList(items: any[]): void {
    let html: string = '';

    if (!items) {

      html = `<br /><p class="${styles.fontMPlus}">The selected list does not exist.</p>`;

    } else if (items.length === 0) {

      html = `<br /><p class="${styles.fontMPlus}">The selected list is empty</p>`;

    } else {
      items = items.sort((n1, n2) => n2.Id - n1.Id);

      items.forEach((item: any) => {
        let title: string = '';

        if (item.Title === null) {
          title = "Missing title for item with ID= " + item.Id;
        } else {
          title = item.Title;
        }

        let id: any = item.Id;
        let created: any = item.FieldValuesAsText.Created;
        let date: any = item.Date;
        let content: any = item.Content === null ? ' ' : item.Content;
        // let author: any = item.author.title;

        html += `
          <hr />
          <div class="${styles.row}">
                <div class="${styles.column}">
                  ${date.substring(0, 10)}
                </div>
                <div class="${styles.column}">${title}</div>
                <div class="${styles.column}">${created}</div>

          </div>
          <br />
          <div class="${styles.row}">
                <div class="${styles.column}">${content}</div>
          </div>
          <div id="spListDocuments-${id}" class="${styles.row}">
          </div>
          <div id="spListComments-${id}" class="${styles.row}">
          </div>
          <br />`;
      });
    }

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;

    items.forEach((item: any) => {
      let id: any = item["ID"];

      this._getListComments(this.properties.listTitle, id).then((response) => {
        this._renderListComments(response, id);
      }).catch((err) => {
        Log.error('js-display-List', err);
        this.context.statusRenderer.renderError(this.domElement, err);
      });

      if (item["Attachments"] === true) {
        console.info(id + " has attachments");
        this._getListDocuments(this.properties.listTitle, id).then((response) => {
          this._renderListDocuments(response, id);
        }).catch((err) => {
          Log.error('js-display-List', err);
          this.context.statusRenderer.renderError(this.domElement, err);
        });
      }
    });


  }

  private _renderListDocuments(items: any[], id: any): void {
    let html: string = '';

    if (!items) {

      html = `<br /><p class="${styles.fontMPlus}">The selected list does not exist.</p>`;

    } else if (items.length === 0) {

      html = ``;

    } else {
      // items = items.sort((n1,n2)=> parseInt(n1.id) - parseInt(n2.id));
      items.forEach((item: any) => {
        html += `
          <div class="${styles.column}">
            <a id="link" href="${item['ServerRelativeUrl']}" download="">
            ${item['FileName']}
            </a>
          </div>`;
      });
    }
    const listContainer: Element = this.domElement.querySelector(`#spListDocuments-${id}`);
    listContainer.innerHTML = html;
  }

  private _renderListComments(items: any[], id: any): void {
    let html: string = '';

    if (!items) {

      html = `<br /><p class="${styles.fontMPlus}">The selected list does not exist.</p>`;

    } else if (items.length === 0) {

      html = `<br /><p class="${styles.fontMPlus}">No comments</p>`;

    } else {
      items = items.sort((n1, n2) => parseInt(n1.id) - parseInt(n2.id));

      items.forEach((item: any) => {
        html += `
          <div class="${styles.row}" style="padding: 5px; margin: 7px;background: lightgrey;"}>
                <div class="${styles.column}">${item['author'].name}: </div>
                <div class="${styles.column}">${item['text']}</div>
          </div>`;
      });
    }
    html += `
      <div class="${styles.row}"}>
            <div class="${styles.column}"><input id="newCommentInput-${id}" /></div>
            <div class="${styles.column}">
              <button class="${styles.button} create-Button-${id}">
                <span class="">Send comment</span>
              </button>
      </div>`;
    const listContainer: Element = this.domElement.querySelector(`#spListComments-${id}`);
    listContainer.innerHTML = html;

    const webPart: JsDisplayListWebPart = this;
    this.domElement.querySelector(`button.create-Button-${id}`).addEventListener('click', () => { webPart._createComment(id); });
  }

  private _createComment(id: any): void {
    let comment: string = (<HTMLInputElement>document.getElementById(`newCommentInput-` + id)).value;

    sp.web.lists.getByTitle(this.properties.listTitle).items.getById(id).comments.add({
      text: comment
    }).then((result: CommentData): void => {
      // Success
      this.render();
    }, (error: any): void => {
      // Failure
    });
  }

  private _renderWebPart(): void {
    this.domElement.innerHTML = `
        <div class='wrapper ${styles.jsDisplayList}' style='min-height: 500px; max-height: 1000px; overflow-y: scroll; padding-right: 10px;'>
          <div class="${styles.grid} ${styles.jsDisplayList}">
              <select id="${styles.searchFieldSelect}" name = "dropdown">
                <option value = "Date" selected>Date</option>
                <option value = "Title">Title</option>
              </select>
              <input id="searchInput" />
              <button class="${styles.button} searchButton">
                <span class="">Search</span>
              </button>
              <a href="${this.context.pageContext.web.absoluteUrl}/Lists/${this.properties.listTitle}/NewForm.aspx?Source=${escape(document.location.href)}" style="float: right;">
                <button class="${styles.button} newDiscussionButton" >
                  <span class="">New Discussion</span>
                </button>
              </a>
              <div id="spListContainer"></div>
          </div>
        </div>`;

    this._renderListAsync();
    this.domElement.querySelector(`button.searchButton`).addEventListener('click', () => { this._renderListAsync(); });
  }

  private _renderListAsync(): void {
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = "";
    console.info(`Environment.type ${Environment.type}`);
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      let html: string = '<p> Local test environment [No connection to SharePoint]</p>';
      listContainer.innerHTML = html;
    } else {
      this._getListData(this.properties.listTitle).then((response) => {
        this._renderList(response);
      }).catch((err) => {
        Log.error('js-display-List', err);
        this.context.statusRenderer.renderError(this.domElement, err);
      });
    }
  }
  private _renderListTitles(): void {
    this._getListTitles()
      .then((response) => {
        console.info(response);
        this._dropdownOptions = response.map((list: any) => {
          return {
            key: list.Title,
            text: list.Title
          };
        });
        this.context.propertyPane.refresh();
      });
  }
}
