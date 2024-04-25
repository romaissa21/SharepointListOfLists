import { Environment, EnvironmentType, Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListOfListsWebPart.module.scss';
import * as strings from 'ListOfListsWebPartStrings';

import { SPHttpClient,SPHttpClientResponse } from "@microsoft/sp-http";

export interface IListOfListsWebPartProps {
  description: string;
}

export interface ISharePointList {
  Title: string;
  Id: string;
}

export interface ISharePointLists {
  value: ISharePointList[];
}


export default class ListOfListsWebPart extends BaseClientSideWebPart<IListOfListsWebPartProps> {

  private _getListOfLists(): Promise<ISharePointLists>{
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?filter=Hidden eq false`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  private _getAndRenderLists(): void {
    if (Environment.type === EnvironmentType.Local) {

    }
    else if (Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint){
      this._getListOfLists().then((response) => {this._renderListOfLists(response.value)});
    }
  }

  private _renderListOfLists(items: ISharePointList[]): void{
    let html: string = `<h3 class="${ styles.titleH3 }">List of all lists</h3>`;

    items.forEach((Item: ISharePointList) => {
      html += `
      <ul class="${ styles.list }">
        <li class="${ styles.listItem }">
          <span class="ms-font-l">${Item.Title}</span>
        </li>
        <li class="${ styles.listItem }">
          <span class="ms-font-l">${Item.Id}</span>
        </li>
      </ul>`;
    });

    const listsPlaceholder: Element = this.domElement.querySelector('#SPListPlaceholder');
    listsPlaceholder.innerHTML = html;
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.listOfLists }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <h1 class="${ styles.titleH1 }">List</h1>
              <div id="SPListPlaceholder"></div>
            </div>
          </div>
        </div>        
      </div>`;
      this._getAndRenderLists();
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
