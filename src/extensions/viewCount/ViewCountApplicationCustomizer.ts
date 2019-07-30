import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer
} from "@microsoft/sp-application-base";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { escape } from "@microsoft/sp-lodash-subset";

import ViewCount from "./components/ViewCount";
import * as strings from "ViewCountApplicationCustomizerStrings";
import View from "./Iview";
import { SPHttpClient } from "@microsoft/sp-http";

const LOG_SOURCE: string = "ViewCountApplicationCustomizer";

export interface IViewCountApplicationCustomizerProperties {
}

export default class ViewCountApplicationCustomizer extends BaseApplicationCustomizer<
  IViewCountApplicationCustomizerProperties
> {
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    console.log(this.context.pageContext.web.absoluteUrl);
    this.loadViews().then(this.createControlButton);
    this.incrementViews();
    return Promise.resolve();
  }

  private loadViews(): Promise<View[]> {
    return new Promise<View[]>(
      (resolve: (views: View[]) => void, reject: (error: any) => void) => {
        this.context.spHttpClient
          .get(
            `https://agarb.sharepoint.com/sites/dev2/_api/web/lists/GetByTitle('ViewCountList')/Items?$select=Title,Views,Id`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "odata-version": ""
              }
            }
          )
          .then(response => response.json())
          .then(response => {
            console.log(
              response.value.map(view => {
                return { page: view.Title, views: view.Views, id: view.Id };
              })
            );
            resolve(
              response.value.map(view => {
                return { page: view.Title, views: view.Views, id: view.Id };
              })
            );
          })
          .catch(error => {
            console.log(error);
            reject(error);
          });
      }
    );
  }

  private incrementViews() {
    this.loadViews().then((views: View[]) => {
      let index = -1;
      views.forEach((view, i) => {
        if (view.page === this.context.pageContext.web.absoluteUrl) index = i;
      });
      console.log(index);
      if (index !== -1) this.updateItem(views[index].id, views[index].views+1);
      else this.createItem();
    });
  }

  private createItem() {
    const body: string = JSON.stringify({
      '__metadata': {
        'type': "SP.Data.ViewCountListListItem"
      },
      'Title': this.context.pageContext.web.absoluteUrl,
      'Views': 1
    });
    return this.context.spHttpClient.post(`https://agarb.sharepoint.com/sites/dev2/_api/web/lists/getbytitle('ViewCountList')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: body
      });
  }

  private updateItem(index:string, views: number): void {
    const body: string = JSON.stringify({
      __metadata: {
        type: "SP.Data.ViewCountListListItem"
      },
      Views: views
    });
    this.context.spHttpClient
      .post(
        `https://agarb.sharepoint.com/sites/dev2/_api/web/lists/getbytitle('ViewCountList')/items(${index})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-type": "application/json;odata=verbose",
            "odata-version": "",
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
          },
          body: body
        }
      )
      .then(response => console.log(response))
      .catch(error => console.error(error));
  }

  private createControlButton() {
    const element = document.createElement("div");
    element.classList.toggle("ms-OverflowSet-item");
    element.classList.toggle("item-279");

    document
      .querySelector(".ms-OverflowSet.ms-CommandBar-primaryCommand")
      .appendChild(element);
    ReactDOM.render(React.createElement(ViewCount, { views: 1 }), element);
  }
}
