import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { escape } from "@microsoft/sp-lodash-subset";

import ViewCount from "./components/ViewCount";
import * as strings from "ViewCountApplicationCustomizerStrings";
import View from "./Iview";
import { SPHttpClient } from "@microsoft/sp-http";

const LOG_SOURCE: string = "ViewCountApplicationCustomizer";

export interface IViewCountApplicationCustomizerProperties {}
export default class ViewCountApplicationCustomizer extends BaseApplicationCustomizer<
  IViewCountApplicationCustomizerProperties
> {
  private pageURL = window.location.href;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    this.loadViews()
      .then((view: View) => this.incrementViews(view))
      .then((views: number) => this.createControlButton(views, 0));
    return Promise.resolve();
  }

  private loadViews(): Promise<View> {
    return new Promise<View>(
      (resolve: (views: View) => void, reject: (error: any) => void) => {
        this.context.spHttpClient
          .get(
            `https://agarb.sharepoint.com/sites/dev2/_api/web/lists/GetByTitle('ViewCountList')/Items?$select=Views,Id,Title&$filter=Title eq '${escape(
              this.pageURL
            )}'`,
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
            if (response.value[0])
              resolve({
                page: response.value[0].Title,
                views: response.value[0].Views,
                id: response.value[0].Id
              });
            else resolve({ page: "", views: undefined, id: undefined });
          })
          .catch(error => {
            console.error(error);
            reject(error);
          });
      }
    );
  }

  private incrementViews(view: View): number {
    if (view.page === this.pageURL)
      return this.updateItem(view.id, view.views + 1);
    else return this.createItem();
  }

  private createItem(): number {
    const body: string = JSON.stringify({
      __metadata: {
        type: "SP.Data.ViewCountListListItem"
      },
      Title: this.pageURL,
      Views: 1
    });
    this.context.spHttpClient
      .post(
        `https://agarb.sharepoint.com/sites/dev2/_api/web/lists/getbytitle('ViewCountList')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-type": "application/json;odata=verbose",
            "odata-version": ""
          },
          body: body
        }
      )
      .catch(error => console.log(error));
    return 1;
  }

  private updateItem(index: string, views: number): number {
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
      .catch(error => console.error(error));
    return views;
  }

  private createControlButton(views: number, attempt: number) {
    let container = document.querySelector(
      ".ms-OverflowSet.ms-CommandBar-primaryCommand"
    );
    let id: string;
    if (
      (!container || !container.querySelector(".ms-OverflowSet-item")) &&
      attempt < 10
    ) {
      setTimeout(this.createControlButton(views, attempt + 1), 500);
      return;
    }
    if (attempt === 10) {
      console.log(
        `error(garbuz-spfx-extension-client-side-solution): can't find container's .ms-OverflowSet.ms-CommandBar-primaryCommand child`
      );
      id = "-320";
    } else {
      const classes = container.querySelector(".ms-OverflowSet-item").classList;
      Array.prototype.forEach.call(classes, className => {
        if (className.indexOf("item-") !== -1) id = className.slice(4);
      });
    }
    if (!container) {
      console.log(
        `error(garbuz-spfx-extension-client-side-solution): can't find container .ms-OverflowSet.ms-CommandBar-primaryCommand`
      );
      container = document.querySelector(
        "ms-CommandBar"
      );
      if (!container) return;
    }
    const element = document.createElement("div");
    element.classList.toggle("ms-OverflowSet-item");
    element.classList.toggle("item" + id);
    container.appendChild(element);
    ReactDOM.render(React.createElement(ViewCount, { views }), element);
  }
}
