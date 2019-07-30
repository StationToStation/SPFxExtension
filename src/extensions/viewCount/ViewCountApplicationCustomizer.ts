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
import { array } from "prop-types";

const LOG_SOURCE: string = "ViewCountApplicationCustomizer";

export interface IViewCountApplicationCustomizerProperties {}
export default class ViewCountApplicationCustomizer extends BaseApplicationCustomizer<
  IViewCountApplicationCustomizerProperties
> {
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    console.log(window.location.href);
    this.loadViews()
      .then((view: View) => this.incrementViews(view))
      .then((views: number) => this.createControlButton(views));
    return Promise.resolve();
  }

  private loadViews(): Promise<View> {
    return new Promise<View>(
      (resolve: (views: View) => void, reject: (error: any) => void) => {
        this.context.spHttpClient
          .get(
            `https://agarb.sharepoint.com/sites/dev2/_api/web/lists/GetByTitle('ViewCountList')/Items?$select=Views,Id,Title&$filter=Title eq '${escape(
              window.location.href
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
            resolve({
              page: response.value[0].Title,
              views: response.value[0].Views,
              id: response.value[0].Id
            });
          })
          .catch(error => {
            console.log(error);
            reject(error);
          });
      }
    );
  }

  private incrementViews(view: View): number {
    if (view.page === window.location.href)
      return this.updateItem(view.id, view.views + 1);
    else return this.createItem();
  }

  private createItem(): number {
    console.log("create");
    const body: string = JSON.stringify({
      __metadata: {
        type: "SP.Data.ViewCountListListItem"
      },
      Title: window.location.href,
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
    console.log("update");
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

  private createControlButton(views: number) {
    const container = document.querySelector(
      ".ms-OverflowSet.ms-CommandBar-primaryCommand"
    );
    const classes = container.firstElementChild.classList;
    let id: string;
    Array.prototype.forEach.call(classes, className => {
      if (className.indexOf("item-") !== -1) id = className.slice(4);
    });
    const element = document.createElement("div");
    element.classList.toggle("ms-OverflowSet-item");
    element.classList.toggle("item" + id);
    container.appendChild(element);
    ReactDOM.render(React.createElement(ViewCount, { views }), element);
  }
}
