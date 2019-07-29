import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from "@microsoft/sp-application-base";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { Dialog } from "@microsoft/sp-dialog";

import ViewCount from "./components/ViewCount";
import { IButtonProps } from "office-ui-fabric-react/lib/Button";
import * as strings from "ViewCountApplicationCustomizerStrings";
import styles from "./AppCustomizer.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";

const LOG_SOURCE: string = "ViewCountApplicationCustomizer";

// id: b54969ee-5581-4ad9-a109-74a79ac4a580

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IViewCountApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ViewCountApplicationCustomizer extends BaseApplicationCustomizer<
  IViewCountApplicationCustomizerProperties
> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = "(No properties were provided.)";
    }

    Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);

    const element: React.ReactElement<IButtonProps> = React.createElement(
      ViewCount
    );

    ReactDOM.render(
      element,
      document.querySelector(".ms-OverflowSet.ms-CommandBar-primaryCommand")
    );

    return Promise.resolve();
    // Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // // Wait for the placeholders to be created (or handle them being changed) and then
    // // render.
    // this.context.placeholderProvider.changedEvent.add(
    //   this,
    //   this._renderPlaceHolders
    // );

    // return Promise.resolve<void>();
  }
}
