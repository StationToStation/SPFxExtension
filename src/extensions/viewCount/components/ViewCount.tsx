import * as React from "react";
import { CommandBarButton } from "office-ui-fabric-react/lib/Button";
import Popup from "reactjs-popup";

interface IViewCountProps {
  views: number;
}

export default class ViewCount extends React.Component<IViewCountProps, {}> {
  public popupText = () => {
    if (this.props.views > 1)
      return "This page has been visited " + this.props.views + " times.";
    return "This page has never been visited before";
  };

  public render() {
    return (
      <Popup
        contentStyle={{ width: "fit-content" }}
        trigger={
          <CommandBarButton
            className="full-height"
            data-automation-id="views-count"
            iconProps={{ iconName: "View" }}
            text={this.props.views.toString()}
            ariaLabel="Views count"
          />
        }
        position="top center"
        on="hover"
      >
        <div className="card">
          <div className="header">{this.popupText()}</div>
        </div>
      </Popup>
    );
  }
}
