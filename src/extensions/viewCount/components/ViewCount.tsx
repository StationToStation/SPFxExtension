import * as React from "react";
import {
  CommandBarButton,
  IButtonProps
} from "office-ui-fabric-react/lib/Button";
import Popup from "reactjs-popup";

const Card = ({ title }) => (
  <div className="card">
    <div className="header">{title} position </div>
    <div className="content">
      Lorem ipsum dolor sit amet consectetur adipisicing elit. Suscipit autem
      sapiente labore architecto exercitationem optio quod dolor cupiditate
    </div>
  </div>
);

export default class ViewCount extends React.Component<IButtonProps, {}> {
  public render() {
    return (
      <Popup contentStyle={{"width": "fit-content"}}
        trigger={
          <CommandBarButton
            className="full-height"
            data-automation-id="views-count"
            iconProps={{ iconName: "View" }}
            text="Nastia Garbuz"
            ariaLabel="Views count"
            onMouseOver={() => console.log("mouse over")}
          />
        }
        position="top center"
        on="hover"
      >
        <div className="card">
          <div className="header">Hi there</div>
        </div>
      </Popup>
    );
  }
}
