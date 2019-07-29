import * as React from 'react';
import { CommandBarButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';

export default class ViewCount extends React.Component<IButtonProps, {}> {
  public render(): JSX.Element {
    const { disabled, checked } = this.props;

    return (
          <CommandBarButton
            data-automation-id="views-count"
            disabled={disabled}
            checked={checked}
            iconProps={{ iconName: 'View' }}
            text="Nastia Garbuz"
            ariaLabel="Views count"
          />
    );
  }
}