import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Dropdown, IDropdownOption } from '@fluentui/react';
import { IPropertyPaneDropdownOption } from '@microsoft/sp-property-pane';
import { IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';

export interface IPropertyPaneMultiSelectProps {
  label?: string;
  options: IPropertyPaneDropdownOption[];
  selectedKeys: string[];
  onChanged: (selectedKeys: string[]) => void;
  key?: string;
}

class MultiSelectHost extends React.Component<IPropertyPaneMultiSelectProps, { selectedKeys: string[] }> {
  constructor(props: IPropertyPaneMultiSelectProps) {
    super(props);
    this.state = { selectedKeys: props.selectedKeys || [] };
  }

  public render(): React.ReactElement<IPropertyPaneMultiSelectProps> {
    const options: IDropdownOption[] = this.props.options.map(o => ({ key: o.key as string, text: o.text }));

    return (
      <div style={{ paddingTop: 8 }}>
        <label style={{ display: 'block', marginBottom: 6 }}>{this.props.label}</label>
        <Dropdown
          multiSelect
          options={options}
          selectedKeys={this.state.selectedKeys}
          onChange={this._onChange}
        />
      </div>
    );
  }

  private _onChange = (ev: React.FormEvent<HTMLDivElement>, option?: IDropdownOption, index?: number): void => {
    if (!option) return;
    const selected = (option.selected === true);
    let keys = this.state.selectedKeys.slice();
    const key = option.key as string;
    if (selected) {
      if (keys.indexOf(key) === -1) keys.push(key);
    } else {
      keys = keys.filter(k => k !== key);
    }
    this.setState({ selectedKeys: keys }, () => {
      this.props.onChanged(keys);
    });
  }
}

class PropertyPaneMultiSelectBuilder implements IPropertyPaneField<IPropertyPaneMultiSelectProps> {
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneMultiSelectProps & { onRender: (elem: HTMLElement) => void, onDispose: (elem: HTMLElement) => void, key: string };

  constructor(targetProperty: string, properties: IPropertyPaneMultiSelectProps) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.key || targetProperty,
      ...properties,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this)
    } as any;
  }

  private onRender(elem: HTMLElement): void {
    ReactDom.render(React.createElement(MultiSelectHost, this.properties), elem);
  }

  private onDispose(elem: HTMLElement): void {
    ReactDom.unmountComponentAtNode(elem);
  }
}

export function PropertyPaneMultiSelect(targetProperty: string, properties: IPropertyPaneMultiSelectProps): IPropertyPaneField<IPropertyPaneMultiSelectProps> {
  return new PropertyPaneMultiSelectBuilder(targetProperty, properties) as any;
}

export default PropertyPaneMultiSelect;
