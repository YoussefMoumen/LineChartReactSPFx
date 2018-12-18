import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'SuiviFinanceWebPartStrings';
import SuiviFinance from './components/SuiviFinance';
import { ISuiviFinanceProps } from './components/ISuiviFinanceProps';

import { PropertyPaneAsyncDropdown } from './PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { update, get } from '@microsoft/sp-lodash-subset';
import pnp from "@pnp/pnpjs";
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls/lib/PropertyFieldMultiSelect';

export interface ISuiviFinanceWebPartProps {
  description: string;
  listName: string;
  item: string;
  multiSelect: string[];
}

export default class SuiviFinanceWebPart extends BaseClientSideWebPart<ISuiviFinanceWebPartProps> {
  private options: IPropertyPaneDropdownOption[] = [];
  public render(): void {
    const element: React.ReactElement<ISuiviFinanceProps > = React.createElement(
      SuiviFinance,
      {
        description: this.properties.description,
        listName: this.properties.listName,
        item: this.properties.item,
        multiSelect: this.properties.multiSelect
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private loadLists(): Promise<IDropdownOption[]> {
    console.log("fetchLists"); 
     var optionsLits: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();   
     return pnp.sp.web.lists.filter(`Hidden eq false and BaseTemplate eq 100`).get().then(lists => {      
      console.log("fetchLists", lists);           
      lists.map((x, key) =>{
        optionsLits.push( { key: x.Title, index: key, text: x.Title });
      });         
      return optionsLits;         
    });
  }

  private onListChange(propertyPath: string, newValue: any): void {
    console.log("onListChange", propertyPath);
    console.log("onListChange", newValue);
    const oldValue: any = get(this.properties, propertyPath);
    console.log("oldValue", oldValue);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    console.log("this.properties", this.properties);
    console.log("propertyPath", propertyPath);
    console.log("newValue", newValue);
    this.loadColumns(newValue);
    this.context.propertyPane.refresh();
    // refresh web part
    this.render();
  }

  public loadColumns(selectedList: string):Promise <IPropertyPaneDropdownOption[]> {
    console.log("Inside loadColumns function");
    // var options : IPropertyPaneDropdownOption[] = new Array<IPropertyPaneDropdownOption>();
     return pnp.sp.web.lists.getByTitle(selectedList).fields.get().then(fields => {
        fields.forEach(field => {
          this.options.push({key: field.StaticName, text: field.Title});
        });             
        return this.options;
      });
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
                // PropertyPaneTextField('description', {
                //   label: strings.DescriptionFieldLabel
                // }),
                new PropertyPaneAsyncDropdown('listName', {
                  label: "List Name",
                  loadOptions: this.loadLists.bind(this),
                  onPropertyChange: this.onListChange.bind(this),
                  selectedKey: this.properties.listName
                }),
                PropertyFieldMultiSelect('multiSelect', {
                  key: 'multiSelect',
                  label: "Column X",
                  options: this.options,
                  selectedKeys: this.properties.multiSelect,
                  disabled:false                  
                }),
                PropertyFieldMultiSelect('multiSelect', {
                  key: 'multiSelect',
                  label: "Column Y",
                  options: this.options,
                  selectedKeys: this.properties.multiSelect,
                  disabled:false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
