import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AadGroupPeopleWebPartStrings';
import AadGroupPeople from './components/AadGroupPeople';
import { IAadGroupPeopleProps } from './components/IAadGroupPeopleProps';
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { update, get } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IAadGroupPeopleWebPartProps {
  groupName: [string, string];
}

export default class AadGroupPeopleWebPart extends BaseClientSideWebPart <IAadGroupPeopleWebPartProps> {

  public render(): void {

    

    this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {

      let groupMembers: Array<any>[] = [];  
      let groupId: string = this.properties.groupName[0];

      client.api("groups/" + groupId + "/members").get((err, res) => {
        if (err) {
          console.error(err);
          return;
        }

        groupMembers = res.value;

        const element: React.ReactElement<IAadGroupPeopleProps> = React.createElement(
          AadGroupPeople,
          {
            groupName: this.properties.groupName,
            members: groupMembers
          }
        );
    
        ReactDom.render(element, this.domElement);

      });

      

    });

    
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                new PropertyPaneAsyncDropdown('groupName', {
                  label: strings.GroupFieldLabel,
                  loadOptions: this.loadGroups.bind(this),
                  onPropertyChange: this.onGroupChange.bind(this),
                  selectedKey: this.properties.groupName[0]
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private loadGroups(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {
      this.context.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {

        let groups: IDropdownOption[] = [];

        client.api("groups").get((err, res) => {
          if(err){
            console.error("Error in Webpart AADGroup-People:");
            console.error(err);
            return;
          }

          res.value.map((item: any) => {
            groups.push({
              key: item.id,
              text: item.displayName
            });
          });
        });

        resolve(groups);
      });
    });
  }

  private onGroupChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    // refresh web part
    this.render();
  }
}
