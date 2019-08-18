import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart,
         WebPartContext
} from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-property-pane';
import { setup as pnpSetup } from "@pnp/common";
import * as strings from 'ExpenseClaimsWebPartStrings';
import ExpenseClaims from './components/ExpenseClaims';
import { IExpenseClaimsProps } from './components/IExpenseClaimsProps';


export interface IExpenseClaimsWebPartProps {
  listName: string;
  context: WebPartContext;
}

export default class ExpenseClaimsWebPart extends BaseClientSideWebPart<IExpenseClaimsWebPartProps> {


  public render(): void {
    const element: React.ReactElement<IExpenseClaimsProps > = React.createElement(
      ExpenseClaims,
      {
        listName: this.properties.listName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnpSetup({
        spfxContext: this.context
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
                PropertyPaneTextField('listName', {
                  label: strings.DescriptionListName                
                })
              ]
            }
          ]
        }
      ]
    };
  }


  /**
   * Get the dropdown options for the "List" option in 
   
  private async getListDropdownOptions(){
    const lists = await sp.web.lists.get();
    let dropdownOptions = lists.filter(list => list.Hidden === false && !Constants.listsToIgnore.includes(list.Title)).map(list => {
      return { key: list.Title, text: list.Title };
    });
    dropdownOptions.push({ key: "addnewlist", text: "Add new expenses list" });
    return dropdownOptions;
  }*/

}
