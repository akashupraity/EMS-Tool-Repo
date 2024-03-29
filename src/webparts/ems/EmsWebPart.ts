import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'EmsWebPartStrings';
import Ems from './components/Ems';
import { IEmsProps } from './components/IEmsProps';
import { sp, Web } from '@pnp/sp/presets/all'


export interface IEmsWebPartProps {
  description: string;
}

export default class EmsWebPart extends BaseClientSideWebPart<IEmsWebPartProps> {
  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
  
      sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    const element: React.ReactElement<IEmsProps> = React.createElement(
      Ems,
      {
        description: this.properties.description,
        siteUrl:this.context.pageContext.web.absoluteUrl,
        context:this.context,
        expenseListTitle:"Expenses",
        expenseDetailListTitle:"ExpenseDetails",
        logHistoryListTitle:"LogHistory",
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
