import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'FormRequestWebPartStrings';

import FormRequest from './components/FormRequest';
import { IFormRequestProps } from './components/IFormRequestProps';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFormRequestWebPartProps {
  description: string;
  context: WebPartContext
}

export default class FormRequestWebPart extends BaseClientSideWebPart<IFormRequestWebPartProps> {

  
  public render(): void {
    const element: React.ReactElement<IFormRequestProps> = React.createElement(
      FormRequest,
      {
        description: this.properties.description,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // protected onInit(): Promise<void> {
  //   return super.onInit().then(_ => {
  //     sp.setup({
  //       spfxContext: this.context
  //     });
  //   });
  // }
  
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
