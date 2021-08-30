import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpfxSearchWebPartStrings';
import SpfxSearch from './components/SpfxSearch';
import { ISpfxSearchProps } from './components/ISpfxSearchProps';

export interface ISpfxSearchWebPartProps {
  description: string;
  queryTemplate: string;
}

export default class SpfxSearchWebPart extends BaseClientSideWebPart<ISpfxSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxSearchProps> = React.createElement(
      SpfxSearch,
      {
        
        description: this.properties.description,
        wContext: this.context,
        queryTemplate: this.properties.queryTemplate
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
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('queryTemplate', {
                  label: strings.QueryTemplateFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
