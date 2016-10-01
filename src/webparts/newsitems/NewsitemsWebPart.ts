import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-client-preview';

import * as strings from 'newsitemsStrings';
import Newsitems, { INewsitemsProps } from './components/Newsitems';
import { INewsitemsWebPartProps } from './INewsitemsWebPartProps';

export default class NewsitemsWebPart extends BaseClientSideWebPart<INewsitemsWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    const element: React.ReactElement<INewsitemsProps> = React.createElement(Newsitems, {
      description: this.properties.description,
      numberOfItems: this.properties.numberOfItems,
      httpClient: this.context.httpClient,
      title: this.properties.title,
      siteUrl: this.context.pageContext.web.absoluteUrl,
      listName: this.properties.listName,
      environmentType: this.context.environment.type,
      serviceScope: this.context.serviceScope
    });

    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
                }),
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: strings.LsitNameFieldLabel
                }),
                PropertyPaneSlider('numberOfItems', {
                  label: strings.NumberOfItemsFieldLabel,
                  min: 1,
                  max: 10,
                  step: 1
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
