import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version,DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from '@microsoft/sp-webpart-base';

import * as strings from 'RoadmapWebPartStrings';
import Roadmap from './components/Roadmap';
import { IRoadmapProps } from './components/IRoadmapProps';

export interface IRoadmapWebPartProps {
  description: string;
  ListName:string;
  ServiceLine:string;
  SiteUrl:string;
  Context:WebPartContext;
  title: string;
  displayMode:DisplayMode;
  updateProperty:(value: string) => void;
}

export default class RoadmapWebPart extends BaseClientSideWebPart<IRoadmapWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRoadmapProps > = React.createElement(
      Roadmap,
      {
        description: this.properties.description,     
        ServiceLine:this.properties.ServiceLine,
        ListName:this.properties.ListName,
        SiteUrl:this.properties.SiteUrl,
        Context:this.context,
        title: this.properties.title,
    displayMode: this.displayMode,
    updateProperty: (value: string) => {
      this.properties.title = value;
    }
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
                }),
                PropertyPaneTextField('SiteUrl', {
                  label: strings.SiteUrlFieldLabel
                }),
                PropertyPaneTextField('ListName', {
                  label: strings.ListNameFieldLabel
                }),
                PropertyPaneTextField('ServiceLine', {
                  label: strings.ServiceLineFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
