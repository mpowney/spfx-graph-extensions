import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";

import * as strings from 'HelloApiWebPartStrings';
import HelloApi from './components/HelloApi';
import { IHelloApiProps } from './components/IHelloApiProps';

export interface IHelloApiWebPartProps {
  clientResource: string;
  apiEndPoint: string;
}

export default class HelloApiWebPart extends BaseClientSideWebPart<IHelloApiWebPartProps> {

  public render(): void {

    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Verbose;

    const element: React.ReactElement<IHelloApiProps > = React.createElement(
      HelloApi,
      {
        clientResource: this.properties.clientResource,
        apiEndPoint: this.properties.apiEndPoint,
        aadClientFactory: this.context.aadHttpClientFactory
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
                PropertyPaneTextField('clientResource', {
                  label: strings.ClientResourceLabel
                }),
                PropertyPaneTextField('apiEndPoint', {
                  label: strings.ApiEndPointLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
