import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";

import * as strings from 'HelloCustomApplicationWebPartStrings';
import HelloCustomApplication from './components/HelloCustomApplication';
import { IHelloCustomApplicationProps } from './components/IHelloCustomApplicationProps';

export interface IHelloCustomApplicationWebPartProps {
  clientResource: string;
  apiEndPoint: string;
  clientId: string;
  aadTenant: string;
}

export default class HelloCustomApplicationWebPart extends BaseClientSideWebPart<IHelloCustomApplicationWebPartProps> {

  public render(): void {

    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Verbose;

    const element: React.ReactElement<IHelloCustomApplicationProps > = React.createElement(
      HelloCustomApplication,
      {
        clientResource: this.properties.clientResource,
        apiEndPoint: this.properties.apiEndPoint,
        clientId: this.properties.clientId,
        aadTenant: this.properties.aadTenant,
        httpClient: this.context.httpClient,
        webPartId: this.context.instanceId
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
                }),
                PropertyPaneTextField('clientId', {
                  label: strings.ClientIdLabel
                }),
                PropertyPaneTextField('aadTenant', {
                  label: strings.AadTenantLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
