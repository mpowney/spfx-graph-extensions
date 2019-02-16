import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";

import * as strings from 'HelloGraphWebPartStrings';
import HelloGraph from './components/HelloGraph';
import { IHelloGraphProps } from './components/IHelloGraphProps';

export interface IHelloGraphWebPartProps {
  openExtensionName: string;
  openExtensionValue: string;
}

const LOG_SOURCE = `HelloGraphWebPart.ts`;

export default class HelloGraphWebPart extends BaseClientSideWebPart<IHelloGraphWebPartProps> {

  public render(): void {

    Logger.subscribe(new ConsoleListener());
    Logger.activeLogLevel = LogLevel.Verbose;

    const element: React.ReactElement<IHelloGraphProps> = React.createElement(
        HelloGraph,
        {
          openExtensionName: this.properties.openExtensionName,
          openExtensionValue: this.properties.openExtensionValue,
          userLoginName: this.context.pageContext.user.loginName,
          msGraphClientFactory: this.context.msGraphClientFactory
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
                PropertyPaneTextField('openExtensionName', {
                  label: strings.OpenExtensionNameFieldLabel
                }),
                PropertyPaneTextField('openExtensionValue', {
                  label: strings.OpenExtensionValueFieldLabel,
                  multiline: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
