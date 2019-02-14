import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Logger, LogLevel } from "@pnp/logging";

import * as strings from 'HelloGraphWebPartStrings';
import HelloGraph from './components/HelloGraph';
import { IHelloGraphProps } from './components/IHelloGraphProps';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IHelloGraphWebPartProps {
  openExtensionName: string;
}

const LOG_SOURCE = `HelloGraphWebPart.ts`;

export default class HelloGraphWebPart extends BaseClientSideWebPart<IHelloGraphWebPartProps> {

  public render(): void {

    const extensionProvisioned: boolean = false;

    const element: React.ReactElement<IHelloGraphProps> = !extensionProvisioned ? React.createElement(
        Placeholder,
        {
          iconName: 'Edit',
          iconText: 'Provision the Open Graph extension',
          description: 'Please configure the web part.',
          buttonLabel: 'Provision',
          onConfigure: this._provisionButtonClick.bind(this)
        }
      ) : React.createElement(
        HelloGraph,
        {
          description: this.properties.openExtensionName
        }
      );

    ReactDom.render(element, this.domElement);
  }

  private _provisionButtonClick(): void {

    this.context.msGraphClientFactory.getClient().then((graphClient: MSGraphClient): void => {

      graphClient.api(`/users/${this.context.pageContext.user.loginName}/extensions`)
        
        .post(JSON.stringify({ 
          "@odata.type": "microsoft.graph.openTypeExtension",
          "extensionName": this.properties.openExtensionName
        })).then(value => {

          Logger.log({ level: LogLevel.Info, message: `${LOG_SOURCE} graph call to proviison open extension complete`});

        }).catch(error => {

          Logger.log({ level: LogLevel.Error, message: `${LOG_SOURCE} Error occurred calling graph to proviison open extension: ${error}`});

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
                PropertyPaneTextField('openExtensionName', {
                  label: strings.OpenExtensionNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
