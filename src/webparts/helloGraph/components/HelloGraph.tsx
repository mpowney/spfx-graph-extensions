import * as React from 'react';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { DefaultButton, PrimaryButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Logger, LogLevel } from "@pnp/logging";
import { MSGraphClient } from '@microsoft/sp-http';

import styles from './HelloGraph.module.scss';
import { IHelloGraphProps } from './IHelloGraphProps';
import { escape } from '@microsoft/sp-lodash-subset';

const LOG_SOURCE = `HelloGraph.tsx`;

export interface IHelloGraphState {
  isLoading: boolean;
  extensionProvisioned: boolean;
}

export default class HelloGraph extends React.Component<IHelloGraphProps, IHelloGraphState> {

  constructor(props: IHelloGraphProps) {
    
    super(props);

    this.state = {
      isLoading: false,
      extensionProvisioned: false
    }

  }

  private getGraphOpenExtensionValue(value: any = null): any {
    const graphOpenExtensionValue = {
      "@odata.type": "microsoft.graph.openTypeExtension",
      "extensionName": this.props.openExtensionName,
      ...value
    };


  }

  public componentDidMount(): void {

    Logger.log({level: LogLevel.Verbose, message: `${LOG_SOURCE} HelloGraph.tsx calling Graph`})

    this.props.msGraphClientFactory.getClient().then((graphClient: MSGraphClient): void => {

      graphClient.api(`/users/${this.props.userLoginName}?$select=id&$expand=extensions`)

        .get().then(returnValue => {

          Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} HelloGraph.tsx Graph return: `, data: returnValue})

          const foundExtension = returnValue.extensions 
            && returnValue.extensions.filter(extension => { return extension.extensionName == this.props.openExtensionName})
            || [];
          if (foundExtension.length > 0) {
            
            Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} HelloGraph.tsx extension value: `, data: foundExtension[0]});
            this.setState({ extensionProvisioned: true });


          }

        })

    })

  }

  private _provisionButtonClick(): void {

    const graphOpenExtensionValue = {
      "@odata.type": "microsoft.graph.openTypeExtension",
      "extensionName": this.props.openExtensionName,
    };

    this.props.msGraphClientFactory.getClient().then((graphClient: MSGraphClient): void => {

      graphClient.api(`/users/${this.props.userLoginName}/extensions`)
        
        .post(JSON.stringify(graphOpenExtensionValue)).then(value => {

          Logger.log({ level: LogLevel.Info, message: `${LOG_SOURCE} _provisionButtonClick() graph call to provision open extension complete`});

        }).catch(error => {

          Logger.log({ level: LogLevel.Error, message: `${LOG_SOURCE} _provisionButtonClick() Error occurred calling graph to provision open extension:`, data: error});

        });

    });
      
  }

  private _updateValueButtonClick(): void {

    let updateValue = JSON.parse(this.props.openExtensionValue);

    Logger.log({level: LogLevel.Verbose, message: `${LOG_SOURCE} _updateButtonClick() Value to enter: ${JSON.stringify(updateValue)}`});

    this.props.msGraphClientFactory.getClient().then((graphClient: MSGraphClient): void => {

      graphClient.api(`/users/${this.props.userLoginName}/extensions/${this.props.openExtensionName}`)
        
        .patch(JSON.stringify(updateValue)).then(responseValue => {

          Logger.log({ level: LogLevel.Info, message: `${LOG_SOURCE} _updateButtonClick() graph call to update open extension complete`});

        }).catch(error => {

          Logger.log({ level: LogLevel.Error, message: `${LOG_SOURCE} _updateButtonClick() Error occurred calling graph to update open extension:`, data: error});

        });

    });
      


  }



  public render(): React.ReactElement<IHelloGraphProps> {

    return (
      <div className={ styles.helloGraph }>

        { !this.state.extensionProvisioned ?
          <Placeholder
            iconName='Edit'
            iconText='Provision the Open Graph extension'
            description='Please configure the web part.'
            buttonLabel='Provision'
            onConfigure={this._provisionButtonClick.bind(this)} />
        :
          <div className={ styles.container }>
            <div className={ styles.row }>
              <div className={ styles.column }>
                <span className={ styles.title }>Graph Open Extension web part!</span>
                <p className={ styles.subTitle }>Use this web part to update a Graph open extension value.</p>
                <DefaultButton
                  data-automation-id="test"
                  text="Update Value"
                  onClick={this._updateValueButtonClick.bind(this)}
                />
              </div>
            </div>
          </div>
        }
      </div>
    );
  }
}
