import * as React from 'react';
import styles from './HelloApi.module.scss';
import { IHelloApiProps } from './IHelloApiProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Logger, LogLevel } from "@pnp/logging";
import { AadHttpClient, AadHttpClientConfiguration } from '@microsoft/sp-http';

const LOG_SOURCE = `HelloApi.tsx`;

export interface IHelloApiState {
  isLoading: boolean;
}

export default class HelloApi extends React.Component<IHelloApiProps, IHelloApiState> {

  constructor(props: IHelloApiProps) {
    
    super(props);

    this.state = {
      isLoading: false,
    };

  }

  public componentDidMount(): void {

    Logger.log({level: LogLevel.Verbose, message: `${LOG_SOURCE} calling endpoint`});

    this.setState({isLoading: true});

    this.props.aadClientFactory.getClient(this.props.clientResource).then((aadClient: AadHttpClient) => {

      return aadClient.get(this.props.apiEndPoint, AadHttpClient.configurations.v1).then(response => {
        return response.json();
      })
      .then(returnValue => {

          this.setState({isLoading: false});
          Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} endpoint return: `, data: returnValue});

        });

    });

  }

  public render(): React.ReactElement<IHelloApiProps> {
    return (
      <div className={ styles.helloApi }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
