import * as React from 'react';
import styles from './HelloCustomApplication.module.scss';
import { IHelloCustomApplicationProps } from './IHelloCustomApplicationProps';
import IAdalConfig from '../../../lib/adal/IAdalConfig';
import { escape } from '@microsoft/sp-lodash-subset';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import '../../../lib/adal/WebPartAuthenticationContext';
import * as adal from 'adal-angular';
import * as AuthenticationContext from 'adal-angular/lib/adal';
import { Logger, LogLevel } from '@pnp/logging';

export interface IHelloCustomApplicationState {
  isLoading: boolean;
  adalError: string;
  adalSignedIn: boolean;
}

const LOG_SOURCE = `HelloCustomApplication.tsx`;

export default class HelloCustomApplication extends React.Component<IHelloCustomApplicationProps, IHelloCustomApplicationState> {

  public adalConfig: IAdalConfig;
  private authCtx: AuthenticationContext;
  
  constructor(props: IHelloCustomApplicationProps) {

    super(props);

    this.state = {
      isLoading: false,
      adalError: "",
      adalSignedIn: false
    };

    if (props.clientId) {

      this.adalConfig = {
        clientId: props.clientId,
        tenant: props.aadTenant,
        extraQueryParameter: 'nux=1',
        endpoints: {
        },
        postLogoutRedirectUri: window.location.origin,
        cacheLocation: 'sessionStorage',
        popUp: true
      };
  
      this.adalConfig.redirectUri = `${window.location.protocol}//${window.location.host}${window.location.pathname}`;
      //this.adalConfig.redirectUri = `https://localhost:4321`; 
      //this.adalConfig.endpoints[props.apiEndPoint] = props.apiEndPoint;
      this.adalConfig.popUp = true;
      this.adalConfig.webPartId = props.webPartId;
      //this.adalConfig.callback = this.adalCallback.bind(this);
      this.adalConfig.callback = (error: any, token: string): void => {
        Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} adalConfig.callback() called with token ${token}`});
      };

      Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} adalConfig: `, data: this.adalConfig});
  
      this.authCtx = new AuthenticationContext(this.adalConfig);
      AuthenticationContext.prototype._singletonInstance = undefined;
  
    }
  }

  private adalCallback = (error: any, token: string): void => {
    Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} adalCallback() called with token ${token}`});

    this.setState((previousState: IHelloCustomApplicationState, currentProps: IHelloCustomApplicationProps): IHelloCustomApplicationState => {
      previousState.adalError = error;
      previousState.adalSignedIn = !(!this.authCtx.getCachedUser());
      return previousState;
    });
  }

  public componentDidMount(): void {
    
    Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} componentDidMount() executing`});

    if (this.authCtx) {

      if ((!this.authCtx.getCachedUser())) {
        this.signIn();
      }
      else {
        this.loadApiData();
      }

      this.authCtx.handleWindowCallback();

      if (window !== window.top) {
        return;
      }

      this.setState((previousState: IHelloCustomApplicationState, props: IHelloCustomApplicationProps): IHelloCustomApplicationState => {
        previousState.adalError = this.authCtx.getLoginError();
        previousState.adalSignedIn = !(!this.authCtx.getCachedUser());
        return previousState;
      });
    }

  }

  public componentDidUpdate(prevProps: IHelloCustomApplicationProps, prevState: IHelloCustomApplicationState, prevContext: any): void {

    Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} componentDidUpdate(): `, data: {'prevState': prevState, 'state': this.state}});

    if (this.state.adalSignedIn && !this.state.isLoading) {
      this.loadApiData();
    }
  }

  public loadApiData() {

    Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} loadApiData() executing`});

    this.setState((previousState: IHelloCustomApplicationState, props: IHelloCustomApplicationProps): IHelloCustomApplicationState => {
      previousState.isLoading = true;
      return previousState;
    });

    Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} loadApiData() getting access token`});
    this.getAccessToken()
      .then((accessToken: string): Promise<any[]> => {

        Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} loadApiData() access token returned ${accessToken}`});
        Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} loadApiData() getting Api data`});

        return HelloCustomApplication.getApiData(accessToken, this.props.httpClient, this.props.apiEndPoint);
      })
      .then((returnedData: any[]): void => {

        Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} API returned data: `, data: returnedData});

      })
      .catch(error => {

        Logger.log({level: LogLevel.Error, message: `${LOG_SOURCE} API returned error: `, data: error});

      });

  }

  public signIn(): void {
    this.authCtx.login();
  }

  private getAccessToken(): Promise<string> {

    Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} getAccessToken executing`});

    return new Promise<string>((resolve: (accessToken: string) => void, reject: (error: any) => void): void => {

      const apiResource: string = this.props.clientResource;
      const accessToken: string = this.authCtx.getCachedToken(apiResource);

      if (accessToken) {

        Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} getAccessToken() accessToken = ${accessToken}`});

        resolve(accessToken);
        return;
      }

      if (this.authCtx.loginInProgress()) {

        Logger.log({level: LogLevel.Warning, message: `${LOG_SOURCE} getAccessToken() login already in progress`});

        reject('Login already in progress');
        return;
      }

      this.authCtx.acquireToken(apiResource, (error: string, token: string) => {

        Logger.log({level: LogLevel.Verbose, message: `${LOG_SOURCE} getAccessToken() acquireToken executed with apiResource ${apiResource}`});

        if (error) {
          Logger.log({level: LogLevel.Error, message: `${LOG_SOURCE} getAccessToken() acquireToken with apiResource ${apiResource}, returned with error`, data: error});

          reject(error);
          return;
        }

        if (token) {
          Logger.log({level: LogLevel.Info, message: `${LOG_SOURCE} getAccessToken() acquireToken access token returned ${token}`});
          resolve(token);
        }
        else {
          Logger.log({level: LogLevel.Error, message: `${LOG_SOURCE} getAccessToken() acquireToken couldn't retrieve access token`});
          reject('Couldn\'t retrieve access token');
        }
      });
    });
  }

  private static getApiData(accessToken: string, httpClient: HttpClient, apiCall: string): Promise<any[]> {
    return new Promise<any[]>((resolve: (upcomingMeetings: any[]) => void, reject: (error: any) => void): void => {

      httpClient.get(apiCall, HttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata.metadata=none',
          'Authorization': 'Bearer ' + accessToken
        }
      })
        .then((response: HttpClientResponse): Promise<{ value: any[] }> => {
          return response.json();
        })
        .then((data: { value: any[] }): void => {
          const returnData: any[] = [];

          for (let i: number = 0; i < data.value.length; i++) {
            returnData.push(data.value[i]);
          }
          resolve(returnData);
        }, (error: any): void => {
          reject(error);
        });
    });
  }

  public render(): React.ReactElement<IHelloCustomApplicationProps> {
    return (
      <div className={ styles.helloCustomApplication }>
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
