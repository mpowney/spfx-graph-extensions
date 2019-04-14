import { HttpClient } from "@microsoft/sp-http";

export interface IHelloCustomApplicationProps {
  clientResource: string;
  apiEndPoint: string;
  clientId: string;
  aadTenant: string;
  httpClient: HttpClient;
  webPartId: string;
}
