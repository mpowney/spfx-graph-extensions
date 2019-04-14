import { AadHttpClientFactory } from "@microsoft/sp-http";

export interface IHelloApiProps {
  clientResource: string;
  apiEndPoint: string;
  aadClientFactory: AadHttpClientFactory;
}
