import { MSGraphClientFactory } from "@microsoft/sp-http";

export interface IHelloGraphProps {
  openExtensionName: string;
  openExtensionValue: string;
  userLoginName: string;
  msGraphClientFactory: MSGraphClientFactory;
}
