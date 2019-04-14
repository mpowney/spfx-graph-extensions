declare interface IHelloCustomApplicationWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ClientResourceLabel: string;
  ApiEndPointLabel: string;
  ClientIdLabel: string;
  AadTenantLabel: string;
}

declare module 'HelloCustomApplicationWebPartStrings' {
  const strings: IHelloCustomApplicationWebPartStrings;
  export = strings;
}
