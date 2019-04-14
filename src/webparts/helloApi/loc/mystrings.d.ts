declare interface IHelloApiWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  ClientResourceLabel: string;
  ApiEndPointLabel: string;
}

declare module 'HelloApiWebPartStrings' {
  const strings: IHelloApiWebPartStrings;
  export = strings;
}
