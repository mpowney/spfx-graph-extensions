declare interface IHelloGraphWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  OpenExtensionNameFieldLabel: string;
  OpenExtensionValueFieldLabel: string;
}

declare module 'HelloGraphWebPartStrings' {
  const strings: IHelloGraphWebPartStrings;
  export = strings;
}
