declare interface IEmsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'EmsWebPartStrings' {
  const strings: IEmsWebPartStrings;
  export = strings;
}
