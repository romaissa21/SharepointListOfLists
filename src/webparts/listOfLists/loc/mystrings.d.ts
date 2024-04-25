declare interface IListOfListsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'ListOfListsWebPartStrings' {
  const strings: IListOfListsWebPartStrings;
  export = strings;
}
