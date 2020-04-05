declare interface ITabbedUiWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  GroupFieldValueLabel: string;
  LayoutFieldValueLabel:string;
  TitlePlaceholder:string;
  ListNameFieldLabel:string;
  ColumnNameFieldLabel: string;
  ErrorOnLoadingLists: string;
}

declare module 'TabbedUiWebPartStrings' {
  const strings: ITabbedUiWebPartStrings;
  export = strings;
}

