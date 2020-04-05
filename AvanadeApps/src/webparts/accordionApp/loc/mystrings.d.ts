declare interface IAccordionAppWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  SiteUrlFieldLabel:string;   
  ListUrlFieldLabel:string;
  TitleFieldLabel:string;
  TitlePlaceholder:string;
  DescriptionFieldLabel:string;
  GroupFieldLabel: string;
  GroupFieldValueLabel:string;
  ErrorOnLoadingLists: string;
}

declare module 'AccordionAppWebPartStrings' {
  const strings: IAccordionAppWebPartStrings;
  export = strings;
}
