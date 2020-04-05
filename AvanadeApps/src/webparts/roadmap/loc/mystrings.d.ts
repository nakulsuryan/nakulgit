declare interface IRoadmapWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ListNameFieldLabel:string;
  ServiceLineFieldLabel:string;
  SiteUrlFieldLabel:string;
}

declare module 'RoadmapWebPartStrings' {
  const strings: IRoadmapWebPartStrings;
  export = strings;
}
