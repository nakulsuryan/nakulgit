import { DisplayMode } from '@microsoft/sp-core-library';
export interface ITabbedUiProps {
  siteUrl:string;
  view:string;
  title:string;
  titleImageUrl:string;
  description:string;
  tabPublish:string;
  layout:string;
  titleHeader: string;
  setTitle: Function;
  isEditMode: boolean;
  listName:string;
  columnName:string;

       
      
 /* needsConfiguration: boolean;
  configureWebPart: () => void;
  displayMode: DisplayMode;    */  
}

