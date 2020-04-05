import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IRoadmapProps {
  description: string;
  ServiceLine:string;
  ListName:string;
  SiteUrl:string;
  Context:WebPartContext;
  title: string;
  displayMode:DisplayMode;
  updateProperty:(value: string) => void;  
}
