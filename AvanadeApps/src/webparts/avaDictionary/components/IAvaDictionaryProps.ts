import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IAvaDictionaryProps {
  description: string;
  context:WebPartContext;
  urlpanel:string;
  listName:string;
}
