import { MSGraphClientFactory } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IOutlook2SharePointProps {
  mail: any;
  context:WebPartContext;
}
