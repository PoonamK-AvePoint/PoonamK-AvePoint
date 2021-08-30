import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISpfxSearchProps {
  description: string;
  wContext: WebPartContext;
  queryTemplate: string;
}
