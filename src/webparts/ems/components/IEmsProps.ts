import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEmsProps {
  description: string;
  siteUrl:string;
  context:WebPartContext;
  expenseListTitle:string;
  expenseDetailListTitle:string;
  logHistoryListTitle:string;

}
