import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface INewRequestFormProps {
  //description: string;
  siteUrl:string;
  expenseListTitle:string;
  expenseDetailListTitle:string;
  context:WebPartContext;
  logHistoryListTitle:string;
}
