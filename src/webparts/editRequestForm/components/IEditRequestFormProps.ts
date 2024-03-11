import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEditRequestFormProps {
  siteUrl:string;
  expenseListTitle:string;
  expenseDetailListTitle:string;
  context:WebPartContext,
  selectedItem:any;
  tabType:string;
}
