import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IEditRequestFormState {
  IExpenseModel:any;
  Department:string;
  ReportHeader:string;
  Status:string;
  StartDate:any;
  EndDate:any;
  Comments:string;
  DeptErrMsg:string;
  IsDeptErr:boolean;
  ReportHeaderErrMsg:string;
  IsReportErr:boolean;
  StartDateErrMsg:string;
  IsStartDateErr:boolean;
  EndDateErrMsg:string;
  IsEndDateErr:boolean;
  ManagerId: number;
  filePickerResult:any;
  Attachments: any;
  RemoveFiles:any;
  fileInfos:any;
  isMealExpenseCostError:boolean;
  mealExpenseCostErrMsg:string;
  ExpenseDetailErrMsg:string;
  IsExpenseDetailErr:boolean;
  DepartmentOptions:any;
  ExpenseTypeOptions:any;
  FinanceId:any;
  IsFinanceDept:boolean;
  TravelTypeOption:any;
  FilesToDelete:any;
  ExpenseItemsToDelete:any;
  CurrentUserName:string;
  latestComments:string;
  StartEndMileCosts:any,
  IsBtnClicked:boolean;
  MealExpense:any;
  openDialog:boolean;
  openEditDialog:boolean;
  PaidDate:any;
  IsPaidDateErr:boolean;
  PaidDateErrMsg:string;
}
