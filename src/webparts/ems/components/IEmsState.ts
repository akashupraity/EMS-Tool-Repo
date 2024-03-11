import { IExpenseListItems } from "./Model/IExpenseModel";


export interface IEmsState {
MySubmissionItems:IExpenseListItems[];
MyPendingItems:IExpenseListItems[];
MyApprovedItems:IExpenseListItems[];
MyRejectedItems:IExpenseListItems[];
MyPaidItems:IExpenseListItems[];
openDialog:boolean;
openInvoiceDialog:boolean;
selectedExpense:any;
SelectedTabType:string;
IsFinanceDept:boolean,
IsManager:boolean,
//currentList:any;

// Edit form states
IExpenseModel:any;
  Department:string;
  ReportHeader:string;
  Status:string;
  StartDate:any;
  EndDate:any;
  Creator:string;
  CreatorEmail:string;
  Manager:string;
  FinanceStatus:string;
  ManagerStatus:string;
  Comments:any;
  TotalExpense:string;
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
  //IsFinanceDept:boolean;
  TravelTypeOption:any;
  FilesToDelete:any;
  ExpenseItemsToDelete:any;
  CurrentUserName:string;
  latestComments:string;
  StartEndMileCosts:any,
  IsBtnClicked:boolean;
  MealExpense:any;
  //openDialog:boolean;
  openEditDialog:boolean;
  openFinanceDialog:boolean;
  openStatusBarDialog:boolean;
  PaidDate:any;
  IsPaidDateErr:boolean;
  PaidDateErrMsg:string;
  ILogHistoryModel:any;
  Finance:string;
  RequestorResponse:string;
  ManagerResponse:string;
  FinanceResponse:string;
  NewFinanceUserID:string;
  FinanceEmailID:any;
  FinanceItemId:number;
  CurrentUserEmail:any;
  OtherFinanceEmailID:string;
}
