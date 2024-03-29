export interface INewRequestFormState {
  IExpenseModel:any;
  Department:string;
  ReportHeader:string;
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
  SelectedFiles:any;
  fileInfos:any;
  isMealExpenseCostError:boolean;
  mealExpenseCostErrMsg:string;
  ExpenseDetailErrMsg:string;
  IsExpenseDetailErr:boolean;
  DepartmentOptions:any;
  ExpenseTypeOptions:any;
  FinanceId:any;
  TravelTypeOption:any;
  CurrentUserName:string;
  CurrentUserID:string;
  StartEndMileCosts:any;
  IsBtnClicked:boolean;
  MealExpense:any;
  loading:boolean;
}
