export interface IExpenseModel{
  Id:any;
  Expense: string;
  Checkin:any;
  Checkout:any;
  TravelType:string;
  StartMile:string;
  EndMileout:string;
  Description:string;
  ExpenseCost:string;
  isExpenseTypeError:boolean;
  isCheckinError:boolean;
  isCheckoutError:boolean;
  ExpenseTypeErrMsg:string;
  CheckinErrMsg:string;
  CheckoutErrMsg:string;
  isStartMileError:boolean;
  isEndMileError:boolean;
  StartMileErrMsg:string;
  EndMileErrMsg:string;
}

export interface ILogHistoryModel{
  Id:any;
  RequestID:any;
  Expenses:any;
  Author:any;
  CreatedOn:any;
  Status:string;
  CommentsHistory:any;
}