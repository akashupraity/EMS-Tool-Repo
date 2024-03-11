export interface IExpenseModel{
  Id:any;
  Expense: string;
  Checkin:any;
  Checkout:any;
  ExpenseDate:any;
  TravelType:string;
  StartMile:any;
  EndMileout:any;
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