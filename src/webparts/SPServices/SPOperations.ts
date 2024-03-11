import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp, Web } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IExpenseListItems } from "../ems/components/Model/IExpenseModel";
import { IExpenseModel } from "../editRequestForm/components/Model/IExpenseModel";

export class SPOperations {

    public constructor(public siteUrl: string) { }

    /**
     * CreateExpenses Item
     */
     public async CreateItem(listTitle:string,data:any): Promise<any> {
        let web = Web(this.siteUrl);
        return new Promise<string>(async (resolve, reject) => {
          await web.lists.getByTitle(listTitle).items.add(data)
            .then((result: any) => {
              resolve(result)
            }, (error: any) => {
              reject("error occured " + error);
            })
        })
      };
         /**
     * updateItem
     */
    public UpdateItem(listTitle:string,data:any,itemId:any):Promise<string> {
        let web=Web(this.siteUrl);
        return new Promise<string>(async(resolve,reject)=>{
          web.lists.getByTitle(listTitle).items.getById(itemId).update(data)
          .then((result:any)=>{
              resolve("Updated")
          },(error:any)=>{
              reject("error occured "+error);
          })
        })
    };
    //** Get All Users */
    public getUserIDByEmailFromAllUsers(email:string) :Promise<any> {
      let web = Web(this.siteUrl);
      let userId:number=null;
      return new Promise<any>(async(resolve,reject)=>{
      web.siteUsers.get().then((users: any) => {
        let UsersEmailId=[];
        users.map((userInfo)=>{
          
          UsersEmailId.push({Email:userInfo.Email,Title:userInfo.Title,LoginName:userInfo.LoginName})
          
        })
        console.log(UsersEmailId);
        const user = users.find(u => u.Email === email);
            if (user) {
              userId= user.id;
            } else {
              throw new Error("User not found.");
            }
        resolve(userId);
      },(error:any)=>{
          reject("error occured "+error);
      })
    })
};
    //**Get Current User**/
    public GetCurrentUser() :Promise<string> {
        let web = Web(this.siteUrl);
        return new Promise<string>(async(resolve,reject)=>{
        web.currentUser.get().then((result: any) => {
          resolve(result)
        },(error:any)=>{
            reject("error occured "+error);
        })
      })
  };
        //*get Current User Details **/
        public getCurrentUserDetails(empName:string):Promise<string>{
            return new Promise<string>(async(resolve,reject)=>{
          sp.profiles.getPropertiesFor(empName).then((profile: any) => {
          var properties = {};
          profile.UserProfileProperties.forEach(function(prop) {
          properties[prop.Key] = prop.Value;
          });
          resolve(properties["Manager"])
        },(error:any)=>{
            reject("error occured "+error);
        })
      })
  };
    
       //* get Manager Details**/
       public getManagerDetails(user:string):Promise<string>{
          return new Promise<string>(async(resolve,reject)=>{
          sp.profiles.getPropertiesFor(user).then((profile: any) => {
          var properties = {};
          profile.UserProfileProperties.forEach(function(prop) {
          properties[prop.Key] = prop.Value;
          });
          resolve(properties["WorkEmail"])
        },(error:any)=>{
            reject("error occured "+error);
        })
      })
  };
         //* get User ID by Email **/
         public getUserIDByEmail(email: string):Promise<any> {
          let web = Web(this.siteUrl);
          return new Promise<any>(async(resolve,reject)=>{
          web.siteUsers.getByEmail(email).get().then(user => {
            console.log('User Id: ', user.Id);
            resolve(user.Id)
        },(error:any)=>{
            reject("error occured "+error);
        })
      })
  };

  //** Convert dates */
  public ConvertDate(dateValue){
    var d = new Date(dateValue);
    var strDate =  d.getDate()+ "/" + (d.getMonth()+1) + "/" + d.getFullYear();
    return strDate;
  }
  //* get ListItems by login name **/
  public getListItems(email: string,requestType,userType):Promise<IExpenseListItems[]> {
    let listItems:IExpenseListItems[]=[];
    let query:string="";
    if(requestType=="MySubmission"){
     query=`Author/EMail eq '${email}'`;
    }
    if(requestType=="MyTask"){
      if(userType=="Manager"){
      query=`Manager/EMail eq '${email}'`;
      }
      if(userType=="Finance"){
        query=`(Finance/EMail eq '${email}' and ReviewForFinace eq 'Yes') and ((ManagerStatus eq 'Approved') or (ManagerStatus eq 'Rejected by Finance') or (ManagerStatus eq 'Manager Approval Not Required'))`;
        }
    }
    let web = Web(this.siteUrl);
    return new Promise<any>(async(resolve,reject)=>{
    web.lists.getByTitle('Expenses').items
    .filter(query)
    .select("*","Manager/Title","Manager/ID","Manager/EMail","Author/Title","Author/ID","Author/EMail","Finance/Title","Finance/ID","Finance/EMail").expand("Manager,Author,Finance")
    .orderBy("Modified",false)
    .top(4999)
    .get().then(results => {
      console.log(listItems);

      results.map((item)=>{
        let StatusValue="";
        if(requestType=="MySubmission"){
          StatusValue=item.Status;
         }
        if(requestType=="MyTask"){
          if(userType=="Manager"){
            StatusValue=item.ManagerStatus;
          }
          if(userType=="Finance"){
            StatusValue=item.FinanceStatus;
            }
        }
        listItems.push({
          Department:item.Department,
          ReportHeader: item.Title,
          StartDate:item.StartDate!=null?this.ConvertDate(item.StartDate):null,
          EndDate:item.EndDate!=null?this.ConvertDate(item.EndDate):null,
          RequestorID:item.RequestorID,
          Status:StatusValue,
          Manager:item.Manager!=undefined?item.Manager.Title:"",
          Creator:item.Author!=undefined?item.Author.Title:"",
          ID:item.ID,
          FinanceStatus:item.FinanceStatus,
          ReviewForFinace:item.ReviewForFinace,
          ManagerStatus:item.ManagerStatus,
          AmountPaidDate:item.AmountPaidDate!=null?this.ConvertDate(item.AmountPaidDate):null,
          TotalExpense:item.TotalExpense

        });
      })
      resolve(listItems)
  },(error:any)=>{
      reject("error occured "+error);
  })
})
};
//* get ListItem by Item ID **/
public GetListItemByID(itemId: any,listName:string):Promise<any> {
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
    web.lists.getByTitle(listName).items.getById(itemId).select("*","Manager/Title","Manager/ID","Manager/EMail","Author/Title","Author/ID","Author/EMail","Finance/Title","Finance/ID","Finance/EMail").expand("Manager,Author,Finance").get().then(results => {
    console.log(results);
    resolve(results)
},(error:any)=>{
    reject("error occured "+error);
})
})
};
//* get Log History by Expense ID **/
public GetLogHistoryItems(itemId: any,listName:string):Promise<any> {
  let logHistoryItems:any[]=[];
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
    web.lists.getByTitle(listName).items.filter(`ExpensesId eq `+itemId).select("*","Author/Title","Author/ID","Author/EMail","Expenses/ID","Expenses/Title").expand("Author,Expenses").orderBy("Id",false).get().then(results => {
      results.map((item)=>{
        logHistoryItems.push({
          Id:item.ID,
          RequestID:item.Title,
          Expenses:item.Expenses!=undefined?item.Expenses.Title:"",
          Author:item.Author.Title,
          CreatedOn:this.ConvertDateYYMMDD(item.Created),
          Status:item.Status,
          CommentsHistory:item.CommentsHistory!=undefined?item.CommentsHistory.replace(/<\/?[^>]+(>|$)/g, ""):"",
        });
      })
      resolve(logHistoryItems);
},(error:any)=>{
    reject("error occured "+error);
})
})
};
 //** Convert dates */
 public ConvertDateYYMMDD(dateValue){
  var d = new Date(dateValue),
     month = '' + (d.getMonth() + 1),
     day = '' + d.getDate(),
     year = d.getFullYear();

 if (month.length < 2) month = '0' + month;
 if (day.length < 2) day = '0' + day;

 return [year, month, day].join('-');
};
//* get expense detail by lookup ID **/
public GetExpenseDetails(itemId: any,listName:string):Promise<any> {
  let listItems:any[]=[];
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
    web.lists.getByTitle(listName).items.select("*").filter("ExpensesId eq "+itemId+"").get().then(results => {
      results.map((item)=>{
        listItems.push({
          Expense: item.ExpenseTypes,
          Checkin: item.CheckIn?this.ConvertDateYYMMDD(item.CheckIn):null,
          Checkout: item.CheckOut?this.ConvertDateYYMMDD(item.CheckOut):null,
          ExpenseDate: item.ExpenseDate?this.ConvertDateYYMMDD(item.ExpenseDate):null,
          TravelType: item.TravelType,
          StartMile: item.StartMile,
          EndMileout: item.EndMile,
          Description: item.Description,
          ExpenseCost: item.ExpenseCost,
          id:item.Id,
          Id:item.Id
        });
      })
      resolve(listItems);
},(error:any)=>{
    reject("error occured "+error);
})
})
};
//* get ListItems by login name **/
public getEMSConfigListItems():Promise<any[]> {
  let web = Web(this.siteUrl);
  return new Promise<any>(async(resolve,reject)=>{
  web.lists.getByTitle('EMSConfiguration').items
  .select("*","Finance/Title","Finance/ID","Finance/EMail","OtherFinance/Title","OtherFinance/ID","OtherFinance/EMail").expand("Finance,OtherFinance")
  .orderBy("Modified",false)
  .top(4999)
  .get().then(results => {
    resolve(results)
},(error:any)=>{
    reject("error occured "+error);
})
})
};
//** Get Today Date */
public GetTodaysDate (){
  var d = new Date();
  let month = '' + (d.getMonth() + 1),
  day = '' + d.getDate(),
  year = d.getFullYear();

if (month.length < 2) month = '0' + month;
if (day.length < 2) day = '0' + day;

return [day, month, year].join('-');
}
}