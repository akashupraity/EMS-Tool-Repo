import * as React from 'react';
import styles from './NewRequestForm.module.scss';
import { INewRequestFormProps } from './INewRequestFormProps';
import { INewRequestFormState } from './INewRequestFormState'
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label, Icon, PrimaryButton } from 'office-ui-fabric-react';
import { Shimmer, ShimmerElementsGroup, ShimmerElementType } from 'office-ui-fabric-react/lib/Shimmer'; 
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import * as $ from "jquery";
import * as bootstrap from "bootstrap";
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { sp, Web } from '@pnp/sp/presets/all'
import { SPOperations } from '../../SPServices/SPOperations';
import { ListItemAttachments } from "@pnp/spfx-controls-react/lib";
import { float } from 'html2canvas/dist/types/css/property-descriptors/float';


const stackTokens = { childrenGap: 50 };
const iconProps = { iconName: 'Calendar' };
const stackStyles: Partial<IStackStyles> = { root: { width: 1080 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 700 } },
};

export default class NewRequestForm extends React.Component<INewRequestFormProps, INewRequestFormState> {
  public _spOps: SPOperations;
  constructor(props: INewRequestFormProps) {
    super(props)
    this.state = {
      IExpenseModel: [],
      Department: "",
      ReportHeader: "",
      StartDate: null,
      EndDate: null,
      Comments: "",
      DeptErrMsg: "",
      IsDeptErr: false,
      ReportHeaderErrMsg: "",
      IsReportErr: false,
      StartDateErrMsg: "",
      IsStartDateErr: false,
      EndDateErrMsg: "",
      IsEndDateErr: false,
      ManagerId: null,
      filePickerResult: [],
      Attachments: [],
      SelectedFiles: [],
      fileInfos: [],
      isMealExpenseCostError: false,
      mealExpenseCostErrMsg: "",
      ExpenseDetailErrMsg: "",
      IsExpenseDetailErr: false,
      DepartmentOptions: [],
      ExpenseTypeOptions: [],
      FinanceId: null,
      TravelTypeOption: [],
      CurrentUserName: "",
      StartEndMileCosts: null,
      IsBtnClicked: false,
      MealExpense: null,
      CurrentUserID:null,
      loading:false,

    }

    this._spOps = new SPOperations(this.props.siteUrl);

  }
  GetCurrentUserManagerId = () => {
    this._spOps.GetCurrentUser().then((result) => {
      this.setState({
        CurrentUserName: result["Title"],
        CurrentUserID:result["ID"]
      })
      this._spOps.getCurrentUserDetails(result["LoginName"]).then((manager) => {
        if(manager!=""){
        this._spOps.getManagerDetails(manager).then((managerEmail) => {
          this._spOps.getUserIDByEmail(managerEmail).then((managerId) => {
            this.setState({
              ManagerId: managerId,
            })
          })
        })
      }
      })
    
    })
  }
  public GetEMSConfig() {
    this._spOps.getEMSConfigListItems().then((response: any) => {

      let AllFinanceEmail = [];
      let AllDepartment = [];
      let AllExpenseType = [];
      let AllTravelType = [];
      let MileCost = [];
      let MealExpense = [];
      response.map((item) => {
        if (item.Title == "FinanceDepartment") {
          AllFinanceEmail.push(item);
        }
        if (item.Title == "MileCost") {
          MileCost.push(item);
        }
        if (item.Title == "MealExpense") {
          MealExpense.push(item);
        }
        if (item.Title == "Department") {
          let dept = { key: "", text: "" };
          dept.key = item.Key;
          dept.text = item.Value
          AllDepartment.push(dept);
        }
        if (item.Title == "ExpenseType") {
          let expensetypeItems = { key: "", text: "" };
          expensetypeItems.key = item.Key;
          expensetypeItems.text = item.Value
          AllExpenseType.push(expensetypeItems);
        }
        if (item.Title == "TravelType") {
          let travelTypeItems = { key: "", text: "" };
          travelTypeItems.key = item.Key;
          travelTypeItems.text = item.Value
          AllTravelType.push(travelTypeItems);
        }

      })
      let FinanceUserId = AllFinanceEmail[0].Finance["ID"];
      let startEndMileCost = MileCost[0].Key;
      let mealExpenseCost = MealExpense[0].Key;
      this.setState({
        DepartmentOptions: AllDepartment,
        ExpenseTypeOptions: AllExpenseType,
        FinanceId: FinanceUserId,
        TravelTypeOption: AllTravelType,
        StartEndMileCosts: startEndMileCost,
        MealExpense: parseInt(mealExpenseCost)
      })
    })
  }
  componentDidMount(): void {
    this.GetCurrentUserManagerId();
    this.GetEMSConfig();
  }

  /* This event will fire on remove row */
  handleRemoveRow = () => {
    try {
      var rowsArray = this.state.IExpenseModel;
      if (rowsArray.length > 1) {
        var newRow = rowsArray.slice(0, -1);
        this.setState({ IExpenseModel: newRow });
      }
      //this.props.tableData(newRow);    
    } catch (error) {
      console.log("Error in React Table handle Remove Row : " + error);
    }
  };
  RefreshPage = () => {
    window.location.reload();
    // $("#file").val('');
    // $('input[type=date]').val('');
    // this.setState({
    //   IExpenseModel: [],
    //   Department: "",
    //   ReportHeader: "",
    //   StartDate: null,
    //   EndDate: null,
    //   Comments: "",
    //   DeptErrMsg: "",
    //   IsDeptErr: false,
    //   ReportHeaderErrMsg: "",
    //   IsReportErr: false,
    //   StartDateErrMsg: "",
    //   IsStartDateErr: false,
    //   EndDateErrMsg: "",
    //   IsEndDateErr: false,
    //   ManagerId: null,
    //   filePickerResult: [],
    //   Attachments: [],
    //   SelectedFiles: [],
    //   fileInfos: [],
    //   isMealExpenseCostError: false,
    //   mealExpenseCostErrMsg: "",
    //   ExpenseDetailErrMsg: "",
    //   IsExpenseDetailErr: false,
    //   DepartmentOptions: [],
    //   ExpenseTypeOptions: [],
    //   FinanceId: null,
    //   TravelTypeOption: [],
    //   CurrentUserName: "",
    //   StartEndMileCosts: null,
    //   IsBtnClicked: false,
    //   MealExpense: null,
    // })
    // this.componentDidMount();
  }
  //** Generate Requestor unique ID */
  public GetUniqueRequestorID = (expenseItemId: number) => {
    let expItemId = expenseItemId.toString();
    var uniqueID = "";
    if (expenseItemId < 10) {
      uniqueID = "000" + expItemId
    }
    if (expenseItemId >= 10 && expenseItemId < 100) {
      uniqueID = "00" + expItemId
    }
    if (expenseItemId >= 100 && expenseItemId < 1000) {
      uniqueID = "0" + expItemId
    }
    if (expenseItemId >= 1000) {
      uniqueID = expItemId;
    }
    return "EMS-" + uniqueID;
  }
 // save all expenses into "ExpenseDetails" list in batches
  private async AddExpensesDetailsAsBatch(Expenses: any[], expenseItemId: number) {
    let requestorUniqueID = this.GetUniqueRequestorID(expenseItemId);
    let sourceWeb = Web(this.props.siteUrl);
    let taskList = sourceWeb.lists.getByTitle(this.props.expenseDetailListTitle);
    let batch = sourceWeb.createBatch();
    console.log("batch = ", JSON.stringify(batch));
    console.log("batch baseURL = ", batch["baseUrl"]);
    for (let i = 0; i < Expenses.length; i++) {
      taskList.items.inBatch(batch).add(
          {
        ExpenseTypes: Expenses[i].Expense,
        CheckIn: Expenses[i].Checkin != "" ? Expenses[i].Checkin : null,
        CheckOut: Expenses[i].Checkout != "" ? Expenses[i].Checkout : null,
        ExpenseDate:Expenses[i].ExpenseDate!="" ? Expenses[i].ExpenseDate:null,
        TravelType: Expenses[i].TravelType,
        StartMile: Expenses[i].StartMile,
        EndMile: Expenses[i].EndMileout,
        Description: Expenses[i].Description,
        ExpenseCost: String(Expenses[i].ExpenseCost),
        RequestorID: requestorUniqueID,
        ExpensesId: expenseItemId
          }
        )
        .then((result:any) => {
          console.log("Item created with id", result.data.Id);
        })
        .catch((ex) => {
          console.log(ex);
        });
    }
    await batch.execute();
    console.log("Done");
    }
  // async function with await to save all expenses into "ExpenseDetails" list
  // private async AddExpensesDetails(Expenses: any[], expenseItemId: number) {
  //   let web = Web(this.props.siteUrl);
  //   let requestorUniqueID = this.GetUniqueRequestorID(expenseItemId);
  //   for (const expense of Expenses) {
  //     await web.lists.getByTitle(this.props.expenseDetailListTitle).items.add({
  //       ExpenseTypes: expense.Expense,
  //       CheckIn: expense.Checkin != "" ? expense.Checkin : null,
  //       CheckOut: expense.Checkout != "" ? expense.Checkout : null,
  //       ExpenseDate:expense.ExpenseDate!="" ? expense.ExpenseDate:null,
  //       TravelType: expense.TravelType,
  //       StartMile: expense.StartMile,
  //       EndMile: expense.EndMileout,
  //       Description: expense.Description,
  //       ExpenseCost: String(expense.ExpenseCost),
  //       RequestorID: requestorUniqueID,
  //       ExpensesId: expenseItemId
  //     });
  //   }
  // }

  //**Validation on fields */
  ValidateForm = (submissionType) => {
    var tableArr = this.state.IExpenseModel;
    var isErrExists = false;
    this.setState({ IsDeptErr: false, IsReportErr: false, isMealExpenseCostError: false, IsStartDateErr: false, IsEndDateErr: false });


    if (this.state.Department == "") {
      this.setState({
        DeptErrMsg: "Please select Department",
        IsDeptErr: true
      })
      isErrExists = true;
    }
    if (submissionType == "Submitted" || submissionType == "Draft") {
      if (this.state.ReportHeader == "") {
        this.setState({
          ReportHeaderErrMsg: "Please enter Description",
          IsReportErr: true
        })
        isErrExists = true;
      }
      if (this.state.StartDate != null && this.state.EndDate != null) {
        var new_start_date = new Date(this.state.StartDate);
        var new_end_date = new Date(this.state.EndDate);
        if (new_start_date > new_end_date) {
          this.setState({
            StartDateErrMsg: "Start Date should be less than End Date",
            IsStartDateErr: true
          })
          isErrExists = true;
        }
        if (new_end_date < new_start_date) {
          this.setState({
            EndDateErrMsg: "End Date should be greater than Start Date",
            IsEndDateErr: true
          })
          isErrExists = true;
        }
      }
      if (this.state.IExpenseModel.length == 0) {
        this.setState({
          ExpenseDetailErrMsg: "Add Expense Details",
          IsExpenseDetailErr: true
        })
        isErrExists = true;
      }
      // let {fileInfos}=this.state;
      tableArr.map((item, key) => {
        item.isExpenseTypeError = false;
        item.isExpenseCostError=false;
        item.isCheckinError = false;
        item.isCheckoutError = false;
        item.isEndMileError = false;
        if (item.Expense == "") {
          item.isExpenseTypeError = true;
          item.ExpenseTypeErrMsg = "Please select Expense Type";
          isErrExists = true;
        }
        if (item.ExpenseCost == ""||item.ExpenseCost == 0) {
          item.isExpenseCostError = true;
          item.ExpenseCostMsg = "Please Type Expense Cost";
          isErrExists = true;
        }
        

        if (item.Expense == "Hotel" && item.Checkin == "") {
          item.isCheckinError = true;
          item.CheckinErrMsg = "Please select chekin date";
          isErrExists = true;
        }
        if (item.Expense == "Hotel" && item.Checkout == "") {
          item.isCheckoutError = true;
          item.CheckoutErrMsg = "Please select checkout date";
          isErrExists = true;
        }
        if (item.Expense == "Hotel" && item.Checkout != "" && item.Checkin != "") {
          var new_checkout = new Date(item.Checkout);
          var new_checkin = new Date(item.Checkin);
          if (new_checkout <= new_checkin) {
            item.isCheckoutError = true;
            item.CheckoutErrMsg = "Check Out date should be greater than Check in date";
            isErrExists = true;
          }

        }
        if (item.Expense == "Travel" && item.StartMile != null && item.EndMileout != null) {
          let startMile = parseFloat(item.StartMile);
          let endMile = parseFloat(item.EndMileout);
          if (startMile > endMile) {
            item.isEndMileError = true;
            item.EndMileErrMsg = "End Mile should be greater than Start Mile";
            isErrExists = true;
          }
        }
        if (item.Expense == "Meal" && item.ExpenseCost >= this.state.MealExpense && this.state.Attachments.length == 0) {
          this.setState({
            isMealExpenseCostError: true,
            mealExpenseCostErrMsg: "Please attach attachement"
          })
          isErrExists = true;
        }
      })
    }
    this.setState({ IExpenseModel: tableArr });
    return isErrExists
  }
  //* Update Unqique Id and Add expense details in ExpenseDetails list*/
  updateUniqueID = (requestorUniqueID, itemId, submissionType: string) => {
    let updatePostDate = {
      RequestorID: requestorUniqueID,
    }
    this._spOps.UpdateItem(this.props.expenseListTitle, updatePostDate, itemId).then((response) => {
      if (this.state.IExpenseModel.length > 0) {
        this.AddExpensesDetailsAsBatch(this.state.IExpenseModel, itemId).then(() => {
          this.setState({
            loading:false
          })

        $('#loader').hide();
          alert(submissionType == "Submitted" ? "Request submitted sucessfully" : "Request drafted sucessfully");
          this.RefreshPage();
        });
      } else {
        this.setState({
          loading:false
        })
       $('#loader').hide();
        alert(submissionType == "Submitted" ? "Request submitted sucessfully" : "Request drafted sucessfully");
        this.RefreshPage();
      }
    })
  }
  //** call validation method and create item into list */
  SubmitRequest = (submissionType) => {
    var isError = this.ValidateForm(submissionType)
    if (!isError) {
      // show indicator 
      $('#loader').show();
      this.setState({
        IsBtnClicked: true,
        loading:true
      })
      let commentsHTML = "";
      let todayDate = this._spOps.GetTodaysDate();
      commentsHTML = '<strong>' + this.state.CurrentUserName + ' : ' + todayDate + '</strong>' + '<div>' + this.state.Comments + '</div>';
      let totalAmount: any = 0;
      if (this.state.IExpenseModel.length > 0) {
        this.state.IExpenseModel.map((expense) => {
          totalAmount += expense.ExpenseCost != undefined ? parseFloat(expense.ExpenseCost) : 0;
        })
      }
      let createPostData: any = {};
       createPostData = {
        Department: this.state.Department,
        Title: this.state.ReportHeader,
        StartDate: this.state.StartDate != "" ? this.state.StartDate : null,
        EndDate: this.state.EndDate != "" ? this.state.EndDate : null,
        Status: submissionType == "Submitted" && this.state.ManagerId==null?"InProgress":submissionType,
        ManagerId: submissionType == "Submitted"?this.state.ManagerId:null,
        FinanceId: submissionType == "Submitted"?this.state.FinanceId:null,
        Comments: this.state.Comments != "" ? commentsHTML : this.state.Comments,
        ManagerStatus: submissionType == "Submitted" && this.state.ManagerId!=null? "Pending for Manager" : "",
        TotalExpense: totalAmount.toFixed(2),
      }
      if(submissionType == "Submitted" && this.state.ManagerId==null){
        createPostData.FinanceStatus="Pending for Finance";
        createPostData.ReviewForFinace="Yes";
        createPostData.ManagerStatus= 'Manager Approval Not Required';
      }
      this._spOps.CreateItem(this.props.expenseListTitle, createPostData).then((result: any) => {
        console.log(result.data.ID);
        let itemId = result.data.ID;
        let requestorUniqueID = this.GetUniqueRequestorID(itemId);
        let logHistoryPostData:any={};
        logHistoryPostData={
          Title:requestorUniqueID,
          ExpensesId:itemId,
          CommentsHistory:this.state.Comments,
          Status:submissionType == "Submitted" && this.state.ManagerId==null?"InProgress":submissionType,
          //NameId:this.state.CurrentUserID
        }
        this._spOps.CreateItem(this.props.logHistoryListTitle, logHistoryPostData).then((result: any) => {});
        let { fileInfos } = this.state;
        if (fileInfos.length > 0) {
          let web = Web(this.props.siteUrl);
          web.lists.getByTitle(this.props.expenseListTitle).items.getById(itemId).attachmentFiles.addMultiple(fileInfos).then(() => {
            this.updateUniqueID(requestorUniqueID, itemId, submissionType);
          });
        }
        else {
          this.updateUniqueID(requestorUniqueID, itemId, submissionType);
        }
      })

    }
  }

  /* This event will fire on remove specific row */
  handleRemoveSpecificRow = (idx) => () => {
    try {
      const rows = this.state.IExpenseModel
      if (rows.length > 1) {
        rows.splice(idx, 1);
      }

      this.setState({ IExpenseModel: rows });
    } catch (error) {
      console.log("Error in React Table handle Remove Specific Row : " + error);
    }
  }
  /* This event will fire on adding new row */
  handleAddRow = () => {
    try {
      var id = (+ new Date() + Math.floor(Math.random() * 999999)).toString(36);
      const tableColProps = {
        id: id,
        Expense: "",
        Checkin: "",
        Checkout: "",
        ExpenseDate:"",
        TravelType: "",
        StartMile: null,
        EndMileout: null,
        Description: "",
        ExpenseCost: "",
        isExpenseTypeError: false,
        isExpenseCostError:false,
        ExpenseCostMsg:"",
        isCheckinError: false,
        isCheckoutError: false,
        ExpenseTypeErrMsg: "",
        CheckinErrMsg: "",
        CheckoutErrMsg: "",

      }
      this.state.IExpenseModel.push(tableColProps);
      this.setState(this.state.IExpenseModel);
    } catch (error) {
      console.log("Error in React Table handle Add Row : " + error)
    }
  };
  /* This event will fire on change of every fields on form */
  handleChange = (index: any) => evt => {
    try {
      var item = {
        id: evt.target.id,
        name: evt.target.name,
        value: evt.target.value
      };
      if (item.name == "Department") {
        this.setState({
          Department: item.value
        })
      }
      if (item.name == "ReportHeader") {
        this.setState({
          ReportHeader: item.value
        })
      }
      if (item.name == "StartDate") {
        this.setState({
          StartDate: item.value
        })
      }
      if (item.name == "EndDate") {
        this.setState({
          EndDate: item.value
        })
      }
      if (item.name == "Comments") {
        this.setState({
          Comments: item.value
        })
      }
      var rowsArray = this.state.IExpenseModel;
      var newRow = rowsArray.map((row, i) => {
        for (var key in row) {
          if (key == item.name && row.id == item.id) {
            row[key] = item.value;
            if (item.name == "TravelType") {
              if (item.value == "Cab/Uber/Lyft") {
                row['StartMile'] = "N/A";
                row['EndMileout'] = "N/A";
              }
              if (item.value == "Airfare") {
                row['StartMile'] = "N/A";
                row['EndMileout'] = "N/A";
              }
              row["ExpenseCost"] = "";
            }
            if (item.name == "EndMileout") {
              let mileDiff = parseFloat(row.EndMileout) - parseFloat(row.StartMile);
              let milageAmt = this.state.StartEndMileCosts * mileDiff;
              row['ExpenseCost'] = milageAmt.toFixed(2);

            }
          }
        }
        return row;
      });
      this.setState({ IExpenseModel: newRow });

    } catch (error) {
      console.log("Error in React Table handle change : " + error);
    }
  };


  // * On Select of File. Read filename and content*/
  private addFile(event) {
    //let resultFile = document.getElementById('file');
    let resultFile = event.target.files;
    console.log(resultFile);
    //let fileInfos = [];
    let fileInformations = [];
    let selectedFileNames = [];
    for (var i = 0; i < resultFile.length; i++) {
      var fileName = resultFile[i].name;
      selectedFileNames.push(fileName);
      console.log(fileName);
      var file = resultFile[i];
      var reader = new FileReader();
      reader.onload = (function (file) {
        return function (e) {
          //Push the converted file into array
          fileInformations.push({
            "name": file.name,
            "content": e.target.result
          });

        }

      })(file);

      reader.readAsArrayBuffer(file);
    }
    setTimeout(
      function () {
        let tempArr = this.state.fileInfos;
        tempArr.push.apply(tempArr, fileInformations)
        this.setState({ fileInfos: tempArr, Attachments: resultFile });
      }
        .bind(this),
      100
    );
  };

  //**Remove specific attachement */
  removeSpecificAttachment = (idx) => () => {
    try {
      const rows = this.state.fileInfos;
      rows.splice(idx, 1);
      this.setState({ fileInfos: rows });
    } catch (error) {
      console.log("Error in React Table handle Remove Specific Row : " + error);
    }
  }

  //**render selected attachments */
  renderAttachmentName() {
    return this.state.fileInfos.map((item, idx) => {
      return (<div><a href="javascript:void(0)" target="_blank">{item.name}</a>&nbsp;&nbsp; <span id="delete-spec-row" onClick={this.removeSpecificAttachment(idx)} className={styles.deleteIcon}>
        <Icon iconName="delete" className="ms-IconExample" />
      </span></div>)
    })
  }

  //**Common function to render Department and Expense Types DropDowns */
  renderDropdown = (options) => {
    return options.map((item, idx) => {
      return (<option value={item.key}>{item.text}</option>)
    })
  }
  // Render Expense Table
  renderTableData() {
    return this.state.IExpenseModel.map((item, idx) => {
      let costTxtCss="form-group col-md-3";
      if(item.Expense == "Travel"){
        costTxtCss="form-group col-md-1"
      }
      if(item.Expense == "Hotel"){
        costTxtCss="form-group col-md-2"
      }
      return (<div key={idx}>
        <div className={styles.renderExpenseTbl}>
          <div className="form-group col-md-2">
            <label className="control-label">Expense Type</label><span className={styles.star}>*</span>
            <select className='form-control ExpTypeSelectOptions' name="Expense" value={this.state.IExpenseModel[idx].Expense} id={this.state.IExpenseModel[idx].id} onChange={this.handleChange(idx)}>
              <option value="">Select</option>
              {this.renderDropdown(this.state.ExpenseTypeOptions)}
            </select>
            {item.isExpenseTypeError == true &&
              <span className={styles.errMsg}>{item.ExpenseTypeErrMsg}</span>
            }
          </div>
          {item.Expense == "Hotel" &&
            <div className="form-group col-md-2">
              <label className="control-label">Check In &nbsp;<span className={styles.noteMsg}>(MM-DD-YYYY)</span><span className={styles.star}>*</span></label>
              <input
                placeholder='check in'
                type="date"
                className='form-control'
                name="Checkin"
                value={this.state.IExpenseModel[idx].Checkin}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
                onKeyDown={(e) => e.preventDefault()}
              />

              {item.isCheckinError == true && 
              <span className={styles.errMsg}>{item.CheckinErrMsg}</span>}
            </div>
          }
          {item.Expense == "Hotel" &&
            <div className="form-group col-md-2">
              <label className="control-label">Check Out &nbsp;<span className={styles.noteMsg}>(MM-DD-YYYY)</span> <span className={styles.star}>*</span></label>
              <input
                placeholder='Check out'
                type="date"
                className='form-control'
                name="Checkout"
                value={this.state.IExpenseModel[idx].Checkout}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
                onKeyDown={(e) => e.preventDefault()}
              />
              {item.isCheckoutError == true && <span className={styles.errMsg}>{item.CheckoutErrMsg}</span>}
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-2">
              <label className="control-label">Travel Type</label>
              <select className='form-control' name="TravelType" value={this.state.IExpenseModel[idx].TravelType} id={this.state.IExpenseModel[idx].id} onChange={this.handleChange(idx)}>
                <option value="">Select</option>
                {this.renderDropdown(this.state.TravelTypeOption)}
              </select>
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-1">
              <label className={styles.lblWidth}>Start Mile</label>
              <input
                placeholder={item.TravelType == "Cab/Uber/Lyft" || item.TravelType == "Airfare"?'N/A':'Start'}
                type="number"
                min="1"
                disabled={item.TravelType == "Cab/Uber/Lyft" || item.TravelType == "Airfare" ? true : false}
                className={styles.formControl}
                name="StartMile"
                value={this.state.IExpenseModel[idx].StartMile}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-1">
              <label className="control-label">End Mile</label>
              <input
                placeholder={item.TravelType == "Cab/Uber/Lyft" || item.TravelType == "Airfare"?'N/A':'End'}
                type="number"
                min="1"
                disabled={(item.TravelType == "Cab/Uber/Lyft" || item.TravelType == "Airfare") ? true : false}
                className={styles.formControl}
                name="EndMileout"
                value={this.state.IExpenseModel[idx].EndMileout}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
              {item.isEndMileError && <span className={styles.errMsg}>{item.EndMileErrMsg}</span>}
            </div>
          }
          {/* {(item.Expense == "Meal" || item.Expense == "Others") && */}
            <div className={item.Expense == "Travel"?"form-group col-md-2":"form-group col-md-3"}>
              <label className="control-label">Description</label>
              <input
                placeholder='Description'
                type="text"
                className='form-control'
                name="Description"
                value={this.state.IExpenseModel[idx].Description}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
                maxLength={254}
              />
            </div>
          {/* } */}
            {(item.Expense != "Hotel") &&
           <div className="form-group col-md-2">
              <label className="control-label">Date &nbsp;<span className={styles.noteMsg}>(MM-DD-YYYY)</span></label>
              <input
                placeholder='Date'
                type="date"
                className='form-control'
                name="ExpenseDate"
                value={this.state.IExpenseModel[idx].ExpenseDate}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
                onKeyDown={(e) => e.preventDefault()}
              />
            </div>
    }
          <div className={costTxtCss}>
            <label className={styles.amtLbl}>Amount ($)<span className={styles.star}>*</span> &nbsp;{item.Expense == "Travel" && this.state.IExpenseModel[idx].TravelType == "Leased" && <span className={styles.noteMsg}>(${this.state.StartEndMileCosts}/milage)</span>}</label>
            <input
              placeholder='Amount'
              type="number"
              min="1"
              disabled={item.TravelType == "Leased" ? true : false}
              className={item.Expense == "Travel"?styles.amtTxt:'form-control'}
              name="ExpenseCost"
              value={this.state.IExpenseModel[idx].ExpenseCost}
              onChange={this.handleChange(idx)}
              id={this.state.IExpenseModel[idx].id}
            />
            {item.isExpenseCostError == true &&
              <span className={styles.errMsg}>{item.ExpenseCostMsg}</span>
            }
          </div>
          {this.state.IExpenseModel.length > 1 &&
            <div className="form-group col-md-1">
              <label className="control-label"></label>
              <div id="delete-spec-row" onClick={this.handleRemoveSpecificRow(idx)} className={styles.deleteIcon}>
                <Icon iconName="delete" className="ms-IconExample" />
              </div>
            </div>
          } 

        </div>
        <div className={styles.itemLine}></div>
      </div>)
    })
  }
  public render(): React.ReactElement<INewRequestFormProps> {

    return (
      <div className={styles.newRequestForm}>
        <div className={styles.ml8}>
          <div className={this.state.IExpenseModel.length > 0 ? styles.viewDeptSection : styles.deptSection}>

            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>Department</label><span className={styles.star}>*</span>
              <select className='form-control' name="Department" value={this.state.Department} id="Department" onChange={this.handleChange(1)}>
                <option value="">Select</option>
                {this.renderDropdown(this.state.DepartmentOptions)}
              </select>
              {this.state.IsDeptErr == true && <span className={styles.errMsg}>{this.state.DeptErrMsg}</span>}
            </div>
            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>Description</label><span className={styles.star}>*</span>
              <input
                placeholder='Description'
                type="text"
                className='form-control'
                name="ReportHeader"
                value={this.state.ReportHeader}
                onChange={this.handleChange(2)}
                id="ReportHeader"
                maxLength={254}
              />
              {this.state.IsReportErr == true && <span className={styles.errMsg}>{this.state.ReportHeaderErrMsg}</span>}
            </div>
            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>Start Date &nbsp;<span className={styles.noteMsg}>(MM-DD-YYYY)</span></label>
              <input
                placeholder='Start Date'
                type="date"
                className='form-control'
                name="StartDate"
                value={this.state.StartDate}
                onChange={this.handleChange(3)}
                id="StartDate"
                onKeyDown={(e) => e.preventDefault()}
              />
              {this.state.IsStartDateErr == true && <span className={styles.errMsg}>{this.state.StartDateErrMsg}</span>}
            </div>
            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>End Date&nbsp;<span className={styles.noteMsg}>(MM-DD-YYYY)</span></label>
              <input
                placeholder='End Date'
                type="date"
                className='form-control'
                name="EndDate"
                value={this.state.EndDate}
                onChange={this.handleChange(4)}
                id="EndDate"
                onKeyDown={(e) => e.preventDefault()}
              />
              {this.state.IsEndDateErr == true && <span className={styles.errMsg}>{this.state.EndDateErrMsg}</span>}
            </div>

          </div>

          <table className={styles.newRequestTable}>
            {this.renderTableData()}
          </table>
          <button className='btn btn-primary addExpenseRow' disabled={this.state.IsBtnClicked} id="addDetailRow" onClick={this.handleAddRow}>Add Expense</button>&nbsp;
          {this.state.IsExpenseDetailErr == true && <span className={styles.errMsg}>{this.state.ExpenseDetailErrMsg}</span>}


          <div>
            <div className={styles.line}></div>
          </div>
          <table className={this.state.IExpenseModel.length > 0 ? styles.newRequestTable : styles.newRequestCmtTable}>
            <tr>
              <td className={styles.cmt}>
                <label className="control-label">Comments</label>
                <textarea
                  className='form-control'
                  name="Comments"
                  value={this.state.Comments}
                  onChange={this.handleChange(5)}
                  id="Comments">
                </textarea>
              </td>
              <td className={styles.attachFile} id="inputAttachment">
                <label className="control-label">Attachment(s)</label>
                <input type="file" multiple={true} id="file" onChange={this.addFile.bind(this)} />
                {this.state.isMealExpenseCostError == true && <span className={styles.errMsg}>{this.state.mealExpenseCostErrMsg}</span>}

              </td>
              <td className={styles.attachedFile}>
                {this.state.fileInfos.length > 0 &&
                  <label id="fileName">Attached Files </label>
                }
                {this.renderAttachmentName()}
              </td>
            </tr>
          </table>
          {/* {
          this.state.loading &&
          <Spinner label="Loading items..." size={SpinnerSize.large} />
        } */}
          <div className={this.state.IExpenseModel.length > 0 ? styles.btnSection : ""}>
            <span className={styles.btnRt}>
              <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.SubmitRequest("Submitted")}>Submit</button>  &nbsp;&nbsp;
              <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.SubmitRequest("Draft")}>Save as Draft</button> &nbsp;&nbsp;
              <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={this.RefreshPage}>Cancel</button>
            </span>
          </div>
          <div id="loader" className={styles.loader}></div>
        </div>
      </div>
    );
  }
}
