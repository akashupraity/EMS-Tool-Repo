import * as React from 'react';
import styles from './EditRequestForm.module.scss';
import { IEditRequestFormProps } from './IEditRequestFormProps';
import { IEditRequestFormState } from './IEditRequestFormState';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label, Icon, PrimaryButton } from 'office-ui-fabric-react';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
import { IStyleSet, mergeStyleSets, getTheme, FontWeights, } from 'office-ui-fabric-react/lib/Styling';
import * as $ from "jquery";
import * as bootstrap from "bootstrap";
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IAttachmentInfo, IItem, sp, Web } from '@pnp/sp/presets/all'
import { SPOperations } from '../../SPServices/SPOperations';
import { ListItemAttachments } from "@pnp/spfx-controls-react/lib";
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import {
  DefaultButton,
  Modal,
  IconButton, IButtonStyles, IIconProps, IDetailsRowStyles, DetailsRow
} from 'office-ui-fabric-react';

const cancelIcon: IIconProps = { iconName: 'Cancel' };

const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    // flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '450px',
    //height: '260px',
    color: '#000',
    padding: '10px',
    overflow: 'hidden',
  },
  header: [
    // eslint-disable-next-line deprecation/deprecation
    theme.fonts.xLargePlus,
    {
      flex: '1 1 auto',
      //borderTop: `4px solid ${theme.palette.themePrimary}`,
      color: theme.palette.neutralPrimary,
      display: 'flex',
      alignItems: 'center',
      fontWeight: FontWeights.semibold,
      // padding: '12px 12px 14px 24px',
      fontSize: '20px',
      overflow: 'hidden',
    },
  ],
  body: {
    flex: '4 4 auto',
    padding: '0 24px 24px 24px',
    overflowY: 'hidden',
    selectors: {
      p: { margin: '14px 0' },
      'p:first-child': { marginTop: 0 },
      'p:last-child': { marginBottom: 0 },
    },
  },
})
const iconButtonStyles: Partial<IButtonStyles> = {
  root: {
    color: theme.palette.neutralPrimary,
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};
export default class EditRequestForm extends React.Component<IEditRequestFormProps, IEditRequestFormState> {
  public _spOps: SPOperations;
  constructor(props: IEditRequestFormProps) {
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
      RemoveFiles: [],
      fileInfos: [],
      isMealExpenseCostError: false,
      mealExpenseCostErrMsg: "",
      ExpenseDetailErrMsg: "",
      IsExpenseDetailErr: false,
      DepartmentOptions: [],
      ExpenseTypeOptions: [],
      TravelTypeOption: [],
      FinanceId: null,
      IsFinanceDept: false,
      Status: "",
      FilesToDelete: [],
      ExpenseItemsToDelete: [],
      CurrentUserName: "",
      latestComments: "",
      StartEndMileCosts: null,
      IsBtnClicked: false,
      MealExpense: null,
      openDialog: false,
      openEditDialog: false,
      PaidDate: null,
      IsPaidDateErr: false,
      PaidDateErrMsg: ""
    }

    this._spOps = new SPOperations(this.props.siteUrl);

  }
  //** Convert dates */
  public ConvertDate(dateValue) {
    var d = new Date(dateValue),
      month = '' + (d.getMonth() + 1),
      day = '' + d.getDate(),
      year = d.getFullYear();

    if (month.length < 2) month = '0' + month;
    if (day.length < 2) day = '0' + day;

    return [year, month, day].join('-');
  };

  //**Get all attachement of Item */
  public getAttachments = () => {
    (async () => {
      // get list item by id
      const item: IItem = sp.web.lists.getByTitle(this.props.expenseListTitle).items.getById(this.props.selectedItem.ID);
      // get all attachments
      const attachments: any[] = await item.attachmentFiles();
      console.table(attachments);
      attachments.map((file) => {
        file.name = file.FileName
      })
      this.setState({
        fileInfos: attachments
      })
    })().catch(console.log)
  };

  //**get Selected Expense Item Detail */
  getSelectedExpenseDetail = () => {
    this._spOps.GetListItemByID(this.props.selectedItem.ID, this.props.expenseListTitle).then((result) => {
      this._spOps.GetExpenseDetails(this.props.selectedItem.ID, this.props.expenseDetailListTitle).then((expenseDetails) => {
        this.setState({
          Department: result.Department,
          ReportHeader: result.Title,
          StartDate: result.StartDate != null ? this.ConvertDate(result.StartDate) : null,
          EndDate: result.EndDate != null ? this.ConvertDate(result.EndDate) : null,
          Comments: result.Comments,
          Status: result.Status,
          IExpenseModel: expenseDetails,

        })
      })
    })
    this.getAttachments()
  };

  GetCurrentUserManagerId = () => {
    this._spOps.GetCurrentUser().then((result) => {
      this.setState({
        CurrentUserName: result["Title"]
      });
      this._spOps.getCurrentUserDetails(result["LoginName"]).then((manager) => {
        this._spOps.getManagerDetails(manager).then((managerEmail) => {
          this._spOps.getUserIDByEmail(managerEmail).then((managerId) => {
            this.setState({
              ManagerId: managerId,
            })
          })
        })
      })
    })
  }
  //**Get cogfigurable data from List */
  public GetEMSConfig(result) {
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
      let FinanceUserId = AllFinanceEmail[0].Finance["Id"];
      let financeEmail = AllFinanceEmail[0].Finance["EMail"];
      let startEndMileCost = MileCost[0].Key;
      let mealExpenseCost = MealExpense[0].Key;
      if (financeEmail == result["Email"] && this.props.selectedItem.ManagerStatus == "Approved" && this.props.selectedItem.ReviewForFinace == "Yes") {
        this.setState({
          IsFinanceDept: true
        })
      }
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
  //**Called on Pageload */
  componentDidMount(): void {
    this.GetCurrentUserManagerId();
    this.getSelectedExpenseDetail();
    this._spOps.GetCurrentUser().then((result) => {
      this.GetEMSConfig(result);
    })
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
    //window.location.reload();
    this.setState({
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
      RemoveFiles: [],
      fileInfos: [],
      isMealExpenseCostError: false,
      mealExpenseCostErrMsg: "",
      ExpenseDetailErrMsg: "",
      IsExpenseDetailErr: false,
      DepartmentOptions: [],
      ExpenseTypeOptions: [],
      TravelTypeOption: [],
      FinanceId: null,
      IsFinanceDept: false,
      Status: "",
      FilesToDelete: [],
      ExpenseItemsToDelete: [],
      CurrentUserName: "",
      latestComments: "",
      StartEndMileCosts: null,
      IsBtnClicked: false,
      MealExpense: null,
      openDialog: false,
      openEditDialog: false,
      PaidDate: null,
      IsPaidDateErr: false,
      PaidDateErrMsg: ""
    })
    this.componentDidMount();
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
  // async function with await to save all expenses into "ExpenseDetails" list
  private async AddExpensesDetails(Expenses: any[], expenseItemId: number) {
    let web = Web(this.props.siteUrl);
    let requestorUniqueID = this.GetUniqueRequestorID(expenseItemId);
    for (const expense of Expenses) {
      await web.lists.getByTitle(this.props.expenseDetailListTitle).items.add({
        ExpenseTypes: expense.Expense,
        CheckIn: expense.Checkin != "" ? expense.Checkin : null,
        CheckOut: expense.Checkout != "" ? expense.Checkout : null,
        TravelType: expense.TravelType,
        StartMile: expense.StartMile,
        EndMile: expense.EndMileout,
        Description: expense.Description,
        ExpenseCost: String(expense.ExpenseCost),
        RequestorID: requestorUniqueID,
        ExpensesId: expenseItemId
      });
    }
  }
  // async function with await to delete expenses from "ExpenseDetails" list
  private async DeleteExpensesDetails(Expenses: any[]) {
    let web = Web(this.props.siteUrl);
    for (const expense of Expenses) {
      await web.lists.getByTitle(this.props.expenseDetailListTitle).items.getById(expense.Id).delete();
    }
  }
  // async function with await to update all expenses into "ExpenseDetails" list
  private async UpdateExpenseDetails(Expenses: any[]) {
    let web = Web(this.props.siteUrl);
    for (const expense of Expenses) {
      await web.lists.getByTitle(this.props.expenseDetailListTitle).items.getById(expense.Id).update({
        ExpenseTypes: expense.Expense,
        CheckIn: expense.Checkin != "" ? expense.Checkin : null,
        CheckOut: expense.Checkout != "" ? expense.Checkout : null,
        TravelType: expense.TravelType,
        StartMile: expense.StartMile,
        EndMile: expense.EndMileout,
        Description: expense.Description,
        ExpenseCost: expense.ExpenseCost,
      });
    }
  }
  //**Validation on fields */
  ValidateForm = (submissionType) => {
    var tableArr = this.state.IExpenseModel;
    var isErrExists = false;
    this.setState({ IsDeptErr: false, IsReportErr: false, isMealExpenseCostError: false, IsStartDateErr: false, IsEndDateErr: false });

    if (this.state.Department == null) {
      this.setState({
        DeptErrMsg: "Please select Department",
        IsDeptErr: true
      })
      isErrExists = true;
    }
    if (submissionType == "Submitted" || submissionType == "Draft") {
      if (this.state.ReportHeader == null) {
        this.setState({
          ReportHeaderErrMsg: "Please enter Report Header",
          IsReportErr: true
        })
        isErrExists = true;
      }
      if (this.state.StartDate != null && this.state.EndDate != null) {
        var new_start_date = new Date(this.state.StartDate);
        var new_end_date = new Date(this.state.EndDate);
        if (new_start_date >= new_end_date) {
          this.setState({
            StartDateErrMsg: "Start Date should be less than End Date",
            IsStartDateErr: true
          })
          isErrExists = true;
        }
        if (new_end_date <= new_start_date) {
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
        item.isExpenseCostError=false,
        item.isCheckinError = false;
        item.isCheckoutError = false;
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
        if (item.Expense == "Meal" && item.ExpenseCost >= this.state.MealExpense && this.state.fileInfos.length == 0) {
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
  private isModalOpen(): boolean {
    return this.state.openEditDialog;
  };
  private hideModal = () => {
    this.setState({
      openEditDialog: false
    })
  };
  //* Add expense details in ExpenseDetails list*/
  addUpdateAllExpensesDetails = (itemId, submissionType: string) => {
    if (this.state.IExpenseModel.length > 0) {
      let expenseItemsCreate = [];
      let expenseItemsUpdate = [];
      this.state.IExpenseModel.map((expense) => {
        if (expense.Id != undefined) {
          expenseItemsUpdate.push(expense);
        } else {
          expenseItemsCreate.push(expense);
        }
      })
      if (expenseItemsCreate.length > 0 && expenseItemsUpdate.length > 0) {
        this.AddExpensesDetails(expenseItemsCreate, itemId).then(() => {
          this.UpdateExpenseDetails(expenseItemsUpdate).then(() => {
            $('#loaderEdit').hide();
            alert("Request " + submissionType + " sucessfully");
            this.RefreshPage();
          })

        });
      }
      if (expenseItemsCreate.length > 0 && expenseItemsUpdate.length == 0) {
        this.AddExpensesDetails(expenseItemsCreate, itemId).then(() => {
          $('#loaderEdit').hide();
          alert("Request " + submissionType + " sucessfully");
          this.RefreshPage();
        })
      }
      if (expenseItemsUpdate.length > 0 && expenseItemsCreate.length == 0) {
        this.UpdateExpenseDetails(expenseItemsUpdate).then(() => {
          $('#loaderEdit').hide();
          alert("Request " + submissionType + " sucessfully");
          this.RefreshPage();
          //alert(submissionType=="Submited"?"Request submitted sucessfully":"Request drafted sucessfully");
        })
      }
    }
    else {
      alert("Request " + submissionType + " sucessfully");
      $('#loaderEdit').hide();
      this.RefreshPage();
    }

  }
  //Update Expense
  UpdateExpenseRequest = (submissionType) => {
    var isError = this.ValidateForm(submissionType);
    if (!isError) {
      this.setState({
        IsBtnClicked: true,
      })
      $('#loaderEdit').show();
      let previousComments = this.state.Comments == null ? "" : this.state.Comments;
      let todayDate = this._spOps.GetTodaysDate();
      let latestCommentsHTML = '<strong>' + this.state.CurrentUserName + ' : ' + todayDate + '</strong>' + '<div>' + this.state.latestComments + '</div>';
      let totalAmount: any = 0;
      if (this.state.IExpenseModel.length > 0) {
        this.state.IExpenseModel.map((expense) => {
          totalAmount += expense.ExpenseCost != undefined ? parseInt(expense.ExpenseCost) : 0;
        })
      }
      let updatePostData: any = {};
      updatePostData = {
        Department: this.state.Department,
        Title: this.state.ReportHeader,
        StartDate: this.state.StartDate != "" ? this.state.StartDate : null,
        EndDate: this.state.EndDate != "" ? this.state.EndDate : null,
        FinanceId: this.state.FinanceId,
        TotalExpense: totalAmount.toString(),
      }
      if (this.state.latestComments != "") {
        updatePostData.Comments = latestCommentsHTML.concat(previousComments);
      }
      if (this.props.tabType == "MySubmission") {
        if (submissionType == "Submitted") {
          updatePostData.ManagerStatus = "Pending for Manager";
          updatePostData.Status = "Submitted";
          updatePostData.ManagerId = this.state.ManagerId;
          if (this.props.selectedItem.FinanceStatus == "Clarification") {
            updatePostData.ManagerStatus = "Approved";
            updatePostData.FinanceStatus = "Pending for Finance";
            updatePostData.Status = "InProgress";
          }
        }
        if (submissionType == "Draft") {
          updatePostData.ManagerStatus = "";
          updatePostData.Status = "Draft";
          updatePostData.ManagerId = this.state.ManagerId;
        }
      }
      if (this.props.tabType == "MyTask" && !this.state.IsFinanceDept) {
        updatePostData.ManagerStatus = submissionType;
        if (submissionType == "Rejected") {
          updatePostData.Status = "Rejected";
        }
        if (submissionType == "Approved") {
          updatePostData.ReviewForFinace = "Yes";
          updatePostData.Status = "InProgress";
          updatePostData.FinanceStatus = "Pending for Finance"
        }
      }
      if (this.props.tabType == "MyTask" && this.state.IsFinanceDept) {
        updatePostData.FinanceStatus = submissionType;
        updatePostData.Status = submissionType;
        if (submissionType == "Paid") {
          updatePostData.Status = "Paid";
          updatePostData.AmountPaidDate = this.state.PaidDate;
        }
        if (submissionType == "Rejected") {
          updatePostData.ManagerStatus = "Rejected by Finance";
          updatePostData.Status = "Rejected by Finance";
        }
      }
      this._spOps.UpdateItem(this.props.expenseListTitle, updatePostData, this.props.selectedItem.ID).then((result: any) => {
        console.log(this.props.selectedItem.ID);
        let itemId = this.props.selectedItem.ID;
        let { fileInfos } = this.state;
        let web = Web(this.props.siteUrl);
        if (fileInfos.length > 0) {
          let fileToAttach = [];
          fileInfos.map((fileItem) => {
            if (fileItem.ServerRelativeUrl == undefined) {
              fileToAttach.push(fileItem);
            }
          })

          web.lists.getByTitle(this.props.expenseListTitle).items.getById(itemId).attachmentFiles.addMultiple(fileToAttach).then(() => {
            this.addUpdateAllExpensesDetails(itemId, submissionType);
          });
        }
        else {
          this.addUpdateAllExpensesDetails(itemId, submissionType);
        }
        if (this.state.FilesToDelete.length) {
          web.lists.getByTitle(this.props.expenseListTitle).items.getById(this.props.selectedItem.ID).attachmentFiles.deleteMultiple(...this.state.FilesToDelete);
        }
        if (this.state.ExpenseItemsToDelete.length > 0) {
          this.DeleteExpensesDetails(this.state.ExpenseItemsToDelete);
        }
      })

    }
  }

  //** call validation method and create item into list */
  UpdateRequest = (submissionType) => {
    if (submissionType == "Rejected" && confirm('Are you sure, you want to reject')) {
      this.UpdateExpenseRequest(submissionType);
    } else {
      this.UpdateExpenseRequest(submissionType);
    }
  }

  /* This event will fire on remove specific row */
  handleRemoveSpecificRow = (idx) => () => {
    if (confirm('Are you sure, you want to remove it ?')) {
      try {
        const rows = this.state.IExpenseModel
        let expenseItems = [];
        this.state.IExpenseModel.map((item, index) => {
          if (item.Id != undefined && index === idx) {
            expenseItems.push(item);
          }
        })
        if (rows.length > 1) {
          rows.splice(idx, 1);
        }
        let tempItemsToDeleteArr = this.state.ExpenseItemsToDelete;
        tempItemsToDeleteArr.push.apply(tempItemsToDeleteArr, expenseItems);
        this.setState({ IExpenseModel: rows, ExpenseItemsToDelete: tempItemsToDeleteArr });
      } catch (error) {
        console.log("Error in React Table handle Remove Specific Row : " + error);
      }
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
        TravelType: "",
        StartMile: null,
        EndMileout: null,
        Description: "",
        ExpenseCost: "",
        isExpenseTypeError: false,
        isCheckinError: false,
        isCheckoutError: false,
        ExpenseTypeErrMsg: "",
        isExpenseCostError:false,
        ExpenseCostMsg:"",
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
          latestComments: item.value
        })
      }
      if (item.name == "PaidDate") {
        this.setState({
          PaidDate: item.value
        })
      }
      var rowsArray = this.state.IExpenseModel;
      var newRow = rowsArray.map((row, i) => {
        for (var key in row) {
          if (key == item.name && row.id == item.id) {
            row[key] = item.value;
            if (item.name == "TravelType") {
              if (item.value == "Cab") {
                row['StartMile'] = "";
                row['EndMileout'] = "";
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
        tempArr.push.apply(tempArr, fileInformations);
        this.setState({ fileInfos: tempArr, Attachments: resultFile });
      }
        .bind(this),
      100
    );

    console.log(this.state.Attachments.fileInfos);
  };

  //**Remove specific attachement */
  removeSpecificAttachment = (idx, item) => () => {
    if (confirm('Are you sure, you want to remove it ?')) {
      try {
        const rows = this.state.fileInfos;
        let files = [];
        this.state.fileInfos.map((item, index) => {
          if (item.ServerRelativeUrl != undefined && index === idx) {
            files.push(item.FileName);
          }
        })
        rows.splice(idx, 1);
        let tempFileToDeleteArr = this.state.FilesToDelete;
        tempFileToDeleteArr.push.apply(tempFileToDeleteArr, files);
        console.log("Files to Remove: " + this.state.FilesToDelete);
        this.setState({ fileInfos: rows, FilesToDelete: tempFileToDeleteArr });
      } catch (error) {
        console.log("Error in React Table handle Remove Specific Row : " + error);
      }
    }
  }

  //**render selected attachments */
  renderAttachmentName() {
    let isShowAttachDelete = this.props.tabType != "MyTask" && this.props.selectedItem.formType == "Edit" ? true : false;
    return this.state.fileInfos.map((item, idx) => {
      return (<div><a href={item.ServerRelativeUrl != undefined ? item.ServerRelativeUrl : "javascript:void(0)"} target="_blank">{item.name}</a>&nbsp;&nbsp;
        <span id="delete-spec-row" className={styles.deleteIcon}>
          {isShowAttachDelete &&
            <Icon iconName="delete" className="ms-IconExample" onClick={this.removeSpecificAttachment(idx, item)} />
          }
        </span></div>)
    })
  }
  //** Open Paid PopUp to select paid date*/
  OpenPaidPopUp = () => {
    this.setState({
      openEditDialog: true
    })
  }
  //** Validate Paid Date and save it */
  CheckPaidDate = () => {
    if (this.state.PaidDate == null || this.state.PaidDate == "" || this.state.PaidDate == undefined) {
      this.setState({
        PaidDateErrMsg: "Select Paid Date",
        IsPaidDateErr: true
      })
    }
    else {
      this.UpdateRequest("Paid")
    }
  }
  //**Common function to render Department and Expense Types DropDowns */
  renderDropdown = (options) => {
    return options.map((item, idx) => {
      return (<option value={item.key}>{item.text}</option>)
    })
  }
  //** onClick of Add Expense, Render Expense details fields in table*/
  renderTableData() {
    let isExpenseDeleteIcon = this.props.tabType != "MyTask" && this.props.selectedItem.formType == "Edit" ? true : false;
    return this.state.IExpenseModel.map((item, idx) => {
      let IsExpenseTypeDisable = true;
      if (item.Id != undefined && this.state.Status == "Draft") {
        IsExpenseTypeDisable = false;
      }
      if (item.Id == undefined) {
        IsExpenseTypeDisable = false;
      }
      return (<span key={idx}>
        {/* <div className={styles.expenselbl}>Expense {idx+1}</div> */}
        <div className={this.props.tabType != "MyTask" && this.props.selectedItem.formType == "Edit" ? styles.formRow : styles.editRequestTable}>
          <div className="form-group col-md-2">
            <label className="control-label">Expense Type</label><span className={styles.star}>*</span>
            <select disabled={IsExpenseTypeDisable || this.props.selectedItem.formType == "View" || this.props.tabType == "MyTask"} className='form-control ExpTypeSelectOptions' name="Expense" defaultValue={this.state.IExpenseModel[idx].Expense} value={this.state.IExpenseModel[idx].Expense} id={this.state.IExpenseModel[idx].id} onChange={this.handleChange(idx)}>
              <option value="">Select</option>
              {this.renderDropdown(this.state.ExpenseTypeOptions)}
            </select>
            {item.isExpenseTypeError == true &&
              <span className={styles.errMsg}>{item.ExpenseTypeErrMsg}</span>
            }
          </div>
          {item.Expense == "Hotel" &&
            <div className="form-group col-md-3">
              <label className="control-label">Check In</label><span className={styles.star}>*</span>
              <input
                placeholder='check in'
                type="date"
                disabled={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false}
                className='form-control'
                name="Checkin"
                defaultValue={this.state.IExpenseModel[idx].Checkin}
                value={this.state.IExpenseModel[idx].Checkin}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />

              {item.isCheckinError == true && <span className={styles.errMsg}>{item.CheckinErrMsg}</span>}
            </div>
          }
          {item.Expense == "Hotel" &&
            <div className="form-group col-md-3">
              <label className="control-label">Check Out</label><span className={styles.star}>*</span>
              <input
                placeholder='Check out'
                type="date"
                disabled={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false}
                className='form-control'
                name="Checkout"
                value={this.state.IExpenseModel[idx].Checkout}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
              {item.isCheckoutError == true && <span className={styles.errMsg}>{item.CheckoutErrMsg}</span>}
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-2">
              <label className="control-label">Travel Type</label>
              <select disabled={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false} className='form-control ExpTypeSelectOptions' name="TravelType" defaultValue={this.state.IExpenseModel[idx].TravelType} value={this.state.IExpenseModel[idx].TravelType} id={this.state.IExpenseModel[idx].id} onChange={this.handleChange(idx)}>
                <option value="">Select</option>
                {this.renderDropdown(this.state.TravelTypeOption)}
              </select>
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-2">
              <label className="control-label">Start Mile</label>
              <input
                placeholder='Start Mile'
                type="number"
                min="1"
                disabled={this.props.tabType == "MyTask" || item.TravelType == "Cab" || this.props.selectedItem.formType == "View" ? true : false}
                className='form-control'
                name="StartMile"
                value={this.state.IExpenseModel[idx].StartMile}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-2">
              <label className="control-label">End Mile</label>
              <input
                placeholder='End Mile'
                type="number"
                min="1"
                disabled={this.props.tabType == "MyTask" || item.TravelType == "Cab" || this.props.selectedItem.formType == "View" ? true : false}
                className='form-control'
                name="EndMileout"
                value={this.state.IExpenseModel[idx].EndMileout}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
              {item.isEndMileError && <span className={styles.errMsg}>{item.EndMileErrMsg}</span>}
            </div>
          }
          {(item.Expense == "Meal" || item.Expense == "Others") &&
            <div className="form-group col-md-6">
              <label className="control-label">Description</label>
              <input
                placeholder='Description'
                type="text"
                disabled={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false}
                className='form-control'
                name="Description"
                value={this.state.IExpenseModel[idx].Description}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
            </div>
          }
          <div className="form-group col-md-3">
            <label className="control-label">Amount ($)<span className={styles.star}>*</span> &nbsp;</label>{item.Expense == "Travel" && this.state.IExpenseModel[idx].TravelType == "Leased" && <span className={styles.noteMsg}>(${this.state.StartEndMileCosts}/milage)</span>}
            <input
              placeholder='Expense Cost'
              type="number"
              min="1"
              disabled={item.TravelType == "Leased" || this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false}
              className='form-control'
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
              <div id="delete-spec-row" className={styles.deleteIcon}>
                {isExpenseDeleteIcon &&
                  <Icon iconName="delete" className="ms-IconExample" onClick={this.handleRemoveSpecificRow(idx)} />
                }
              </div>
            </div>
          }
        </div>

      </span>)
    })
  }
  public render(): React.ReactElement<IEditRequestFormProps> {
    let isShowAddExpenseBtn = this.props.tabType != "MyTask" && this.props.selectedItem.formType == "Edit" ? true : false;
    return (
      <div className={styles.editRequestForm}>
        <div className={styles.ml8}>
          <div className={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? styles.viewDeptSection : styles.deptSection}>
            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>Department</label><span className={styles.star}>*</span>
              <select disabled={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false} className='form-control' name="Department" value={this.state.Department} id="Department" onChange={this.handleChange(1)}>
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
                disabled={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false}
                className='form-control'
                name="ReportHeader"
                value={this.state.ReportHeader}
                onChange={this.handleChange(2)}
                id="ReportHeader"
              />
              {this.state.IsReportErr == true && <span className={styles.errMsg}>{this.state.ReportHeaderErrMsg}</span>}
            </div>
            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>Start Date</label>
              <input
                placeholder='Start Date'
                type="date"
                disabled={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false}
                className='form-control'
                name="StartDate"
                value={this.state.StartDate}
                onChange={this.handleChange(3)}
                id="StartDate"
              />
              {this.state.IsStartDateErr == true && <span className={styles.errMsg}>{this.state.StartDateErrMsg}</span>}
            </div>
            <div className="form-group col-md-3">
              <label className={styles.lblCtrl}>End Date</label>
              <input
                placeholder='End Date'
                type="date"
                disabled={this.props.tabType == "MyTask" || this.props.selectedItem.formType == "View" ? true : false}
                className='form-control'
                name="EndDate"
                value={this.state.EndDate}
                onChange={this.handleChange(4)}
                id="EndDate"
              />
              {this.state.IsEndDateErr == true && <span className={styles.errMsg}>{this.state.EndDateErrMsg}</span>}
            </div>


          </div>
          {this.renderTableData()}
          {isShowAddExpenseBtn &&
            <div className='btn btn-primary' id="add-row" onClick={this.handleAddRow}>Add Expense</div>
          }

          {this.state.IsExpenseDetailErr == true && <span className={styles.errMsg}>{this.state.ExpenseDetailErrMsg}</span>}
          <div>
            <div className={styles.line}></div>
          </div>
          <div className="form-row">
            <div className="form-group col-md-4">

              <label className="control-label">Comments</label>
              <textarea
                className='form-control'
                name="Comments"
                cols={6}
                rows={3}
                disabled={this.props.selectedItem.Status == "Approved" || this.props.selectedItem.formType == "View" ? true : false}
                value={this.state.latestComments}
                onChange={this.handleChange(5)}
                id="Comments">
              </textarea>
            </div>
            {this.props.tabType != "MyTask" && this.props.selectedItem.formType == "Edit" &&
              <div className="form-group col-md-3">
                <label className="control-label">Attachements</label>
                <input type="file" disabled={this.props.selectedItem.formType == "View" ? true : false} multiple={true} id="file" onChange={this.addFile.bind(this)} />
                {this.state.isMealExpenseCostError == true && <span className={styles.errMsg}>{this.state.mealExpenseCostErrMsg}</span>}
              </div>
            }
            <div className="form-group col-md-5">
              {this.state.fileInfos.length > 0 &&
                <label id="fileName">Attached Files </label>
              }
              {this.renderAttachmentName()}

            </div>
          </div>
          <div className="form-row">
            <div className="form-group col-md-6">
              {this.state.Comments != null &&
                <span> <label className="control-label">Comments History</label>
                  <div className={styles.commentContainer}>
                    <div className={styles.cmtHistoryRow}>
                      <div className={styles.comment}>
                        <div dangerouslySetInnerHTML={{ __html: this.state.Comments }}></div>
                      </div>
                    </div>
                  </div>
                </span>
              }
            </div>
          </div>
          <div>
            <span className={styles.btnRt}>
              {this.props.selectedItem.Status != "Approved" && <span>
                {this.props.tabType == "MySubmission" && this.props.selectedItem.formType == "Edit" &&
                  <span><button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Submitted")}>Submit</button>  &nbsp;&nbsp;
                    <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Draft")}>Save as Draft</button> &nbsp;&nbsp;
                  </span>
                }
                {this.props.tabType == "MyTask" && this.state.IsFinanceDept == false && this.props.selectedItem.formType == "Edit" &&
                  <span><button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Approved")}>Approve</button>  &nbsp;&nbsp;
                    <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Rejected")}>Reject</button> &nbsp;&nbsp;
                  </span>
                }
                {this.props.tabType == "MyTask" && this.state.IsFinanceDept == true && this.props.selectedItem.formType == "Edit" &&
                  <span><button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.OpenPaidPopUp()}>Paid</button>  &nbsp;&nbsp;
                    <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Clarification")}>Clarification</button> &nbsp;&nbsp;
                    <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Rejected")}>Reject</button> &nbsp;&nbsp;

                  </span>
                }
              </span>
              }

              <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={this.RefreshPage}>Close</button>
            </span>
            <div id="loaderEdit" className={styles.loaderEdit}></div>
          </div>
        </div>
        <Modal isOpen={this.isModalOpen()} isBlocking={false} containerClassName={contentStyles.container}>
          <div className={contentStyles.body}>
            <div className={contentStyles.header}>
              <span className={styles.label}> Select Paid Date</span>
              <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideModal} />
            </div>

            <div className="form-row">
              <div className="form-group col-md-12">
                {/* <label className="control-label">Paid Date</label> */}
                <input
                  placeholder='Paid Date'
                  type="date"
                  className='form-control'
                  name="PaidDate"
                  value={this.state.PaidDate}
                  onChange={this.handleChange(5)}
                  id="PaidDate"
                />
                {this.state.IsPaidDateErr == true && <span style={{ 'color': 'Red' }}>{this.state.PaidDateErrMsg}</span>}
              </div>
            </div>

            <div className="form-row">
              <div className="form-group col-md-12">
                <button className='btn btn-primary float-right' id="add-row" onClick={() => this.CheckPaidDate()}>Ok</button>
              </div>
            </div>
          </div>
        </Modal>
      </div>
    );
  }
}
