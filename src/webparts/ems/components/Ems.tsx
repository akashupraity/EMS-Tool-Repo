import * as React from 'react';
import styles from './Ems.module.scss';
import { IEmsProps } from './IEmsProps';
import { IEmsState } from './IEmsState';
import { escape } from '@microsoft/sp-lodash-subset';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DateTimePicker, DateConvention, TimeConvention, TimeDisplayControlType } from '@pnp/spfx-controls-react/lib/dateTimePicker';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Stack, IStackProps, IStackStyles } from 'office-ui-fabric-react/lib/Stack';
//import * as jsPDF from 'jspdf';  
import { jsPDF } from "jspdf";
import html2canvas from 'html2canvas';
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IAttachmentInfo, IItem, Web } from '@pnp/sp/presets/all'
import { SPComponentLoader } from '@microsoft/sp-loader';
import { sp } from "@pnp/sp";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { MSGraphClient } from '@microsoft/sp-http';
import { RichText } from "@pnp/spfx-controls-react/lib/RichText";
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');

import { IPivotItemProps, Pivot, PivotItem, PivotLinkFormat } from 'office-ui-fabric-react/lib/Pivot';
import { Label, ILabelStyles } from 'office-ui-fabric-react/lib/Label';
import { IStyleSet, mergeStyleSets, getTheme, FontWeights, } from 'office-ui-fabric-react/lib/Styling';
import NewRequestForm from './../../newRequestForm/components/NewRequestForm'
import { SPOperations } from '../../SPServices/SPOperations';
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import * as $ from "jquery";
import * as bootstrap from "bootstrap";
require('../../../../node_modules/bootstrap/dist/css/bootstrap.min.css');
import {
  DefaultButton,
  Modal,
  IconButton, IButtonStyles, IIconProps, IDetailsRowStyles, DetailsRow, Icon
} from 'office-ui-fabric-react';
//import EditRequestForm from '../../editRequestForm/components/EditRequestForm';
//import { ListItemAttachments } from '@pnp/spfx-controls-react';
//import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/controls/peoplepicker';
import * as moment from 'moment'
//import { BaseWebPartContext, WebPartContext } from '@microsoft/sp-webpart-base';

//const cancelIcon: IIconProps = { iconName: 'Cancel' };

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
  }
})
const contentStatusBarStyles = mergeStyleSets({
  container: {
    display: 'flex',
    // flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '1060px',
    height: '250px',
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
  label: {
    'font-family': 'inherit',
    'font-weight': '600',
  },
  'h5':{
    'margin-left': '-12px',
    'margin-top': '15px',
    'font-size': 'small',
  },
  'h4':{
    'margin-left': '-13px',
    'font-size': '16px',
    'margin-top': '-20px',
    'color': 'blue',
    'width': 'max-content',
  },
  'mailIcon':{
    'padding-left': '36px',
  },
  'progressbar': {
    'width': '809px',
    'height': '9px',
    'background-color': 'lightgray',
    'border-radius': '10px',
    'display': 'flex',
    'align-items': 'center',
    'justify-content': 'space-between',
    'padding': '0px',
    'margin-top': '5%',
    'margin-left': '5%',
},
'Paid':{
  'margin-left': '18px',
},
'statuscircle': {
  'width': '40px', /* Adjust the size as per your requirement */
  'height': '40px', /* Adjust the size as per your requirement */
  'border-radius': '50%',
  'background-color': 'lightgray', /* Default color */
  'transition': 'background-color 0.3s',
},
'approved': {
  'background-color': 'green', /* Change color when approved */
},
'rejected': {
  'background-color': 'red', /* Change color when rejected */
},
'pending': {
  'background-color': '#ffb100', /* Change color when pending */
},
'clarify':{
  'background-color': '#09cceb', /* Change color when clarification */
},
'icon':{
  'margin-top': '12px',
    'margin-left': '12px',
    'width': '10px',
    'height': '10px',
    'color': 'white',
}
})


const contentInvoiceStyles = mergeStyleSets({
  container: {
    display: 'flex',
    // flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '1060px',
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
  'invoiceNum':{
    'margin-top': '60px',
    'margin-left': '60px',
  },
  'table' : {
    'border-collapse': 'collapse',
    'width': '90%',
    'margin-top': '0px',
    'border': '2px solid #787878',
    'align-items': 'center',
  },
  'table2': {
    'border-collapse': 'collapse',
    'width': '90%',
    'margin-top': '-10px',
    'border': '2px solid #787878',
  },
  'th, td' : {
    'padding': '8px',
    'text-align': 'left',
    'border-bottom': '1px solid #ddd',
    'font-weight': '100',
    'font-size': 'small',
  },
  'td' : {
    'padding': '8px',
    'text-align': 'left',
    'border-bottom': '1px solid #ddd',
    'font-weight': '100',
    'font-size': 'small',
  },
  'th': {
    'background-color': 'white',
  },
  h1 : {
    'text-align': 'center',
  },
  'invoiceinfo': {
    'margin-bottom': '20px',
  },
  'invoiceinfo p': {
    'margin': '0'
  },
  'invoiceinfo p strong' :{
    'margin-right': '10px',
  },
  'total': {
    'font-weight': 'bold',
  },
  'Container':{
    'width': '90%',
    'height': '35px',
    'color': 'white',
    'background-color': '#484848',
    'text-align': 'left',
    'border-radius': '0px',
    'font-size': '15px',
  },
  'Container2':{
    'width': '90%',
    'height': '45px',
    'color': 'white',
    'background-color': '#484848',
    'text-align': 'left',
    'border-radius': '1px',
    'font-size': '15px',
  },
  h3:{
    'padding-left': '20px',
    'padding-top': '10px',
  },
  'Containerbox': {
    'width': '100%',
    'margin': '0 auto', /* Center the container horizontally */
    'padding': '20px',
    'border': '1px solid #ccc',
    'padding-right': '-5%',
    'padding-left': '80px',
  },
  'footeritem': {
    'text-align':'right',
   'padding-right': '10%',
   'padding-bottom': '5%',
		//'margin-top': '-10%',
  }
})
const iconButtonEditFormStyles: Partial<IButtonStyles> = {
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


const liTheme = getTheme();
const labelStyles: Partial<ILabelStyles> = {
  root: { marginTop: 10 },
};
const cancelIcon: IIconProps = { iconName: 'Cancel' };

//const theme = getTheme();
const contentEditFormStyles = mergeStyleSets({
  container: {
    display: 'flex',
    // flexFlow: 'column nowrap',
    alignItems: 'stretch',
    width: '1200px',
    height: '800px',
    color: '#000',
    padding: '10px',
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
      //  padding: '12px 12px 14px 24px',
      fontSize: '20px',
    },
  ],
  viewInvoice:{
    'margin-bottom': '7px',
  },
  loaderEdit: {
    'display': 'none',
    'position': 'fixed',
    'top': '0',
    'left': '0',
    'right': '0',
    'bottom': '0',
    'width': '100%',
    'background': "rgba(0,0,0,0.75) url('https://bbsmidwestcom.sharepoint.com/sites/HR/SiteAssets/ICONS/loading2.gif') no-repeat center center",
    'z-index': '10000',
  },
  statusImg:{
    'width':'18px',
  },
  commentContainer: {
    'margin': '0px auto',
    'background': '#f5f4f7',
    'border-radius': '8px',
    'padding': '14px',
  },
  cmtHistoryRow: {
    //'display': '-ms-flexbox',
    'display': 'flex',
    '-ms-flex-wrap': 'wrap',
    'flex-wrap': 'wrap',
    'margin-right': '-13px',
    'margin-left': '-5px',
    'overflow-y': 'auto',
    'max-height': '120px',
  },
  comment: {
    'display': 'block',
    'transition': 'all 1s',
  },
  label: {
    'font-family': 'inherit',
    'font-weight': '600',
  },
  lblClr: {
    'font-size': '15px',
    'color': '#007bff',
    'font-family': 'inherit',
    'font-weight': '600',
    'width': '80px'
  },
  newRequestTable: {
    'width': '100%',
  },
  deptSection: {
    'background-color': '#f7f7f7',
    //  'display': '-ms-flexbox',
    'display': 'flex !important',
    '-ms-flex-wrap': 'wrap',
    'flex-wrap': 'wrap',
    'margin-right': '-5px',
    'margin-left': '-5px',
    'margin-bottom': '10px',
    'margin-top': '5px',
  },
  viewDeptSection: {
    'background-color': '#f7f7f7',
    'display': 'flex !important',
    'flex-wrap': 'wrap',
    'margin-right': '-5px',
    'margin-left': '-5px',
    'margin-bottom': '15px',
  },
  editRequestTable: {
    // 'display': '-ms-flexbox',
    'display': 'flex',
    '-ms-flex-wrap': 'wrap',
    'flex-wrap': 'wrap',
    'margin-right': '-5px',
    'margin-left': '-5px',
    'margin-top': '-10px',
    'background-color': '#fdfdfd',
  },
  line: {
    'width': '100%',
    'border-top': '1px solid rgba(0,0,0,.1)',
    'display': 'inline-block',
  },
  noteMsg: {
    'font-family': 'inherit',
    'font-weight': '400',
    'font-size': '13px',
    'color': '#a5a3a3',
  },
  formRow: {
    //'display': '-ms-flexbox',
    'display': 'flex',
    '-ms-flex-wrap': 'wrap',
    'flex-wrap': 'wrap',
    'margin-right': '-5px',
    'margin-left': '-5px',
    'margin-bottom': '-10px',
    'background-color': '#fdfdfd',
  },
  lblCtrl: {
    'margin-top': '5px',
    'font-family': 'inherit',
    'font-weight': '600',
    'width': '170px'
  },
  pr15: {
    'padding-right': '15px',
  },
  btnbr9: {
    'border-radius': '9px',
    'color': '#fff',
    'background-color': '#007bff',
    'border-color': '#007bff',
  },
  ml8: {
    'margin-left': '8px',
  },


  star: {
    'color': 'red',
  },
  errMsg: {
    'color': 'red',
  },
  deleteIcon: {
    'color': 'red',
    'margin-top': '15px',
  },
  btnRt: {
    float: 'left',
    'margin-right': '15px',
    'margin-bottom': '15px',
  },


  inputAttachment: {
    'padding-bottom': '25px',
  },
  expenselbl: {
    'font-weight': '600',
    'color': 'white',
    'font-size': 'initial',
    'background-color': '#6492c3',
    'padding-left': '5px',
  },
  cmt: {
    'width': '300px',
    'padding-right': '15px',
  },
  attachFile: {
    'width': '150px',
    'padding-right': '15px',
  },
  attachedFile: {
    'width': '500px',
    'padding-right': '15px',
  },
  formControl: {
    'display': 'block',
    'width': '130%',
    'height': 'calc(1.5em + .75rem + 2px)',
    'padding': '.375rem .75rem',
    'font-size': '1rem',
    'font-weight': '400',
    'line-height': '1.5',
    'color': '#495057',
    'background-color': '#fff',
    'background-clip': 'padding-box',
    'border': '1px solid #ced4da',
    'border-radius': '.25rem',
    'transition': 'border-color .15s ease-in-out,box-shadow .15s ease-in-out'
  },
  itemLine: {
    'width': '100%',
    'border-top': '1px solid rgba(0,0,0,.1)',
    'display': 'inline-block',
    'margin-bottom': '15px',
  }
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
export default class Ems extends React.Component<IEmsProps, IEmsState> {
  public _spOps: SPOperations;
  constructor(props: IEmsProps) {
    super(props)
    sp.setup({
      spfxContext: this.props.context
    });
    // import third party css file from cdn
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css');
    this.state = {
      MySubmissionItems: [],
      MyPendingItems: [],
      MyApprovedItems: [],
      MyRejectedItems: [],
      MyPaidItems: [],
      openDialog: false,
      selectedExpense: {},
      SelectedTabType: "",
      IsManager: false,
      IsFinanceDept: false,

      // Edit form 
      IExpenseModel: [],
      Department: "",
      ReportHeader: "",
      StartDate: null,
      EndDate: null,
      Creator: "",
      Manager: "",
      TotalExpense: "",
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
      //IsFinanceDept:false,
      Status: "",
      FilesToDelete: [],
      ExpenseItemsToDelete: [],
      CurrentUserName: "",
      latestComments: "",
      StartEndMileCosts: null,
      IsBtnClicked: false,
      MealExpense: null,
      //openDialog:false,
      openEditDialog: false,
      openFinanceDialog:false,
      openStatusBarDialog:false,
      PaidDate: null,
      IsPaidDateErr: false,
      PaidDateErrMsg: "",
      openInvoiceDialog: false,
      CreatorEmail:"",
      ILogHistoryModel:[],
      ManagerStatus:"",
      FinanceStatus:"",
      Finance:"",
      RequestorResponse:"",
      ManagerResponse:"",
      FinanceResponse:"",
      NewFinanceUserID:null,
      FinanceEmailID:[],
      FinanceItemId:null,
      CurrentUserEmail:"",
      OtherFinanceEmailID:"",
    }

    this._spOps = new SPOperations(this.props.siteUrl);
  };
  public GetMySubmittedItems(result) {
    this._spOps.getListItems(result["Email"], "MySubmission", "Owner").then((response) => {
      this.setState({
        MySubmissionItems: response
      })
    })
  };
  public itemExists(item, arr) {
    let isExists = false;
    arr.map((value) => {
      if (value.ID == item.ID) {
        isExists = true;
        return false;
      }
    });
    return isExists;

  }
  public getUniqueRequests(responses) {
    let uniqueResponses = [];
    responses.map((item) => {
      if (item.ManagerStatus === "Approved" && item.FinanceStatus === "Approved" && this.itemExists(item, uniqueResponses)) {
        return false;
      }

      uniqueResponses.push(item);
      return true;
    })
    return uniqueResponses;
  }

  public getMasterDetails(result) {
    this._spOps.getEMSConfigListItems().then((response: any) => {
      let AllFinanceEmail = [];
      let AllDepartment = [];
      let AllExpenseType = [];
      let AllTravelType = [];
      let MileCost = [];
      let MealExpense = [];
      let financeItemId=null;
      response.map((item) => {
        if (item.Title == "FinanceDepartment") {
          financeItemId=item.ID
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
      let financeEmail = AllFinanceEmail[0].Finance["EMail"];
      let otherFinanceEmail = AllFinanceEmail[0].OtherFinance!=undefined?AllFinanceEmail[0].OtherFinance["EMail"]:"";
      let financeEmailArr=[];
      financeEmailArr.push(financeEmail);
      let startEndMileCost = MileCost[0].Key;
      let mealExpenseCost = MealExpense[0].Key;
      if (financeEmail == result["Email"] && this.state.selectedExpense.ManagerStatus == "Approved" && this.state.selectedExpense.ReviewForFinace == "Yes") {
        this.setState({
          IsFinanceDept: true
        })
      }
      this.setState({
        DepartmentOptions: AllDepartment,
        ExpenseTypeOptions: AllExpenseType,
        FinanceId: FinanceUserId,
        FinanceEmailID:financeEmailArr,
        OtherFinanceEmailID:otherFinanceEmail,
        FinanceItemId:financeItemId,
        TravelTypeOption: AllTravelType,
        StartEndMileCosts: startEndMileCost,
        MealExpense: parseInt(mealExpenseCost)
      })
    })
  }
  public GetEMSConfig(result) {
    this._spOps.getEMSConfigListItems().then((response: any) => {

      let AllFinanceEmail = [];
      response.map((item) => {
        if (item.Title == "FinanceDepartment") {
          AllFinanceEmail.push(item);
        }
      })
      let financeEmail = AllFinanceEmail[0].Finance["EMail"];
      let pendingItems = [];
      let rejectedItems = [];
      let approvedItems = [];
      let paidItems = [];
      if (financeEmail == result["Email"]) { //if manager and finance are same person
        this._spOps.getListItems(result["Email"], "MyTask", "Finance").then((financeResponse) => {
          this._spOps.getListItems(result["Email"], "MyTask", "Manager").then((managerResponse) => {
            let responses = [...financeResponse, ...managerResponse];
            let requests = this.getUniqueRequests(responses);
            requests.map((item) => {
              if ((item.Status == "Approved")) {
                approvedItems.push(item);
              }
              if (item.Status == "Paid") {
                paidItems.push(item);
              }
              if (item.Status == "Rejected") {
                rejectedItems.push(item);
              }
              if (item.Status == "Pending for Manager" || item.Status == "Pending for Finance" || item.Status == "Clarification") {
                pendingItems.push(item);
              }
            })
            this.setState({
              MyPendingItems: pendingItems,
              MyRejectedItems: rejectedItems,
              MyApprovedItems: approvedItems,
              MyPaidItems: paidItems,
              IsFinanceDept: true
            })
          })
        })
        // this._spOps.getEMSConfigListItems().then((response:any)=>{
        this.getMasterDetails(result);

        // })
      }
      else {
        this._spOps.getListItems(result["Email"], "MyTask", "Manager").then((response) => {
          response.map((item) => {
            if (item.Status == "Approved") {
              approvedItems.push(item);
            }
            if (item.Status == "Rejected") {
              rejectedItems.push(item);
            }
            if (item.ManagerStatus == "Pending for Manager") {
              pendingItems.push(item);
            }
          })
          this.setState({
            MyPendingItems: pendingItems,
            MyRejectedItems: rejectedItems,
            MyApprovedItems: approvedItems,
            IsManager: response.length > 0 ? true : false,
          })
        })
        this.getMasterDetails(result);
      }
      console.log("config " + response);
    })
  }
  



  componentDidMount(): void {
    this._spOps.GetCurrentUser().then((result) => {
      this.GetMySubmittedItems(result);
      this.GetEMSConfig(result);
      //this.getAllO365Users();
      this.GetCurrentUserManagerId();
      this.setState({
        CurrentUserEmail:result["Email"],
      })
    })
  
    // this._spOps.GetCurrentUser().then((result)=>{
    // this.GetEMSConfig(result);
    // })
  }
  public selectedTabType(tabType: string) {
    this.setState({
      SelectedTabType: tabType
    })
  };

  public OpenEditForm(formType: string, selectedItem) {
    if (formType == "EditMySubmission") {
      selectedItem.formType = "Edit";
    }
    else {
      selectedItem.formType = "View";
    }
    setTimeout(
      function () {
        this.setState({ openEditDialog: true, selectedExpense: selectedItem })
        this.getSelectedExpenseDetail(selectedItem);
      }
        .bind(this),
      100
    );
  };
  // public updateStatusCircles(statusCircles,currentStatus){
  //   statusCircles.forEach((circle, index) => {
  //     if(index==3){
  //      circle.classList.add('active1');
  //     }
  //       if (index <= currentStatus) {
  //         circle.classList.add('active');
  //       } else {
  //         circle.classList.remove('active');
  //       }
  //     });
  // }
  public OpenStatusBarPopUp(selectedItem){
      setTimeout(
      function () {
        this.setState({ openStatusBarDialog: true, selectedExpense: selectedItem })
        const statusCircles = document.querySelectorAll('#statuscircleId');
        //this.updateStatusCircles(statusCircles[0].childNodes,2);
        this.getSelectedExpenseDetail(selectedItem);
      }
        .bind(this),
      100
    );
  }
  OpenViewFrom = (selectedItem) => {
    setTimeout(
      function () {
        selectedItem.formType = "View";
        this.setState({ openEditDialog: true, selectedExpense: selectedItem })
        this.getSelectedExpenseDetail(selectedItem);
      }
        .bind(this),
      100
    );

  }
  OpenFinanceModal = ()=>{
    this.setState({
      openFinanceDialog:true
    })
  }
  private hideModal = () => {
    this.setState({
      openDialog: false
    })
  };
  //hide Invoice popup
  private hideInvoiceModal = () => {
    this.setState({
      openInvoiceDialog: false
    })
  };
  /** Render color with rows */
  private OnListViewRenderRow(props: any) {
    const customStyles: Partial<IDetailsRowStyles> = {};
    if (props) {
      if (props.itemIndex % 2 === 0) {
        //Every other row render with different background
        customStyles.root = { backgroundColor: liTheme.palette.themeLighterAlt };
      }
      return <DetailsRow {...props} styles={customStyles} />
    }
    return null;
  };

  private onPivotItemClick = (item: PivotItem) => {
    if (item.props.itemKey) {
      this.setState({
        SelectedTabType: item.props.itemKey
      })
      this.componentDidMount();
    }
  }
  public viewFields() {
    const viewFields: IViewField[] = [{
      name: this.state.SelectedTabType == "MySubmission" ? "Edit" : this.state.SelectedTabType == "MyTask" ? "Action" : "",
      displayName: "",
      minWidth: 45,
      maxWidth: 45,
      render: (item: any) => {
        let isEditBtnDisable = false;
        let button;
        if (this.state.SelectedTabType == "MySubmission") {
          isEditBtnDisable = item.Status == "Rejected by Finance" || item.Status == "Rejected" || item.Status == "Draft" || item.Status == "Clarification" ? false : true;
          button = <button disabled={isEditBtnDisable} type="button"  id="add-rowEdit" onClick={() => this.OpenEditForm("EditMySubmission", item)}><i className="fa fa-pencil-square-o" title='Edit'></i></button>
          //button    = <i className="fa fa-pencil-square-o" title='Edit'  id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}></i>
    
        }
        if (this.state.SelectedTabType == "MyTask") {
          button = <button disabled={isEditBtnDisable} type="button" id="add-rowAction" onClick={() => this.OpenEditForm("EditMySubmission", item)}><i className="fa fa-tasks" title='Action' ></i></button>
         // button    = <i className="fa fa-tasks" title='Take Action' id="add-row9" onClick={() => this.OpenEditForm("EditMySubmission", item)}></i>
        
        }
        return <span>
          {button}
        </span>;
      }
    },
    {
      name: "View",
      displayName: "",
      minWidth: 35,
      maxWidth: 35,
      render: (item: any) => {
        //return <div className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("ViewMySubmission", item)}>View</div>;
       // return  <i className="fa fa-eye" title='View'  id="add-row" onClick={() => this.OpenEditForm("ViewMySubmission", item)}></i>
        return   <button  type="button"  id="add-rowProgress" onClick={() => this.OpenEditForm("ViewMySubmission", item)}><i className="fa fa-eye" title='View' ></i></button>
      }
    },
    {
      name: this.state.SelectedTabType == "MySubmission" ? "Progress" : "",
      displayName: "",
      minWidth: 60,
      maxWidth: 60,
      render: (item: any) => {
        let statusIcon;
        // if (this.state.SelectedTabType == "MySubmission") {
        //   statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><i className="fa fa-spinner" title='Show Progress' ></i></button>//<i className="fa fa-spinner" onClick={() => this.OpenStatusBarPopUp(item)} title='Show Progress'></i> 
        // }
        if ((item.Status=="Rejected" || item.Status=="Rejected by Finance")) {
          
          if (this.state.SelectedTabType == "MySubmission") {
           statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/decline.png'} /></button>
       }
         }
          if ((item.Status=="InProgress" || item.Status=="Approved" || item.Status=="Submitted")) {
           if (this.state.SelectedTabType == "MySubmission") {
           statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/timer.png'} /></button>
         }
         }
          if (item.Status=="Paid") {
           if (this.state.SelectedTabType == "MySubmission") {
           statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/checked.png'} /></button> 
         }
         }
          if (item.Status=="Clarification") {
           if (this.state.SelectedTabType == "MySubmission") {
           statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/clear.png'} /></button> 
         }
         }
         if (item.Status=="Draft") {
          if (this.state.SelectedTabType == "MySubmission") {
          statusIcon = <button  type="button"  id="add-row" onClick={() => this.OpenStatusBarPopUp(item)}><img title={item.Status} className={contentEditFormStyles.statusImg} src={this.props.siteUrl + '/SiteAssets/ICONS/notepad.png'} /></button> 
        }
        }
        return <span>
          {statusIcon}
        </span>;
      }
    },
      {
      name: "Department",
      displayName: "Department",
      isResizable: true,
      sorting: true,
      minWidth: 150,
      maxWidth: 150,
      render: (item: any) => {
        return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Department}</a>;
      }
    },
    {
      name: "RequestorID",
      displayName: "Request ID",
      isResizable: true,
      sorting: true,
      minWidth: 130,
      maxWidth: 130,
      render: (item: any) => {
        return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.RequestorID}</a>;
      }
    },
    // {
    //   name: "ReportHeader",
    //   displayName: "Description",
    //   isResizable: true,
    //   sorting: true,
    //   minWidth: 130,
    //   maxWidth: 200,
    //   render: (item: any) => {
    //     return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.ReportHeader}</a>;
    //   }
    // },

    // {
    //   name: "StartDate",
    //   displayName: "Start Date",
    //   isResizable: true,
    //   sorting: true,
    //   minWidth: 100,
    //   maxWidth: 140,
    //   render: (item: any) => {
    //     return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.StartDate}</a>;
    //   }
    // },
    // {
    //   name: "EndDate",
    //   displayName: "End Date",
    //   isResizable: true,
    //   sorting: true,
    //   minWidth: 100,
    //   maxWidth: 140,
    //   render: (item: any) => {
    //     return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.EndDate}</a>;
    //   }
    // },
    
    {
      name: "TotalExpense",
      displayName: "Amount($)",
      isResizable: true,
      sorting: true,
      minWidth: 130,
      maxWidth: 130,
      render: (item: any) => {
        return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.TotalExpense}</a>;
      }
    },
    {
      name: "Status",
      displayName: "Status",
      isResizable: true,
      sorting: true,
      minWidth: 200,
      maxWidth: 200,
      render: (item: any) => {
        let statusHtml;
        if (item.Status == "Paid") {
          statusHtml = <span><a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Status}</a><div>({item.AmountPaidDate?moment(item.AmountPaidDate,'DD/MM/YYYY').format('MM/DD/YYYY'):""})</div></span>
        } else {
          statusHtml = <span><a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Status}</a></span>;
        }
        return <span>
          {statusHtml}
        </span>;
      }
    },
    {
      name: "Submitted By",
      displayName: "",
      minWidth: 130,
      maxWidth: 130,
      render: (item: any) => {
        return <a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.Creator}</a>;

      }
    },
    {
      name: "Approved By",
      displayName: "",
      isResizable: true,
      sorting: true,
      minWidth: 130,
      maxWidth: 130,
      render: (item: any) => {
        let statusHtml;
        if (item.ManagerStatus == "Approved" && item.ReviewForFinace == "Yes") {
          statusHtml = <span><a href="javascript:void(0)" onClick={() => this.OpenViewFrom(item)}>{item.manager}</a></span>
        }
        return <span>
          {statusHtml}
        </span>;
      }
    },
    // {
    //   name: this.state.SelectedTabType == "MySubmission" ? "" : this.state.SelectedTabType == "MyTask" ? "Action" : "",
    //   displayName: "",
    //   minWidth: 65,
    //   maxWidth: 65,
    //   render: (item: any) => {
    //     let isEditBtnDisable = false;
    //     let button;
    //     if (this.state.SelectedTabType == "MySubmission") {
    //       isEditBtnDisable = item.Status == "Rejected by Finance" || item.Status == "Rejected" || item.Status == "Draft" || item.Status == "Clarification" ? false : true;
    //       button = <button disabled={isEditBtnDisable} type="button" className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}>Edit</button>
    //     }
    //     if (this.state.SelectedTabType == "MyTask") {
    //       button = <button disabled={isEditBtnDisable} type="button" className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("EditMySubmission", item)}>Action</button>
    //     }
    //     return <span>
    //       {button}
    //     </span>;
    //   }
    // },

    // {
    //   name: "",
    //   displayName: "",
    //   minWidth: 60,
    //   maxWidth: 60,
    //   render: (item: any) => {
    //     return <div className='btn btn-primary' id="add-row" onClick={() => this.OpenEditForm("ViewMySubmission", item)}>View</div>;
    //   }
    // },
 
    ];
    return viewFields;
  };
  public groupByFields() {
    const groupByFields: IGrouping[] = [
      {
        name: "Status",
        order: GroupOrder.descending
      }
    ];
    return groupByFields
  };
  public renderTabsforApprovers() {
    return <span>{(this.state.IsFinanceDept || this.state.IsManager) &&
      <span>
        <PivotItem itemKey="MyTask" key="MyTask" headerText="Pending" itemCount={this.state.MyPendingItems.length} onClick={() => this.selectedTabType("MyTask")}>
          <ListView
            items={this.state.MyPendingItems}
            showFilter={true}
            filterPlaceHolder="Search..."
            compact={true}
            selectionMode={SelectionMode.single}
            onRenderRow={this.OnListViewRenderRow}
            listClassName={styles.listViewStyle}
            // selection={this.OpenViewFrom}    
            groupByFields={this.groupByFields()}
            viewFields={this.viewFields()}
          />
        </PivotItem>
        <PivotItem itemKey="MyApprovedTask" key="MyApprovedTask" headerText="Approved" itemCount={this.state.MyApprovedItems.length} onClick={() => this.selectedTabType("MyApprovedTask")}>
          <ListView
            items={this.state.MyApprovedItems}
            showFilter={true}
            filterPlaceHolder="Search..."
            compact={true}
            selectionMode={SelectionMode.single}
            onRenderRow={this.OnListViewRenderRow}
            listClassName={styles.listViewStyle}
            // selection={this.OpenViewFrom}    
            groupByFields={this.groupByFields()}
            viewFields={this.viewFields()}
          />
        </PivotItem>
        <PivotItem itemKey="MyRejectedTask" key="MyRejectedTask" headerText="Rejected" itemCount={this.state.MyRejectedItems.length} onClick={() => this.selectedTabType("MyRejectedTask")}>
          <ListView
            items={this.state.MyRejectedItems}
            showFilter={true}
            filterPlaceHolder="Search..."
            compact={true}
            selectionMode={SelectionMode.single}
            onRenderRow={this.OnListViewRenderRow}
            listClassName={styles.listViewStyle}
            // selection={this.OpenViewFrom}    
            groupByFields={this.groupByFields()}
            viewFields={this.viewFields()}
          />
        </PivotItem>

      </span>
    }
    </span>
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
      const item: IItem = sp.web.lists.getByTitle(this.props.expenseListTitle).items.getById(this.state.selectedExpense.ID);
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
  getSelectedExpenseDetail = (selectedItem) => {
    this._spOps.GetListItemByID(selectedItem.ID, this.props.expenseListTitle).then((result) => {
      this._spOps.GetExpenseDetails(selectedItem.ID, this.props.expenseDetailListTitle).then((expenseDetails) => {
        this.setState({
          Department: result.Department,
          ReportHeader: result.Title,
          StartDate: result.StartDate != null ? this.ConvertDate(result.StartDate) : null,
          EndDate: result.EndDate != null ? this.ConvertDate(result.EndDate) : null,
          Manager: result.Manager != undefined ? result.Manager.Title : "",
          Finance: result.Finance != undefined ? result.Finance.Title : "",
          Creator: result.Author != undefined ? result.Author.Title : "",
          CreatorEmail:result.Author != undefined ? result.Author.EMail : "",
          Comments: result.Comments,
          Status: result.Status,
          ManagerStatus:result.ManagerStatus,
          FinanceStatus:result.FinanceStatus,
          TotalExpense: result.TotalExpense,
          IExpenseModel: expenseDetails,
          RequestorResponse:result.RequestorResponse,
          ManagerResponse:result.ManagerResponse,
          FinanceResponse:result.FinanceResponse,

        })
      })
    })
    this.getAttachments();
    this._spOps.GetLogHistoryItems(selectedItem.ID, this.props.logHistoryListTitle).then((logHistory) => {
     this.setState({
      ILogHistoryModel:logHistory
     })
    })
  };
  // public getAllO365Users=()=>{
  //   this.props.context.msGraphClientFactory
  //     .getClient()
  //     .then((client: MSGraphClient): void => {
  //       client
  //         .api('/users/warrensburgpdcamera@besamewellness.com')
  //         .get((error, response: any, rawResponse?: any) => {
  //           console.log(JSON.stringify(response));
  //           // this.setState({
  //           //   displayName: response['displayName']
  //           // })
  //         })
  //     });
  //     this.GetAllusers();
  //  }
  //  public GetAllusers=()=>{
  //   let allusers: any[] = [];
  //     this.props.context.msGraphClientFactory.
  //     getClient().
  //     then((msGraphClient: MSGraphClient) => {
  //       msGraphClient.
  //         api("users").
  //         version("v1.0").
  //         select("displayName,mail").
  //         get((err:any, res:any) => {
  //           if (err) {
  //             console.log("error occurred",err);
  //           }
  //           res.value.map((result: any) => {
  //             allusers.push({
  //               displayName: result.displayName,
  //               mail: result.mail
  //             });
  //           });
  //           //this.setState({userstate:this.allusers});
  //           console.log(allusers);
  //         });
  //     });
  //  }
  GetCurrentUserManagerId = () => {
    this._spOps.GetCurrentUser().then((result) => {
      this.setState({
        CurrentUserName: result["Title"],
      });
      //result["LoginName"]
      this._spOps.getCurrentUserDetails(result["LoginName"]).then((manager) => {
        if (manager != '') {
          this._spOps.getManagerDetails(manager).then((managerEmail) => {
            this._spOps.getUserIDByEmail(managerEmail).then((managerId) => {
              console.log("ManagerEmail",managerEmail)
              //this._spOps.getUserIDByEmailFromAllUsers("kpilarz@besamewellness.com").then((managerId) => {
              this.setState({
                ManagerId: managerId,
              })
            })
          })
        }
      })
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
  // Generate PDF
  public documentprint = (e) => {
    e.preventDefault();
    const myinput = document.getElementById('generatePdf');
    // html2canvas(myinput)
    //   .then((canvas) => {
    //     var imgWidth = 200;
    //     var pageHeight = 290;
    //     var imgHeight = canvas.height * imgWidth / canvas.width;
    //     var heightLeft = imgHeight;
    //     const imgData = canvas.toDataURL('image/png');
    //     const mynewpdf = new jsPDF('p', 'mm', 'a4');
    //     var position = 0;
    //     mynewpdf.addImage(imgData, 'JPEG', 5, position, imgWidth, imgHeight);
    //     mynewpdf.save("SubmittedRecord_" + this.state.selectedExpense.RequestorID + ".pdf");
    html2canvas(myinput, { useCORS: true, allowTaint: true, scrollY: 0 }).then((canvas) => {
      const image = { type: 'jpeg', quality: 0.98 };
      const margin = [0.5, 0.5];
      const filename = 'myfile.pdf';

      var imgWidth = 8.5;
      var pageHeight = 11;

      var innerPageWidth = imgWidth - margin[0] * 2;
      var innerPageHeight = pageHeight - margin[1] * 2;

      // Calculate the number of pages.
      var pxFullHeight = canvas.height;
      var pxPageHeight = Math.floor(canvas.width * (pageHeight / imgWidth));
      var nPages = Math.ceil(pxFullHeight / pxPageHeight);

      // Define pageHeight separately so it can be trimmed on the final page.
      var pageHeight = innerPageHeight;

      // Create a one-page canvas to split up the full image.
      var pageCanvas = document.createElement('canvas');
      var pageCtx = pageCanvas.getContext('2d');
      pageCanvas.width = canvas.width;
      pageCanvas.height = pxPageHeight;

      // Initialize the PDF.
      var pdf = new jsPDF('p', 'in', [8.5, 11]);

      for (var page = 0; page < nPages; page++) {
        // Trim the final page to reduce file size.
        if (page === nPages - 1 && pxFullHeight % pxPageHeight !== 0) {
          pageCanvas.height = pxFullHeight % pxPageHeight;
          pageHeight = (pageCanvas.height * innerPageWidth) / pageCanvas.width;
        }

        // Display the page.
        var w = pageCanvas.width;
        var h = pageCanvas.height;
        pageCtx.fillStyle = 'white';
        pageCtx.fillRect(0, 0, w, h);
        pageCtx.drawImage(canvas, 0, page * pxPageHeight, w, h, 0, 0, w, h);

        // Add the page to the PDF.
        if (page > 0) pdf.addPage();
        debugger;
        var imgData = pageCanvas.toDataURL('image/' + image.type, image.quality);
        pdf.addImage(imgData, image.type, margin[1], margin[0], innerPageWidth, pageHeight);
      }

      pdf.save("SubmittedRecord_"+this.state.selectedExpense.RequestorID+".pdf");
    }); 
    //   });
  }
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
      openFinanceDialog:false,
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
  // // async function with await to save all expenses into "ExpenseDetails" list
  // private async AddExpensesDetails(Expenses: any[], expenseItemId: number) {
  //   let web = Web(this.props.siteUrl);
  //   let requestorUniqueID = this.GetUniqueRequestorID(expenseItemId);
  //   for (const expense of Expenses) {
  //     await web.lists.getByTitle(this.props.expenseDetailListTitle).items.add({
  //       ExpenseTypes: expense.Expense,
  //       CheckIn: expense.Checkin != "" ? expense.Checkin : null,
  //       CheckOut: expense.Checkout != "" ? expense.Checkout : null,
  //       ExpenseDate: expense.ExpenseDate != "" ? expense.ExpenseDate : null,
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
  // async function with await to delete expenses from "ExpenseDetails" list
  private async DeleteExpensesDetails(Expenses: any[]) {
    let web = Web(this.props.siteUrl);
    for (const expense of Expenses) {
      await web.lists.getByTitle(this.props.expenseDetailListTitle).items.getById(expense.Id).delete();
    }
  }
   // update all expenses into "ExpenseDetails" list
   private async UpdateExpenseDetailsAsBatch(Expenses: any[]) {
    let sourceWeb = Web(this.props.siteUrl);
    let taskList = sourceWeb.lists.getByTitle(this.props.expenseDetailListTitle);
    let batch = sourceWeb.createBatch();
    console.log("batch = ", JSON.stringify(batch));
    console.log("batch baseURL = ", batch["baseUrl"]);
    for (let i = 0; i < Expenses.length; i++) {
      taskList.items.getById(Expenses[i].Id).inBatch(batch).update(
          {
        ExpenseTypes: Expenses[i].Expense,
        CheckIn: Expenses[i].Checkin != "" ? Expenses[i].Checkin : null,
        CheckOut: Expenses[i].Checkout != "" ? Expenses[i].Checkout : null,
        ExpenseDate:Expenses[i].ExpenseDate!="" ? Expenses[i].ExpenseDate:null,
        TravelType: Expenses[i].TravelType,
        StartMile: Expenses[i].StartMile,
        EndMile: Expenses[i].EndMileout,
        Description: Expenses[i].Description,
        ExpenseCost: String(Expenses[i].ExpenseCost)
          }
        )
        .then((result:any) => {
          console.log("Item updated with id", Expenses[i].Id);
        })
        .catch((ex) => {
          console.log(ex);
        });
    }
    await batch.execute();
    console.log("Done");
    }
  // async function with await to update all expenses into "ExpenseDetails" list
  // private async UpdateExpenseDetails(Expenses: any[]) {
  //   let web = Web(this.props.siteUrl);
  //   for (const expense of Expenses) {
  //     await web.lists.getByTitle(this.props.expenseDetailListTitle).items.getById(expense.Id).update({
  //       ExpenseTypes: expense.Expense,
  //       CheckIn: expense.Checkin != "" ? expense.Checkin : null,
  //       CheckOut: expense.Checkout != "" ? expense.Checkout : null,
  //       ExpenseDate: expense.ExpenseDate != "" ? expense.ExpenseDate : null,
  //       TravelType: expense.TravelType,
  //       StartMile: expense.StartMile,
  //       EndMile: expense.EndMileout,
  //       Description: expense.Description,
  //       ExpenseCost: expense.ExpenseCost,
  //     });
  //   }
  // }
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
    if (submissionType == "Submitted" ||submissionType == "Draft") {
      if (this.state.ReportHeader == null) {
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
  };

  private isStatusBarModalOpen(): boolean {
    return this.state.openStatusBarDialog;
  };
  private hideStatusBarModal=()=>{
    this.setState({
      openStatusBarDialog:false
    })
    this.RefreshPage();
  };

  private isEditModalOpen(): boolean {
    return this.state.openEditDialog;
  };
  private isFinanceModalOpen():boolean{
    return this.state.openFinanceDialog;
  }
  private hideFinanceModal = () => {
    this.setState({
      openFinanceDialog: false
    })
  };
  private isModalOpen(): boolean {
    return this.state.openDialog;
  };
  private isInvoiceModalOpen(): boolean {
    return this.state.openInvoiceDialog;
  };
  private hideEditModal = () => {
    this.setState({
      openEditDialog: false
    })
  };
 
  UpdateFinanceDetails=()=>{
    let updateFinanceData: any = {};
    updateFinanceData = {
        FinanceId: this.state.NewFinanceUserID,
        OtherFinanceId:this.state.FinanceId
      }
    this._spOps.UpdateItem("EMSConfiguration", updateFinanceData, this.state.FinanceItemId).then((result: any) => {
            alert("Finance has been changed sucessfully.");
            this.setState({
              openFinanceDialog:false
            })
            this.RefreshPage(); 
    })
    
  }
 
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
        this.AddExpensesDetailsAsBatch(expenseItemsCreate, itemId).then(() => {
          this.UpdateExpenseDetailsAsBatch(expenseItemsUpdate).then(() => {
            $('#loaderEdit').hide();
            alert("Request " + submissionType + " sucessfully");
            this.RefreshPage();
          })

        });
      }
      if (expenseItemsCreate.length > 0 && expenseItemsUpdate.length == 0) {
        this.AddExpensesDetailsAsBatch(expenseItemsCreate, itemId).then(() => {
          $('#loaderEdit').hide();
          alert("Request " + submissionType + " sucessfully");
          this.RefreshPage();
        })
      }
      if (expenseItemsUpdate.length > 0 && expenseItemsCreate.length == 0) {
        this.UpdateExpenseDetailsAsBatch(expenseItemsUpdate).then(() => {
          $('#loaderEdit').hide();
          alert("Request " + submissionType + " sucessfully");
          this.RefreshPage();
          //alert(submissionType=="Submited"?"Request submitted sucessfully":"Request drafted sucessfully");
        })
      }
    }
    else {
      $('#loaderEdit').hide();
      alert("Request " + submissionType + " sucessfully");
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
          totalAmount += expense.ExpenseCost != undefined ? parseFloat(expense.ExpenseCost) : 0;
        })
      }
      let updatePostData: any = {};
      updatePostData = {
        Department: this.state.Department,
        Title: this.state.ReportHeader,
        StartDate: this.state.StartDate != "" ? this.state.StartDate : null,
        EndDate: this.state.EndDate != "" ? this.state.EndDate : null,
        FinanceId: this.state.FinanceId,
        TotalExpense: totalAmount.toFixed(2),
      }
      if (this.state.latestComments != "") {
        updatePostData.Comments = latestCommentsHTML.concat(previousComments);                            
      }
      if (this.state.SelectedTabType == "MySubmission") {
        if (submissionType == "Submitted") {
          updatePostData.ManagerStatus = this.state.ManagerId != null ? "Pending for Manager" : "Manager Approval Not Required";
          updatePostData.Status = this.state.ManagerId != null ? "Submitted" : "InProgress";
          updatePostData.FinanceStatus = this.state.ManagerId != null ? "" : "Pending for Finance";
          updatePostData.ReviewForFinace = this.state.ManagerId != null ? "" : "Yes";
          updatePostData.ManagerId = this.state.ManagerId;
        if (this.state.selectedExpense.FinanceStatus == "Clarification" && this.state.ManagerId != null) {
            updatePostData.ManagerStatus = "Approved";
            updatePostData.FinanceStatus = "Pending for Finance";
            updatePostData.Status = "InProgress";
            updatePostData.ReviewForFinace = "Yes";
          }
          if ((this.state.selectedExpense.FinanceStatus == "Clarification" && this.state.ManagerId == null) || (this.state.selectedExpense.FinanceStatus == "Rejected" && this.state.ManagerId == null)) {
            updatePostData.ManagerStatus = "Manager Approval Not Required";
            updatePostData.FinanceStatus = "Pending for Finance";
            updatePostData.Status = "InProgress";
            updatePostData.ReviewForFinace = "Yes";
          }
          if(this.state.FinanceStatus=="Rejected" && this.state.ManagerId!=null){
            updatePostData.FinanceStatus = this.state.FinanceStatus;
          }

        }
        if(submissionType == "Draft"){
          updatePostData.ManagerStatus = "";
          updatePostData.Status = "Draft";
          updatePostData.ManagerId = null;//this.state.ManagerId;
          updatePostData.FinanceId = null;//this.state.FinanceId,
        }
      }
      if (this.state.SelectedTabType == "MyTask" && (!this.state.IsFinanceDept || this.state.ManagerStatus=="Pending for Manager")) {
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
      if (this.state.SelectedTabType == "MyTask" && this.state.IsFinanceDept && (this.state.ManagerStatus=="Approved" || this.state.ManagerStatus=="Manager Approval Not Required")) {
        updatePostData.FinanceStatus = submissionType;
        updatePostData.Status = submissionType;
        if (submissionType == "Paid") {
          updatePostData.Status = "Paid";
          updatePostData.AmountPaidDate = this.state.PaidDate;
        }
        if (submissionType == "Rejected" && (this.state.ManagerStatus=="Approved" || this.state.ManagerStatus=="Manager Approval Not Required")) {
          updatePostData.ManagerStatus = this.state.ManagerStatus;
          updatePostData.Status = "Rejected by Finance";
        }
      }
      let logHistoryPostData:any={};
      logHistoryPostData={
        Title:this.state.selectedExpense.RequestorID,
        ExpensesId:this.state.selectedExpense.ID,
        CommentsHistory:this.state.latestComments,
        Status:updatePostData.Status,
       // NameId:this.state.CurrentUserID
      }
      this._spOps.CreateItem(this.props.logHistoryListTitle, logHistoryPostData).then((result: any) => {});
      this._spOps.UpdateItem(this.props.expenseListTitle, updatePostData, this.state.selectedExpense.ID).then((result: any) => {
        console.log(this.state.selectedExpense.ID);
        let itemId = this.state.selectedExpense.ID;
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
          web.lists.getByTitle(this.props.expenseListTitle).items.getById(this.state.selectedExpense.ID).attachmentFiles.deleteMultiple(...this.state.FilesToDelete);
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
        ExpenseDate: "",
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
  //select user for Finance
  private _getPeoplePickerItems=(items: any[])=> {
     
    let newFinanceUserEmail=items.length>0?items[0].loginName.split('|membership|')[1].toString() :"";
    if(items.length>0){
    this._spOps.getUserIDByEmail(newFinanceUserEmail).then((newFinanceUserID) => {
       this.setState({
        NewFinanceUserID:newFinanceUserID
       })
    })
  }
  }

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
    let isShowAttachDelete = this.state.SelectedTabType != "MyTask" && this.state.selectedExpense.formType == "Edit" ? true : false;
    return this.state.fileInfos.map((item, idx) => {
      return (<div><a href={item.ServerRelativeUrl != undefined ? item.ServerRelativeUrl : "javascript:void(0)"} target="_blank" data-interception="off">{item.name}</a>&nbsp;&nbsp;
        <span id="delete-spec-row" className={contentEditFormStyles.deleteIcon}>
          {isShowAttachDelete &&
            <Icon iconName="delete" className="ms-IconExample" onClick={this.removeSpecificAttachment(idx, item)} />
          }
        </span></div>)
    })
  }
  //** Open Paid PopUp to select paid date*/
  OpenPaidPopUp = () => {
    this.setState({
      openDialog: true
    })
  }
//** Open Invoice PopUp*/
  OpenInvoicePopUp = () => {
    this.setState({
      openInvoiceDialog: true
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
    let isExpenseDeleteIcon = this.state.SelectedTabType != "MyTask" && this.state.selectedExpense.formType == "Edit" ? true : false;
    return this.state.IExpenseModel.map((item, idx) => {
      let IsExpenseTypeDisable = true;
      if (item.Id != undefined && this.state.Status == "Draft") {
        IsExpenseTypeDisable = false;
      }
      if (item.Id == undefined) {
        IsExpenseTypeDisable = false;
      }
      let costTxtCss = "form-group col-md-3";
      if (item.Expense == "Travel") {
        costTxtCss = "form-group col-md-1"
      }
      if (item.Expense == "Hotel") {
        costTxtCss = "form-group col-md-2"
      }
      return (<span key={idx}>
        {/* <div className={styles.expenselbl}>Expense {idx+1}</div> */}
        <div className={this.state.SelectedTabType != "MyTask" && this.state.selectedExpense.formType == "Edit" ? contentEditFormStyles.formRow : contentEditFormStyles.editRequestTable}>
          <div className="form-group col-md-2">
            <label className={contentEditFormStyles.lblCtrl}>Expense Type<span className={contentEditFormStyles.star}>*</span></label>
            <select disabled={IsExpenseTypeDisable || this.state.selectedExpense.formType == "View" || this.state.SelectedTabType == "MyTask"} className='form-control ExpTypeSelectOptions' name="Expense" defaultValue={this.state.IExpenseModel[idx].Expense} value={this.state.IExpenseModel[idx].Expense} id={this.state.IExpenseModel[idx].id} onChange={this.handleChange(idx)}>
              <option value="">Select</option>
              {this.renderDropdown(this.state.ExpenseTypeOptions)}
            </select>
            {item.isExpenseTypeError == true &&
              <span className={contentEditFormStyles.errMsg}>{item.ExpenseTypeErrMsg}</span>
            }
          </div>
          {item.Expense == "Hotel" &&
            <div className="form-group col-md-2">
              <label className={contentEditFormStyles.lblCtrl}>Check In &nbsp;<span className={contentEditFormStyles.noteMsg}>(DD-MM-YYYY)</span><span className={contentEditFormStyles.star}>*</span></label>
              <input
                placeholder='check in'
                type="date"
                disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false}
                className='form-control'
                name="Checkin"
                defaultValue={this.state.IExpenseModel[idx].Checkin}
                value={this.state.IExpenseModel[idx].Checkin}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />

              {item.isCheckinError == true && <span className={contentEditFormStyles.errMsg}>{item.CheckinErrMsg}</span>}
            </div>
          }
          {item.Expense == "Hotel" &&
            <div className="form-group col-md-2">
              <label className={contentEditFormStyles.lblCtrl}>Check Out &nbsp;<span className={contentEditFormStyles.noteMsg}>(DD-MM-YYYY)</span><span className={contentEditFormStyles.star}>*</span></label>
              <input
                placeholder='Check out'
                type="date"
                disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false}
                className='form-control'
                name="Checkout"
                value={this.state.IExpenseModel[idx].Checkout}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
              {item.isCheckoutError == true && <span className={contentEditFormStyles.errMsg}>{item.CheckoutErrMsg}</span>}
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-2">
              <label className={contentEditFormStyles.lblCtrl}>Travel Type</label>
              <select disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false} className='form-control ExpTypeSelectOptions' name="TravelType" defaultValue={this.state.IExpenseModel[idx].TravelType} value={this.state.IExpenseModel[idx].TravelType} id={this.state.IExpenseModel[idx].id} onChange={this.handleChange(idx)}>
                <option value="">Select</option>
                {this.renderDropdown(this.state.TravelTypeOption)}
              </select>
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-1">
              <label className={contentEditFormStyles.lblCtrl}>Start Mile</label>
              <input
                placeholder={item.TravelType == "Cab/Uber/Lyft" || item.TravelType == "Airfare"?'N/A':'Start'}
                type="number"
                min="1"
                disabled={this.state.SelectedTabType == "MyTask" || item.TravelType == "Cab/Uber/Lyft" || item.TravelType == "Airfare" || this.state.selectedExpense.formType == "View" ? true : false}
                className={contentEditFormStyles.formControl}
                name="StartMile"
                value={this.state.IExpenseModel[idx].StartMile}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
            </div>
          }
          {item.Expense == "Travel" &&
            <div className="form-group col-md-1">
              <label className={contentEditFormStyles.lblCtrl}>End Mile</label>
              <input
                placeholder={item.TravelType == "Cab/Uber/Lyft" || item.TravelType == "Airfare"?'N/A':'End'}
                type="number"
                min="1"
                disabled={this.state.SelectedTabType == "MyTask" || item.TravelType == "Cab/Uber/Lyft" || item.TravelType == "Airfare" || this.state.selectedExpense.formType == "View" ? true : false}
                className={contentEditFormStyles.formControl}
                name="EndMileout"
                value={this.state.IExpenseModel[idx].EndMileout}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
              {item.isEndMileError && <span className={contentEditFormStyles.errMsg}>{item.EndMileErrMsg}</span>}
            </div>
          }
          {/* {(item.Expense == "Meal" || item.Expense == "Others") && */}
          <div className={item.Expense == "Travel" ? "form-group col-md-2" : "form-group col-md-3"}>
            <label className={contentEditFormStyles.lblCtrl}>Description</label>
            <input
              placeholder='Description'
              type="text"
              disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false}
              className='form-control'
              name="Description"
              value={this.state.IExpenseModel[idx].Description}
              onChange={this.handleChange(idx)}
              id={this.state.IExpenseModel[idx].id}
            />
          </div>
          {/* } */}
          {(item.Expense != "Hotel") &&
            <div className="form-group col-md-2">
              <label className={contentEditFormStyles.lblCtrl}>Date &nbsp;<span className={contentEditFormStyles.noteMsg}>(DD-MM-YYYY)</span></label>
              <input
                placeholder='Date'
                type="date"
                disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false}
                className='form-control'
                name="ExpenseDate"
                defaultValue={this.state.IExpenseModel[idx].ExpenseDate}
                value={this.state.IExpenseModel[idx].ExpenseDate}
                onChange={this.handleChange(idx)}
                id={this.state.IExpenseModel[idx].id}
              />
            </div>
          }
          <div className={costTxtCss}>
            <label className={contentEditFormStyles.lblCtrl}>Amount($)<span className={contentEditFormStyles.star}>*</span> &nbsp;{item.Expense == "Travel" && this.state.IExpenseModel[idx].TravelType == "Leased" && <span className={contentEditFormStyles.noteMsg}>(${this.state.StartEndMileCosts}/milage)</span>}</label>
            <input
              placeholder='Amount'
              type="number"
              min="1"
              disabled={item.TravelType == "Leased" || this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false}
              className={item.Expense == "Travel" ? contentEditFormStyles.formControl : 'form-control'}
              name="ExpenseCost"
              value={this.state.IExpenseModel[idx].ExpenseCost}
              onChange={this.handleChange(idx)}
              id={this.state.IExpenseModel[idx].id}
            />
             {item.isExpenseCostError == true &&
              <span className={contentEditFormStyles.errMsg}>{item.ExpenseCostMsg}</span>
            }
          </div>
          {this.state.IExpenseModel.length > 1 &&
            <div className="form-group col-md-1">
              <label></label>
              <div id="delete-spec-row" className={contentEditFormStyles.deleteIcon}>
                {isExpenseDeleteIcon &&
                  <Icon iconName="delete" className="ms-IconExample" onClick={this.handleRemoveSpecificRow(idx)} />
                }
              </div>
            </div>
          }
          <div className={contentEditFormStyles.itemLine}></div>
        </div>

      </span>)
    })
  }


  //** onClick of View Expense, Generate html for PDF*/
  renderInvoiceTableData() {
    return this.state.IExpenseModel.map((item, idx) => {
      return (<tbody key={idx}>
                    <tr>
                      <td className={contentInvoiceStyles.td}>{this.state.IExpenseModel[idx].Expense}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.IExpenseModel[idx].Checkin} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IExpenseModel[idx].Checkout} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IExpenseModel[idx].TravelType} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IExpenseModel[idx].StartMile} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IExpenseModel[idx].EndMileout} </td>
                      <td className={contentInvoiceStyles.td}>{this.state.IExpenseModel[idx].Description}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.IExpenseModel[idx].ExpenseCost}</td>
                    </tr>
                  </tbody>)               
    })
  }
//** onClick of View Expense, Generate html for PDF*/
renderInvoiceCommentTableData() {
  return this.state.ILogHistoryModel.map((item, idx) => {
    return (<tbody key={idx}>
                  <tr>
                    <td className={contentInvoiceStyles.td}>{this.state.ILogHistoryModel[idx].CommentsHistory}</td>
                    <td className={contentInvoiceStyles.td}>{this.state.ILogHistoryModel[idx].Author} </td>
                    <td className={contentInvoiceStyles.td}>{this.state.ILogHistoryModel[idx].Status} </td>
                    <td className={contentInvoiceStyles.td}>{this.state.ILogHistoryModel[idx].CreatedOn} </td>
                  </tr>
                </tbody>)               
  })
}
managerSectionCss(){
  let mngrCss="";
  switch (this.state.ManagerStatus) {
    case 'Approved':
      mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ");
      break;
    case 'Manager Approval Not Required':
      mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ");
      break;
    case 'Rejected':
      mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
      break;
    case 'Rejected by Finance':
      mngrCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
      break;
    case 'Pending for Manager':
      mngrCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.pending].join(" ");
      break;
      default:
     mngrCss = [contentStatusBarStyles.statuscircle,""].join(" "); 
  }
  return mngrCss;
};
financeSectionCss(){
  let finCss="";
  switch (this.state.FinanceStatus) {
    case 'Approved':
      finCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ");
      break;
    case 'Paid':
      finCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ");
      break;
    case 'Clarification':
      finCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.clarify].join(" ");
      break;
    case 'Rejected':
      finCss = [contentStatusBarStyles.statuscircle, contentStatusBarStyles.rejected].join(" ");
      break;
    case 'Pending':
      finCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.pending].join(" ");
      break;
    case 'Pending for Finance':
      finCss = [contentStatusBarStyles.statuscircle,contentStatusBarStyles.pending].join(" ");
      break;
    default:
      finCss = [contentStatusBarStyles.statuscircle,""].join(" ");
  }
  return finCss;
}
  public render(): React.ReactElement<IEmsProps> {
   
    let managerCss="";
    let financeCss="";
    managerCss= this.managerSectionCss();
    financeCss=this.financeSectionCss();
    let isShowAddExpenseBtn = this.state.SelectedTabType != "MyTask" && this.state.selectedExpense.formType == "Edit" ? true : false;

    let renderPendingTabs: any;
    let renderApprovedTabs: any;
    let renderRejectedTabs: any;
    let renderPaidTabs: any;
    if (this.state.IsFinanceDept || this.state.IsManager) {
      renderPendingTabs = <PivotItem itemKey="MyTask" key="MyTask" headerText="Pending" itemCount={this.state.MyPendingItems.length} onClick={() => this.selectedTabType("MyTask")}>
        <ListView
          items={this.state.MyPendingItems}
          showFilter={true}
          filterPlaceHolder="Search..."
          compact={true}
          //selectionMode={SelectionMode.single}
          onRenderRow={this.OnListViewRenderRow}
          listClassName={styles.listViewStyle}
          // selection={this.OpenViewFrom}    
          groupByFields={this.groupByFields()}
          viewFields={this.viewFields()}
        />
      </PivotItem>
    }
    if (this.state.IsFinanceDept || this.state.IsManager) {
      renderApprovedTabs = <PivotItem itemKey="MyApprovedTask" key="MyApprovedTask" headerText="Approved" itemCount={this.state.MyApprovedItems.length} onClick={() => this.selectedTabType("MyApprovedTask")}>
        <ListView
          items={this.state.MyApprovedItems}
          showFilter={true}
          filterPlaceHolder="Search..."
          compact={true}
          //selectionMode={SelectionMode.single}
          onRenderRow={this.OnListViewRenderRow}
          listClassName={styles.listViewStyle}
          // selection={this.OpenViewFrom}  
          groupByFields={this.groupByFields()}
          viewFields={this.viewFields()}
        />
      </PivotItem>
    }
    if (this.state.IsFinanceDept) {
      renderPaidTabs = <PivotItem itemKey="MyPaidTask" key="MyPaidTask" headerText="Paid" itemCount={this.state.MyPaidItems.length} onClick={() => this.selectedTabType("MyPaidTask")}>
        <ListView
          items={this.state.MyPaidItems}
          showFilter={true}
          filterPlaceHolder="Search..."
          compact={true}
          // selectionMode={SelectionMode.single}
          onRenderRow={this.OnListViewRenderRow}
          listClassName={styles.listViewStyle}
          // selection={this.OpenViewFrom}    

          groupByFields={this.groupByFields()}
          viewFields={this.viewFields()}
        />
      </PivotItem>
    }
    if (this.state.IsFinanceDept || this.state.IsManager) {
      renderRejectedTabs = <PivotItem itemKey="MyRejectedTask" key="MyRejectedTask" headerText="Rejected" itemCount={this.state.MyRejectedItems.length} onClick={() => this.selectedTabType("MyRejectedTask")}>
        <ListView
          items={this.state.MyRejectedItems}
          showFilter={true}
          filterPlaceHolder="Search..."
          compact={true}
          //selectionMode={SelectionMode.single}
          onRenderRow={this.OnListViewRenderRow}
          listClassName={styles.listViewStyle}
          // selection={this.OpenViewFrom}    
          groupByFields={this.groupByFields()}
          viewFields={this.viewFields()}
        />
      </PivotItem>

    }
    return (
      <div className={styles.ems}>
        <div className="row">
          <div className="form-group col-md-7">
            <h4>Expense Management</h4>
          </div>
          <div className="form-group col-md-5">
            <a href="https://web.microsoftstream.com/video/e3db047a-ad0c-4d18-99be-956b083b8ead" target="_blank" data-interception="off">User Guide </a>
            {this.state.IsManager == true && <a href="https://web.microsoftstream.com/video/e9015206-22a5-41f2-bee0-7fabbbcb2fad" target="_blank" data-interception="off">| Manager Guide </a>}&nbsp;&nbsp;
            {this.state.IsFinanceDept == true && <a href="https://web.microsoftstream.com/video/689fe29c-7c42-4613-ab08-4023f0de0c2e" target="_blank" data-interception="off">| Finance Guide</a>}&nbsp;&nbsp;
            {(this.state.CurrentUserEmail == this.state.FinanceEmailID || this.state.CurrentUserEmail == this.state.OtherFinanceEmailID) &&  <a href="javascript:void(0)" onClick={this.OpenFinanceModal}>| Change Finance</a>}
          </div>
        </div>
        <Pivot aria-label="OnChange Pivot Example" onLinkClick={this.onPivotItemClick}>
          <PivotItem itemKey="NewRequest" key="NewRequest" headerText="Submit new request">
            <NewRequestForm context={this.props.context} siteUrl={this.props.siteUrl} logHistoryListTitle="LogHistory" expenseListTitle="Expenses" expenseDetailListTitle="ExpenseDetails"></NewRequestForm>
          </PivotItem>
          <PivotItem itemKey="MySubmission" key="MySubmission" headerText="My Submission" itemCount={this.state.MySubmissionItems.length} onClick={() => this.selectedTabType("MySubmission")}>
            <ListView
              items={this.state.MySubmissionItems}
              showFilter={true}
              filterPlaceHolder="Search..."
              compact={true}
              //selectionMode={SelectionMode.single}
              onRenderRow={this.OnListViewRenderRow}
              listClassName={styles.listViewStyle}
              // selection={this.OpenViewFrom}    
              groupByFields={this.groupByFields()}
              viewFields={this.viewFields()}
            />
          </PivotItem>
          {renderPendingTabs}
          {renderApprovedTabs}
          {renderPaidTabs}
          {renderRejectedTabs}



        </Pivot>
        {this.state.openEditDialog &&
          <Modal isOpen={this.isEditModalOpen()} isBlocking={false} containerClassName={contentEditFormStyles.container}>
            <div className={contentEditFormStyles.header}>
              <span> {this.state.selectedExpense.formType == "Edit" ? "Edit Expense - " : "View Expense - "}{this.state.selectedExpense.RequestorID}

              </span>&nbsp;
              {(this.state.IsFinanceDept || this.state.Status=="Paid") && <span className={contentEditFormStyles.viewInvoice}>
               
                <button className='btn btn-warning' id="viewInvoice" onClick={this.OpenInvoicePopUp} title="View Invoice">View Invoice</button>
                </span>
                }

              <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideEditModal} />
            </div>
            {/* <EditRequestForm tabType= {this.state.SelectedTabType} selectedItem= {this.state.selectedExpense} context={this.props.context} siteUrl={this.props.siteUrl} expenseListTitle="Expenses" expenseDetailListTitle="ExpenseDetails"></EditRequestForm> */}

            <span id="generatePdfForm">
              <div className={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? contentEditFormStyles.viewDeptSection : contentEditFormStyles.deptSection}>
                <div className="form-group col-md-3">
                  <label className={contentEditFormStyles.lblCtrl}>Department<span className={contentEditFormStyles.star}>*</span></label>
                  <select disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false} className='form-control' name="Department" value={this.state.Department} id="Department" onChange={this.handleChange(1)}>
                    <option value="">Select</option>
                    {this.renderDropdown(this.state.DepartmentOptions)}
                  </select>
                  {this.state.IsDeptErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.DeptErrMsg}</span>}
                </div>
                <div className="form-group col-md-3">
                  <label className={contentEditFormStyles.lblCtrl}>Description<span className={contentEditFormStyles.star}>*</span></label>
                  <input
                    placeholder='Description'
                    type="text"
                    disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false}
                    className='form-control'
                    name="ReportHeader"
                    value={this.state.ReportHeader}
                    onChange={this.handleChange(2)}
                    id="ReportHeader"
                  />
                  {this.state.IsReportErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.ReportHeaderErrMsg}</span>}
                </div>
                <div className="form-group col-md-2">
                  <label className={contentEditFormStyles.lblCtrl}>Start Date&nbsp;<span className={contentEditFormStyles.noteMsg}>(MM-DD-YYYY)</span></label>
                  <input
                    placeholder='Start Date'
                    type="date"
                    disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false}
                    className='form-control'
                    name="StartDate"
                    value={this.state.StartDate}
                    onChange={this.handleChange(3)}
                    id="StartDate"
                  />
                  {this.state.IsStartDateErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.StartDateErrMsg}</span>}
                </div>
                <div className="form-group col-md-2">
                  <label className={contentEditFormStyles.lblCtrl}>End &nbsp;<span className={contentEditFormStyles.noteMsg}>(MM-DD-YYYY)</span></label>
                  <input
                    placeholder='End Date'
                    type="date"
                    disabled={this.state.SelectedTabType == "MyTask" || this.state.selectedExpense.formType == "View" ? true : false}
                    className='form-control'
                    name="EndDate"
                    value={this.state.EndDate}
                    onChange={this.handleChange(4)}
                    id="EndDate"
                  />
                  {this.state.IsEndDateErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.EndDateErrMsg}</span>}
                </div>
                <div className="form-group col-md-2">
                  <label className={contentEditFormStyles.lblCtrl}>Total Amount($)</label>
                  <input
                    placeholder='Total Amount'
                    type="text"
                    disabled={true}
                    className='form-control'
                    name="TotalAmount"
                    value={this.state.TotalExpense}
                    onChange={this.handleChange(5)}
                    id="TotalAmount"
                  />
                </div>

              </div>
              {this.renderTableData()}
              {isShowAddExpenseBtn &&
                <div className='btn btn-primary' id="add-row" onClick={this.handleAddRow}>Add Expense</div>
              }

              {this.state.IsExpenseDetailErr == true && <span className={contentEditFormStyles.errMsg}>{this.state.ExpenseDetailErrMsg}</span>}
              <div>
                <div className={contentEditFormStyles.line}></div>
              </div>
              <div className="form-row">
                <div className="form-group col-md-4">

                  <label className={contentEditFormStyles.lblCtrl}>Comments</label>
                  <textarea
                    className='form-control'
                    name="Comments"
                    cols={6}
                    rows={3}
                    disabled={this.state.selectedExpense.Status == "Approved" || this.state.selectedExpense.formType == "View" ? true : false}
                    value={this.state.latestComments}
                    onChange={this.handleChange(5)}
                    id="Comments">
                  </textarea>
                </div>
                {this.state.SelectedTabType != "MyTask" && this.state.selectedExpense.formType == "Edit" &&
                  <div className="form-group col-md-3">
                    <label className={contentEditFormStyles.lblCtrl}>Attachment(s)</label>
                    <input type="file" disabled={this.state.selectedExpense.formType == "View" ? true : false} multiple={true} id="file" onChange={this.addFile.bind(this)} />
                    {this.state.isMealExpenseCostError == true && <span className={contentEditFormStyles.errMsg}>{this.state.mealExpenseCostErrMsg}</span>}
                  </div>
                }

                <div className="form-group col-md-4">
                  {this.state.fileInfos.length > 0 &&
                    <label id="fileName" className={contentEditFormStyles.lblCtrl}>Attached Files </label>
                  }
                  {this.renderAttachmentName()}

                </div>
                {this.state.selectedExpense.formType == "View" &&
                  <div className="form-group col-md-2">
                    <label className={contentEditFormStyles.lblCtrl}>Submitted By</label>
                    <input
                      placeholder='Submitted By'
                      type="text"
                      disabled={true}
                      className='form-control'
                      name="Creator"
                      value={this.state.Creator}
                      onChange={this.handleChange(2)}
                      id="Creator"
                    />

                  </div>
                }
                {this.state.selectedExpense.formType == "View" &&
                  <div className="form-group col-md-2">
                    <label className={contentEditFormStyles.lblCtrl}>Approved By</label>
                    <input
                      placeholder='Approved By'
                      type="text"
                      disabled={true}
                      className='form-control'
                      name="Manager"
                      value={this.state.Manager}
                      onChange={this.handleChange(2)}
                      id="Manager"
                    />

                  </div>
                }
              </div>
              <div className="form-row">
                <div className="form-group col-md-6">
                  {this.state.Comments != null &&
                    <span> <label className={contentEditFormStyles.lblCtrl}>Comments History</label>
                      <div className={contentEditFormStyles.commentContainer}>
                        <div className={contentEditFormStyles.cmtHistoryRow}>
                          <div className={contentEditFormStyles.comment}>
                            <div dangerouslySetInnerHTML={{ __html: this.state.Comments }}></div>
                          </div>
                        </div>
                      </div>
                    </span>
                  }
                </div>
              </div>
            </span>
            <div>
              <span className={contentEditFormStyles.btnRt}>
                {this.state.selectedExpense.Status != "Approved" && <span>
                  {this.state.SelectedTabType == "MySubmission" && this.state.selectedExpense.formType == "Edit" &&
                    <span><button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Submitted")} title="Submit">Submit</button>  &nbsp;&nbsp;
                      <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Draft")} title="Save as Draft">Save as Draft</button> &nbsp;&nbsp;
                    </span>
                  }
                  {this.state.SelectedTabType == "MyTask" && (this.state.IsManager == true || this.state.ManagerStatus == "Pending for Manager") && this.state.selectedExpense.formType == "Edit" &&
                    <span><button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Approved")} title="Approve">Approve</button>  &nbsp;&nbsp;
                      <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Rejected")} title="Reject">Reject</button> &nbsp;&nbsp;
                    </span>
                  }
                  {this.state.SelectedTabType == "MyTask" && this.state.IsFinanceDept == true && this.state.selectedExpense.formType == "Edit" && (this.state.ManagerStatus=="Approved" || this.state.ManagerStatus=="Manager Approval Not Required") && 
                    <span><button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.OpenPaidPopUp()} title="Paid">Paid</button>  &nbsp;&nbsp;
                      <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Clarification")} title="Clarification">Clarification</button> &nbsp;&nbsp;
                      <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={() => this.UpdateRequest("Rejected")} title="Reject">Reject</button> &nbsp;&nbsp;

                    </span>
                  }
                </span>
                }

                <button disabled={this.state.IsBtnClicked} className='btn btn-primary' id="add-row" onClick={this.RefreshPage} title="Close">Close</button>
              </span>
              <div id="loaderEdit" className={contentEditFormStyles.loaderEdit}></div>
            </div>

          </Modal>
        }
        {this.state.openFinanceDialog &&
          <Modal isOpen={this.isFinanceModalOpen()} isBlocking={false} containerClassName={contentStyles.container}>
            <div className={contentStyles.body}>
              <div className={contentStyles.header}>
                <span className={styles.label}> Select New Finance</span>
                <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideFinanceModal} />
              </div>

              <div className="form-row">
                <div className="form-group col-md-12">
                  {/* <label className="control-label">Paid Date</label> */}
                  <PeoplePicker
                    context={this.props.context as any}
                    titleText=""
                    personSelectionLimit={1}
                    showtooltip={true}
                    onChange={this._getPeoplePickerItems}
                    required={true}
                   // isRequired={true}
                   // selectedItems={this._getPeoplePickerItems}
                    principalTypes={[PrincipalType.User]}
                    defaultSelectedUsers={this.state.FinanceEmailID}
                    resolveDelay={1000} />
                  {/* {this.state.IsPaidDateErr == true && <span style={{ 'color': 'Red' }}>{this.state.PaidDateErrMsg}</span>} */}
                </div>
              </div>

              <div className="form-row">
                <div className="form-group col-md-12">
                  <button className='btn btn-primary float-right' id="add-row" onClick={this.UpdateFinanceDetails}>Change</button>
                </div>
              </div>
            </div>
          </Modal>
        }
        {this.state.openDialog &&
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
        }
        {this.state.openInvoiceDialog &&
          <Modal isOpen={this.isInvoiceModalOpen()} isBlocking={false} containerClassName={contentInvoiceStyles.container}>
            <div>
              <div className={contentInvoiceStyles.header}>
                &nbsp;<img title="Print PDF" src={this.props.siteUrl + '/SiteAssets/ICONS/print.png'} onClick={this.documentprint} />
                <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideInvoiceModal} />
              </div>

              <div className={contentInvoiceStyles.Containerbox} id="generatePdf">
                <h1 className={contentInvoiceStyles.h1}>Expense Invoice</h1>
                <div className="form-row">
                <div className="col-md-8">
                <div className={contentInvoiceStyles.invoiceinfo}>
                  {/* <p><strong>Expense Date:</strong> {this.state.PaidDate}</p> */}
                  <p><strong>Status:</strong> {this.state.Status}</p>
                  {/* </br> */}
                  <p><strong>{this.state.Creator}</strong></p>
                  {/* <p><strong>Tel:</strong> 7900589437</p> */}
                  <p><strong>Email:</strong> {this.state.CreatorEmail}</p>
                </div>
                  </div>
                  <div className="col-md-4">
                  <h4 className={contentInvoiceStyles.invoiceNum}><strong># {this.state.selectedExpense.RequestorID}</strong></h4>
                  </div>
                </div>
                 
               
                {/* <div className={contentInvoiceStyles.footeritem}>
                  <h4><strong># {this.state.selectedExpense.RequestorID}</strong></h4>
                </div> */}
                <div className={contentInvoiceStyles.Container}>
                  <h6 className={contentInvoiceStyles.h3}>OverAll Details</h6>
                </div>

                {/* <!-- This is OverAll Details Table  --> */}
                <table className={contentInvoiceStyles.table}>
                  <thead>
                    <tr className={contentInvoiceStyles.td}>
                      <th className={contentInvoiceStyles.th}><strong>Requestor Name</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Department</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Description</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Start Date</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>End Date</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Purchase Entity ($)</strong></th>
                    </tr>
                  </thead>

                  <tbody>
                    <tr>
                      <td className={contentInvoiceStyles.td}>{this.state.Creator}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.Department}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.ReportHeader}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.StartDate}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.EndDate}</td>
                      <td className={contentInvoiceStyles.td}>{this.state.TotalExpense}</td>
                    </tr>
                  </tbody>

                </table>
                 <br></br>

                {/* <!-- this is expense details  --> */}

                <div className={contentInvoiceStyles.Container}>
                  <h6 className={contentInvoiceStyles.h3}>Expense Details</h6>
                </div>
                <table className={contentInvoiceStyles.table}>
                  <thead>
                    <tr className={contentInvoiceStyles.td}>
                      <th className={contentInvoiceStyles.th}><strong>Expense Type</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Check In</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Check Out</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Travel type</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Start Mile</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>End Mile</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Description</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>$ Amount</strong></th>
                    </tr>
                  </thead>
                  {this.renderInvoiceTableData()}
                  
                  <tfoot>
                    <tr>
                      <td></td>
                      <td></td>
                      <td></td>
                      <td></td>
                      <td></td>
                      <td></td>
                      <td className={contentInvoiceStyles.total}>Total:</td>
                      <td className="total">$ {this.state.TotalExpense}</td>
                    </tr>
                  </tfoot>
                </table>
                <br></br>
                {this.state.ILogHistoryModel.length > 0 &&
                <span>
                <div className={contentInvoiceStyles.Container2}>
                  <h6 className={contentInvoiceStyles.h3}>Comment History</h6>
                </div>    
                 
                {/* <!-- This is OverAll Details Table  --> */}
               
                <table className={contentInvoiceStyles.table2}>
                  <thead>
                    <tr className={contentInvoiceStyles.td}>
                      <th className={contentInvoiceStyles.th}><strong>Comment History</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Name</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Status</strong></th>
                      <th className={contentInvoiceStyles.th}><strong>Date</strong></th>
                    </tr>
                  </thead>
                  {this.renderInvoiceCommentTableData()}
                </table>
                </span>
                }
                {/* <!-- Footer Item --> */}
                <div className={contentInvoiceStyles.footeritem}>
                  <h6>Submitted By</h6>
                  <p>{this.state.Creator}</p>
                  <h6>Approved By</h6>
                  <p>{this.state.Manager}</p>
                </div>
              </div>
            </div>
          </Modal>
        }
          {this.state.openStatusBarDialog &&
          <Modal isOpen={this.isStatusBarModalOpen()} isBlocking={false} containerClassName={contentStatusBarStyles.container}>
              <div className={contentInvoiceStyles.header}>
              <span className={styles.label}> Request Id - {this.state.selectedExpense.RequestorID}</span>
                <IconButton styles={iconButtonStyles} iconProps={cancelIcon} ariaLabel="Close popup modal" onClick={this.hideStatusBarModal} />
              </div>
              <div className={contentStatusBarStyles.progressbar} id="statuscircleId">
                  <div className={[contentStatusBarStyles.statuscircle, contentStatusBarStyles.approved].join(" ")}>
                    <i className= {`${contentStatusBarStyles.icon} fa fa-check`}></i>
                    {this.state.RequestorResponse=="Yes" && <i className={`${contentStatusBarStyles.mailIcon} fa fa-envelope`} title='Email sent' aria-hidden="true"></i>}
                    <h5 className={contentStatusBarStyles.h5}>Requestor</h5><br></br>
                    <h4 className={contentStatusBarStyles.h4}>{this.state.Status=="Draft"?"Draft":"Submitted"}</h4><br></br>
                    <h4 className={contentStatusBarStyles.h4}>({this.state.Creator})</h4>
                  </div>
                  {this.state.ManagerStatus!="Manager Approval Not Required" &&
                  <div className={managerCss}>
                    <i className={`${contentStatusBarStyles.icon} fa fa-check`}></i>
                    {this.state.ManagerResponse=="Yes" && <i className={`${contentStatusBarStyles.mailIcon} fa fa-envelope`} title='Email sent' aria-hidden="true"></i>}
                    <h5 className={contentStatusBarStyles.h5}>Manager</h5>
                   <br></br>
                    <h4 className={contentStatusBarStyles.h4}>{this.state.ManagerStatus}</h4><br></br>
                    {this.state.Manager !="" && <h4 className={contentStatusBarStyles.h4}>({this.state.Manager})</h4>}
                  </div>
                  }
                  {this.state.FinanceStatus!=null  &&
                  <div className={financeCss}>
                    <i className={`${contentStatusBarStyles.icon} fa fa-check`}></i>
                    <div className={contentStatusBarStyles.Paid}>
                    {this.state.FinanceResponse=="Yes" && <i className={`${contentStatusBarStyles.mailIcon} fa fa-envelope`} title='Email sent' aria-hidden="true"></i>}
                      <h5 className={contentStatusBarStyles.h5}>Finance</h5>
                      <br></br>
                      <h4 className={contentStatusBarStyles.h4}>{this.state.FinanceStatus}</h4><br></br>
                      {this.state.Finance !="" && <h4 className={contentStatusBarStyles.h4}>({this.state.Finance})</h4>}
                    </div>
                    
                  </div>
                 }
            </div>

          </Modal>
        }
      </div>
    );
  }
}
