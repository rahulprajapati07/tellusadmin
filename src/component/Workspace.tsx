import * as React from 'react';
import * as ReactIcons from '@fluentui/react-icons-mdl2';
import { TooltipHost, mergeStyles } from '@fluentui/react';
import styles from "./CollaborationWorkspace.module.scss";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import { mergeStyleSets, SelectionMode, TextField } from "@fluentui/react";
import { IColumnConfig } from "fluentui-editable-grid"; 
import "office-ui-fabric-core/dist/css/fabric.min.css";
import InactiveIconTeams from "../Icons/InactiveIconTeams.png";
import ExtUsersIcon from "../Icons/ExtUsersIcon.png";
import NoOwnersIcon from "../Icons/NoOwnersIcon.png";
import TeamsMissingIcon from "../Icons/TeamsMissingIcon.png";
import LockIcon from "../Icons/LockIcon.png";
import InfoIcon from "../Icons/InfoIcon.jpg";
import sharepointImg from "../Icons/sharepointImg.svg";
// import sortIcon from '../Icon/sortIcon.png';
import sortIcon from '../Icons/sortIcon.png';
import Filter from '../Icons/Filter.png';
import { EditableGrid, EventEmitter, EventType } from "fluentui-editable-grid";
import {getClientDetails} from './BackendService';
//import jwtDecode from "jwt-decode";
import {
  Dialog,
  DialogType,
  DialogFooter,
} from "office-ui-fabric-react/lib/Dialog";
import {
  PrimaryButton,
  DefaultButton,
} from "office-ui-fabric-react/lib/Button";
import { getId } from "office-ui-fabric-react/lib/Utilities";
import { hiddenContentStyle } from "office-ui-fabric-react/lib/Styling";
import "../component/TellusAdmin.scss";
import {
  Button,
  IconButton
} from "office-ui-fabric-react/lib/Button";
import {
  HoverCard,
  HoverCardType,
  IPlainCardProps,
} from "office-ui-fabric-react/lib/HoverCard";
import {
  IContextualMenuProps,
  ContextualMenu,
} from "office-ui-fabric-react/lib/ContextualMenu"; //DirectionalHint,
import { callGetPublicTeams, canUserRestoreTeams, deleteWorkspace, archiveWorkspace } from "./BackendService";
import ReactTooltip from "react-tooltip";
import * as microsoftTeams from "@microsoft/teams-js";

const screenReaderOnly = mergeStyles(hiddenContentStyle);
const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: "16px",
  },
  fileIconCell: {
    textAlign: "center",
    selectors: {
      "&:before": {
        content: ".",
        display: "inline-block",
        verticalAlign: "middle",
        height: "100%",
        width: "0px",
        visibility: "hidden",
      },
    },
  },
  fileIconImg: {
    verticalAlign: "middle",
    maxWidth: "24px",
  },
  fileIconImgSharepoint: {
    verticalAlign: 'middle',
    maxWidth: '22px'
    },
  controlWrapper: {
    display: "flex",
    flexWrap: "wrap",
    padding:"20px 0"
  },
  workspaceImage: {
    width: "36px",
    height: "36px",
  },
  exampleToggle: {
    display: "inline-block",
    marginBottom: "10px",
    marginRight: "30px",
  },
  selectionDetails: {
    marginBottom: "20px",
  },
});


const icons = Object.keys(ReactIcons).reduce((acc: React.FC[], exportName) => {
  if ((ReactIcons as any)[exportName]?.displayName) {
    if (exportName === "MoreVerticalIcon") {
      acc.push((ReactIcons as any)[exportName] as React.FunctionComponent);
    }
  }
  return acc;
}, []);

export interface IWorkspaceExampleState {
  columns: IColumnConfig[];
  workspaceItemList : IWorkspace[];
  sortItemCheck: boolean;
  userIsAdmin: string;
  contextualMenuProps?: IContextualMenuProps;
  today: Date;
  inActiveCount: number;
  itemWithNoOwner: number;
  teamsMissingInfo: number;
  teamsExternalUser: number;
  currentItem: any;
  hideDialog: boolean;
  isDraggable: boolean;
  dialog: any;
  showSpinner : boolean;
  currentUserEmail : string;
  teamsMode : string;
  isSort : boolean;
}

export interface IWorkspace {
  key: string;
  test: string;
  name: string;
  businessDepartment: string;
  status: string;
  type: string;
  classification: string;
  businessOwner: string;
  teamsWithNoOwner: number;
  teamsExternalUser: number;
  teamsSiteUrl: string;
  sharePointSiteUrl: string;
  teamsGroupId: string;
}

interface IWorkspaceProps {
  userIsAdmin: any;
}

class WorkspaceDetails extends React.Component<
  IWorkspaceProps,
  IWorkspaceExampleState
  > {
  constructor(props: IWorkspaceProps, state: IWorkspaceExampleState) {
    super(props);
    let today = new Date();

    // EditableGrid columns Details 

    const columns: IColumnConfig[] = [
      {
        key: "test",
        name: "test",
        text: "",
        className: classNames.fileIconCell,
        iconClassName: classNames.fileIconHeaderIcon,
        ariaLabel: "Column operations for File type, Press to sort on File type",
        iconName: "Page",
        isIconOnly: true,
        fieldName: "name",
        minWidth: 20,
        maxWidth: 20,
        onRender: (item: IWorkspace) => (
          <div className="test">
            {
              item.test ?
              <TooltipHost key={item.key}>
              <img
                src={item.test}
                className={classNames.fileIconImg}
                alt={`${item.test} file icon`}
              />
              </TooltipHost>
              : 
              <IconButton
                iconProps={{ iconName: "ProgressLoopOuter" }}
                title="Request Details"
                className={[styles.workspaceImage, styles.requestImage].join(' ')}
              />
            }
          </div>
        ),
      },
      {
        key: "name",
        name: "Name",
        text: "Name",
        fieldName: "name",
        minWidth: 200,
        maxWidth: 200,
        dataType: "string",
        includeColumnInExport: true,
        includeColumnInSearch: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        data: "string",
        //onColumnClick : (ev: React.MouseEvent<HTMLElement>) => this.onColumnClick(ev),
        onRender: (item: IWorkspace) => {
          return (
            <div className="test">
              {" "}
              <span onClick={() => item.teamsSiteUrl ? window.open(item.teamsSiteUrl, "_blank") : null}>
                {" "}
                {item.name}{" "}
              </span>{" "}
            </div>
          );
        },
        isPadded: true,
      },
      {
        key: "column3",
        text: "",
        name: "",
        fieldName: "Options",
        minWidth: 20,
        maxWidth: 20,
        onRender: (item: IWorkspace) => {
          if (item.status !== "Awaiting Approval") {
            const plainCardProps: IPlainCardProps = {
              onRenderPlainCard: this.onRenderPlainCard,
              renderData: item,
            };
            return (
              <div className="test">
                <HoverCard
                sticky
                  plainCardProps={plainCardProps}
                  instantOpenOnClick={true}
                  type={HoverCardType.plain}
                  shouldBlockHoverCard={() => true}
                >
                  {icons.map(
                    (Icon: React.FunctionComponent<ReactIcons.ISvgIconProps>) => (
                      <Icon
                        key={item.key}
                        aria-label={"MoreVertical"?.replace("", "")}
                      />
                    )
                  )}
                </HoverCard>
              </div>
            );
          }
        },
      },
      {
        key: "businessDepartment",
        name: "businessDepartment",
        text: "Business Department",
        editable: true,
        dataType: "string",
        minWidth: 180,
        maxWidth: 180,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        applyColumnFilter: true,
        onRender: (item: IWorkspace) => {
          return (
            <div className="test">
              {" "}
              <span key={item.key}>{item.businessDepartment}</span>{" "}
            </div>
          );
        },
      },
      {
        key: "businessOwner",
        name: "businessOwner",
        text: "Business Owner",
        editable: true,
        dataType: "string",
        minWidth: 140,
        maxWidth: 140,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        applyColumnFilter: true,
      },
      {
        key: "status",
        name: "status",
        text: "Status",
        editable: true,
        dataType: "string",
        minWidth: 120,
        maxWidth: 120,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        applyColumnFilter: true,
      },
      {
        key: "type",
        name: "type",
        text: "Type",
        editable: true,
        dataType: "string",
        minWidth: 120,
        maxWidth: 120,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        applyColumnFilter: true,
      },
      {
        key: "classification",
        name: "classification",
        text: "Classification",
        editable: true,
        dataType: "string",
        minWidth: 100,
        maxWidth: 100,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        applyColumnFilter: true,
      },
      {
        key: "column9",
        name: "test",
        text: "",
        className: classNames.fileIconCell,
        iconClassName: classNames.fileIconHeaderIcon,
        ariaLabel: "Column operations for File type, Press to sort on File type",
        iconName: "Page",
        isIconOnly: true,
        fieldName: "name",
        minWidth: 20,
        maxWidth: 20,
        onRender: (item: IWorkspace) => (
          item.sharePointSiteUrl ?
          <div className="test">
            <TooltipHost key={item.key}>
              <img
                onClick={() => window.open(item.sharePointSiteUrl, "_blank")}
                src={sharepointImg}
                className={classNames.fileIconImgSharepoint}
                alt={`${item.test} file icon`}
              />
            </TooltipHost>
          </div>
          : 
          null
        ),
      },
    ];


    // Initialize the Workspace State 
    this.state = {
      workspaceItemList : [],
      columns: columns,
      contextualMenuProps: undefined,
      sortItemCheck: true,
      userIsAdmin: '',
      dialog: "none",
      today: today,
      inActiveCount: 0,
      itemWithNoOwner: 0,
      teamsMissingInfo: 0,
      teamsExternalUser: 0,
      currentItem: {},
      hideDialog: true,
      isDraggable: false,
      showSpinner : true,
      currentUserEmail : '',
      teamsMode : 'default',
      isSort : true
    };

    microsoftTeams.initialize();

    // Get The Microsoft Teams context And Theme
    microsoftTeams.getContext((context : any) => {

      let teamsContext = context.theme.toString();
      this.setState({
        teamsMode : teamsContext
      });
      let userEmail = context.userPrincipalName;

      this.setState({
        currentUserEmail : userEmail
      });
    });


    // Bind The method's 
    this.checkMode = this.checkMode.bind(this);
    this.onRenderPlainCard = this.onRenderPlainCard.bind(this);
    this.onContextualMenuDismissed = this.onContextualMenuDismissed.bind(this);
    this.renderEditDialog = this.renderEditDialog.bind(this);
    this.renderDialogDelete = this.renderDialogDelete.bind(this);
    this.addClickEvent = this.addClickEvent.bind(this);
    this._updateWorkspaces = this._updateWorkspaces.bind(this);
    this.renderDialog = this.renderDialog.bind(this);
    this._getAccessToken = this._getAccessToken.bind(this);
    this.sortColumn = this.sortColumn.bind(this);
  }

  private _labelId: string = getId("dialogLabel");
  private _subTextId: string = getId("subTextLabel");

  private _dragOptions = {
    moveMenuItemText: "Move",
    closeMenuItemText: "Close",
    menu: ContextualMenu,
  };

  // Dismissed The ContextualMenu
  private onContextualMenuDismissed = (): void => {
    this.setState({
      contextualMenuProps: undefined,
    });
  };

  // Render the PlainCard for the Three Dotes
  private onRenderPlainCard(item: any): JSX.Element {
    return (
      <div className={styles.block + ' elliptical-menu'}>
        {/* edit */}
        <Button
          text="Edit"
          className={styles.createNewButton}
          onClick={() => this.setState({
            currentItem: item,
            dialog: "Update",
          })}
        />
        <br />

        {/* archive */}
        <Button
          text={item.status === "Archived" ? "Unarchive" : "Archive"}
          className={styles.createNewButton}
          onClick={() =>
            this.setState({
              currentItem: item,
              dialog: item.status === "Archived" ? "Unarchive" : "Archive",
            })
          }
        />
        <br />

        {/* delete */}
        <Button
          text="Delete"
          className={styles.createNewButton}
          onClick={() => this.setState({ currentItem: item, dialog: "Delete" })}
        />

        {/* Dialog popup for both archive and delete (e.g. are you sure you want to delete?) */}
        {(this.state.dialog === "Update") ? this.renderEditDialog(item) : (this.state.dialog === "Delete") ? this.renderDialogDelete(item) : (this.state.dialog === "none") ? () => { } : this.renderDialog(item)}
        
      </div>
    );
  }


  // Call The Initilize API For App 
  public async componentDidMount() {

    // This Function Check the user is Admin Or not and set the userRole.
    await this._getUserRole().then((teamsUserRoleStatus: boolean) => {
      if (teamsUserRoleStatus === true) {
        //userRole = teamsUserRoleStatus;
        this.setState({
          userIsAdmin: "true", // true
        });
        console.log("Teams User Role status : " + this.state.userIsAdmin);
      } else {
        //userRole = teamsUserRoleStatus;
        this.setState({
          userIsAdmin: "false",
        });
      }
    });

    // This Function is Get the Teams Details From The API And Apply the 4 Box Count.
    await this._getAllPublicTeams().then((teamsDetails: any[]) => {
      
      let countWorkspaceWithNoOwner = 0;
      let countMissiongInformation = 0;
      let countExternalUser = 0;

      for (let i = 0; i < teamsDetails.length; i++) {
        if (teamsDetails[i].teamsWithNoOwner === 0) {
          countWorkspaceWithNoOwner = countWorkspaceWithNoOwner + 1;
        }
        if (teamsDetails[i].teamsExternalUser > 0) {
          countExternalUser = countExternalUser + 1;
        }
        if (
          teamsDetails[i].businessOwner === "" ||
          teamsDetails[i].businessDepartment === "" ||
          teamsDetails[i].classification === "" ||
          teamsDetails[i].type === ""
        ) {
          countMissiongInformation = countMissiongInformation + 1;
        }
      }

      // Set the by default accending order in title column
      teamsDetails = teamsDetails.sort((a, b) => a.name.localeCompare(b.name));

      // set the workspace Information and Count
      this.setState({
        workspaceItemList: teamsDetails,
        itemWithNoOwner: countWorkspaceWithNoOwner,
        teamsMissingInfo: countMissiongInformation,
        teamsExternalUser: countExternalUser,
      });

      // Change the Export Location 
      var exp: any = document.getElementById("export");
      document
        .getElementsByClassName("ms-TextField-wrapper")[0]
        .appendChild(exp);

      // change the table size after data fill and change the table width.

      if (document.querySelectorAll('.ms-DetailsList-contentWrapper .ms-ScrollablePane')) {
        var gridHeight: any = document.querySelectorAll('.ms-DetailsList-contentWrapper .ms-ScrollablePane')[0];
        // var HeightUnset: any = document.querySelectorAll('.ms-DetailsList-contentWrapper .ms-Fabric div:nth-child(2)')[0];
        let parentNodeOfScrollPane: any = gridHeight.parentElement
        parentNodeOfScrollPane.style.height = "61vh";
      }
    });
    // bind the click event with componentdidmount for column click.
    this.addClickEvent();
  }

  // Set the teams App by theme 
  public checkMode() {
    let bodyEle: any = document.querySelectorAll('html')[0];
    bodyEle.className = this.state.teamsMode;
  }

  // Excecute when the click on columns.
  public addClickEvent() {
    const that = this;

    // Column tooltip Hide
    let Columnhover: any = document.querySelectorAll('.ms-DetailsHeader-cell')
    Columnhover.forEach((element: any) => {
      element.addEventListener('mouseenter', (ev: Event) => {
        let hideTooltip = () => {
          if (document.querySelectorAll('.ms-Tooltip-subtext')[0] || document.querySelectorAll('.ms-Tooltip')[0]) {
            let tooltTipColumn: any = document.querySelectorAll('.ms-Tooltip-subtext')[0]
            let parentTooltip: any = tooltTipColumn.closest('.ms-Tooltip')
            parentTooltip.style.display = "none"
            clearInterval(intervalHide);
          }
        }
        let intervalHide = setInterval(() => { hideTooltip(); }, 100);
      });
    });
    if(document.querySelectorAll('.ms-Tooltip-subtext')[0]){
      let tooltipColumn: any = document.querySelectorAll('.ms-Tooltip-subtext')[0]
      let parentTooltip: any = tooltipColumn.closest('.ms-Tooltip')[0]
      parentTooltip.style.display = 'none'
    }
    
    // Get the HTML body 
    let bodyEle: any = document.querySelectorAll('html')[0]
    bodyEle.className = this.state.teamsMode;
    
    // Sort Icon for Name 
    const sortIconClass = document.createElement("img");
    sortIconClass.src = sortIcon;
    sortIconClass.className = 'SortClass';

    // append the sort class on Name 
    let columnName: any = document.querySelectorAll('.ms-DetailsHeader-cellTitle')[1];
    columnName.appendChild(sortIconClass);

    // Filter Icon For Business Department
    const sortIconBD = document.createElement("img");
    sortIconBD.src = Filter;
    sortIconBD.className =  'SortClassBD';

    // append the sort class on Business Department
    let getbusinessDepartment : any = document.querySelectorAll('.ms-DetailsHeader-cellTitle')[3];
    getbusinessDepartment.appendChild(sortIconBD);

    // Filter Icon For Business Owner
    const sortIconOwner = document.createElement("img");
    sortIconOwner.src = Filter;
    sortIconOwner.className =  'SortClassOwner';

    let getbusinessOwner : any = document.querySelectorAll('.ms-DetailsHeader-cellTitle')[4];
    getbusinessOwner.appendChild(sortIconOwner);

    // Filter Icon For Status
    const sortIconStatus = document.createElement("img");
    sortIconStatus.src = Filter;
    sortIconStatus.className =  'SortClassStatus';

    let getbusinessStatus : any = document.querySelectorAll('.ms-DetailsHeader-cellTitle')[5];
    getbusinessStatus.appendChild(sortIconStatus);

    // Filter Icon For Type
    const sortIconType = document.createElement("img");
    sortIconType.src = Filter;
    sortIconType.className =  'SortClassType';

    let getbusinessType : any = document.querySelectorAll('.ms-DetailsHeader-cellTitle')[6];
    getbusinessType.appendChild(sortIconType);

    // Filter Icon For Classification
    const sortIconClassification = document.createElement("img");
    sortIconClassification.src = Filter;
    sortIconClassification.className =  'SortClassClassification';
   
    let getbusinessClassification : any = document.querySelectorAll('.ms-DetailsHeader-cellTitle')[7];
    getbusinessClassification.appendChild(sortIconClassification);
    
    // SortIcon Change on click name column
    var rottaeDeg = 90;
    columnName.addEventListener('click', function (ev: Event) {
      
      // ev.stopPropagation();
      if (document.querySelectorAll('.SortClass')) {
        let sortImg: any = document.querySelectorAll('.SortClass')[0];
        sortImg.style.transform = `rotate(${rottaeDeg}deg)`
        rottaeDeg += 180
      }
    });

    // Export ToolTip.
    const tooltipDiv: any = document.createElement("div");
    tooltipDiv.className = 'exportTooltip';
    tooltipDiv.textContent = 'Export list of All Teams';
    if (document.getElementById('export')) {
      var exp: any = document.getElementById('export');    
      exp.append(tooltipDiv);
    }

    if (
        document.getElementById('export')
        ) {
        var exportDocumnet: any = document.getElementById('export');
        exportDocumnet.setAttribute('title', 'Export Teams');

    }

    // sort the Business Owner 
    let sortOwner: any = document.querySelectorAll('.ms-DetailsHeader-cell')[4]
    sortOwner.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      console.log(currentEvent);
      //this.sortColumn(false);
    });

    // sort the Business Status 
    let sortStatus: any = document.querySelectorAll('.ms-DetailsHeader-cell')[5]
    sortStatus.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      console.log(currentEvent);
      //this.sortColumn(false);
    });

    // sort the Business Type 
    let sortType: any = document.querySelectorAll('.ms-DetailsHeader-cell')[6]
    sortType.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      console.log(currentEvent);
      //this.sortColumn(false);
    });

    // Get the column click  
    let sortClassification: any = document.querySelectorAll('.ms-DetailsHeader-cell')[1]
    sortClassification.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      console.log(currentEvent);
      let columnsName = "name";
      this.sortColumn(true,columnsName);
    });

    // get the column click for Business Department 
    let onClickBD : any = document.querySelectorAll('.ms-DetailsHeader-cellName')[3]
    //let targetClass = onClickBD.target.className;
    //let targetClasss = onClickBD;
    onClickBD.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      let cellName = "businessDepartment";
      this.sortColumn(true, cellName);
      currentEvent.stopPropagation();
      console.log(currentEvent);
    });

    // get the column click for Business Owner 
    let onClickBO : any = document.querySelectorAll('.ms-DetailsHeader-cellName')[4]
    onClickBO.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      let cellName = "businessOwner";
      this.sortColumn(true, cellName);
      currentEvent.stopPropagation();
      console.log(currentEvent);
    });

    // get tje column click for Status
    let onClickStatus : any = document.querySelectorAll('.ms-DetailsHeader-cellName')[5]
    onClickStatus.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      let cellName = "status";
      this.sortColumn(true, cellName);
      currentEvent.stopPropagation();
      console.log(currentEvent);
    });

    // get tje column click for Type
    let onClickType : any = document.querySelectorAll('.ms-DetailsHeader-cellName')[6]
    onClickType.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      let cellName = "type";
      this.sortColumn(true, cellName);
      currentEvent.stopPropagation();
      console.log(currentEvent);
    });

    // get tje column click for classification
    let onClickClassification : any = document.querySelectorAll('.ms-DetailsHeader-cellName')[7]
    onClickClassification.addEventListener('click',  (ev: Event) => {
      let currentEvent = ev;
      let cellName = "classification";
      this.sortColumn(true, cellName);
      currentEvent.stopPropagation();
      console.log(currentEvent);
    });
    

    let testArr: any = document.querySelectorAll('.ms-DetailsHeader-cell')
    testArr.forEach((element: any) => {
      element.addEventListener('click',  (ev: any) => {

        let checkPopup = () => {
          if (document.querySelectorAll("div[role='filtercallout'] .ms-Button .ms-Button-label") &&
            document.querySelectorAll("div[role='filtercallout'] .ms-Button .ms-Button-label")[1]) {
            
            // apply css for filter dilogbox. 
              that.applyCustomCSS();
            clearInterval(test1);
          }
        };
        let test1 = setInterval(() => { checkPopup(); }, 100);
      });
      

      // Hide sorting arrow on column right click
      element.addEventListener('contextmenu', (ev: Event) => {
        let hideICon = () => {
          let iconSOrtHide: any = document.querySelectorAll('.ms-DetailsHeader-cellTitle .ms-Icon')[0];
          iconSOrtHide.style.display = 'none'
          clearInterval(hideicnn);
        };
        let hideicnn = setInterval(() => { hideICon(); }, 1);
        ev.preventDefault()
      });
    })
  }


  // Sort the workspace on all Column
  // give parametr for column name 
  private sortColumn (currentSort : boolean, checkColumnClass : string = "" )
  {
    let currentWorkspaces:any[] = [...this.state.workspaceItemList];
    if(currentSort){

      if(this.state.isSort === true && checkColumnClass !== "")
      {
        let decWorkspaces = currentWorkspaces.sort((a, b) => b[checkColumnClass].localeCompare(a[checkColumnClass]));
        
        this.setState({
            workspaceItemList : decWorkspaces,
            isSort : false
        });
      }
      else
      {
        let accWorkspaces = currentWorkspaces.sort((a, b) => a[checkColumnClass].localeCompare(b[checkColumnClass]));
        this.setState({
            workspaceItemList : accWorkspaces,
            isSort : true
        });
      }
    }

    else 
    {
      this.setState({
        workspaceItemList : currentWorkspaces
      });
    }
  }

  // This function useed for Apply Filter Dilog Box Css 
  public applyCustomCSS() {

    // change text clear All to clear

    if (document.querySelectorAll("div[role='filtercallout'] .ms-Button .ms-Button-label") &&
      document.querySelectorAll("div[role='filtercallout'] .ms-Button .ms-Button-label")[1]) {
      document.querySelectorAll(
        "div[role='filtercallout'] .ms-Button .ms-Button-label"
      )[1].textContent = 'Clear';
    }

    //descrease Padding of filter box

    if (
      document.querySelectorAll('div[role="filtercallout"]') &&
      document.querySelectorAll('div[role="filtercallout"]')[0]
    ) {
      var filterPadding: any = document
        .querySelectorAll('div[role="filtercallout"]')[0]
        .closest('.ms-Callout')
      filterPadding.style.padding = '13px'
    }

    //change filter search textbox placeholder
    if (
      document.querySelectorAll('.ms-TextField-field') &&
      document.querySelectorAll('.ms-TextField-field')[1]
    ) {
      var placeHolderSearch: any = document.querySelectorAll(
        '.ms-TextField-field'
      )[1]
      placeHolderSearch.setAttribute('placeholder', 'Search')
    }
  }

  // Rendert The HTML For the Application 
  public render() {

    return (
      <div className="container-custom">
        {this.state.userIsAdmin === "true" ? (
          <div className="ms-Grid" style={{marginTop:'30px'}} dir="ltr">
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                <div
                  className="white-wrapper"
                  style={{
                    backgroundColor: "#FFFFFF",
                    margin: '0px 0px 0px 0px ',
                    padding: 2,
                  }}
                >
                  <div className="title_wrapper">
                    <h6
                      style={{
                        textAlign: "left",
                        margin: "10px 15px 0px 15px",
                        fontFamily: "Segoe UI",
                      }}
                    >
                      {" "}
                      Inactive Teams
                    </h6>
                    <div
                      style={{ marginLeft: "0px" }}
                      data-tip="Total Teams with Inactive Status"
                    >
                      <img
                        width="10.0"
                        src={InfoIcon}
                        alt="new"
                      />
                      <ReactTooltip></ReactTooltip>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <h3
                        style={{
                          textAlign: "left",
                          margin: "10px 15px 0px 13px",
                          fontSize: 36,
                        }}
                      >
                        {" "}
                        
                        {this.state.showSpinner ? null : this.state.inActiveCount}

                          { this.state.showSpinner ? this.renderSpinner(
                            "",
                            SpinnerSize.large,
                            "right"
                          ) : null}
                      </h3>
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <img src={InactiveIconTeams} alt="new" />
                    </div>
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                <div
                  className="white-wrapper"
                  style={{
                    backgroundColor: "#FFFFFF",
                    margin: "0px 0px 0px 0px ",
                    padding: 2,
                  }}
                >
                  <div className="title_wrapper">
                    <h6
                      style={{
                        textAlign: "left",
                        margin: "10px 15px 0px 15px",
                        fontFamily: "Segoe UI",
                      }}
                    >
                      {" "}
                      Teams with no Owner
                    </h6>
                    <div data-tip="Total Teams with no owner's">
                      <img
                        width="10.0"
                        src={InfoIcon}
                        alt="new"
                      />
                    </div>
                    <ReactTooltip></ReactTooltip>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <h3
                        style={{
                          textAlign: "left",
                          margin: "10px 15px 0px 13px",
                          fontSize: 36,
                        }}
                      >
                        {" "}
                        {this.state.showSpinner ? null : this.state.itemWithNoOwner}

                        { this.state.showSpinner ? this.renderSpinner(
                          "",
                          SpinnerSize.large,
                          "right"
                        ) : null}
                      </h3>
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <img src={NoOwnersIcon} alt="new" />
                    </div>
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                <div
                  className="white-wrapper"
                  style={{
                    backgroundColor: "#FFFFFF",
                    margin: "0px 0px 0px 0px ",
                    padding: 2,
                  }}
                >
                  <div className="title_wrapper">
                    <h6
                      style={{
                        textAlign: "left",
                        margin: "10px 15px 0px 15px",
                        fontFamily: "Segoe UI",
                      }}
                    >
                      {" "}
                      Teams with external user
                    </h6>
                    <div data-tip="Total Teams with external user's">
                      <img
                        width="10.0"
                        src={InfoIcon}
                        alt="new"
                      />
                    </div>
                    <ReactTooltip></ReactTooltip>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <h3
                        style={{
                          textAlign: "left",
                          margin: "10px 15px 0px 13px",
                          fontSize: 36,
                        }}
                      >
                        {" "}
                        {this.state.showSpinner ? null : this.state.teamsExternalUser}

                        { this.state.showSpinner ? this.renderSpinner(
                          "",
                          SpinnerSize.large,
                          "right"
                        ) : null}
                      </h3>
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <img src={ExtUsersIcon} alt="new" />
                    </div>
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                <div
                  className="white-wrapper"
                  style={{
                    backgroundColor: "#FFFFFF",
                    margin: "0px 0px 0px 0px ",
                    padding: 2,
                  }}
                >
                  <div className="title_wrapper">
                    <h6
                      style={{
                        textAlign: "left",
                        margin: "10px 15px 0px 15px",
                        fontFamily: "Segoe UI",
                      }}
                    >
                      {" "}
                      Teams missing information
                    </h6>
                    <div data-tip="Total Teams missing information">
                      <img
                        width="10.0"
                        src={InfoIcon}
                        alt="new"
                      />
                    </div>
                    <ReactTooltip></ReactTooltip>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <h3
                        style={{
                          textAlign: "left",
                          margin: "10px 15px 0px 13px",
                          fontSize: 36,
                        }}
                      >
                        {" "}
                        {this.state.showSpinner ? null : this.state.teamsMissingInfo}

                        { this.state.showSpinner ? this.renderSpinner(
                          "",
                          SpinnerSize.large,
                          "right"
                        ) : null}
                      </h3>
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <img src={TeamsMissingIcon} alt="new" />
                    </div>
                  </div>
                </div>
              </div>
            </div>
            {/* Render table */}
            <div
                className='darkGird'
                style={{
                margin: '15px 0',
                backgroundColor: '#FFFFFF',
                boxShadow: '1px 2px 7px #0000000f',
                borderRadius: '5px',
                paddingBottom: '25px'
                }}
            >
              <div className="ms-Grid-row" style={{ marginTop: 20 }}>
              </div>
              {/* showing the search teams section */}
              <div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">
                  </div>
                </div>
              </div>
              {this.state.contextualMenuProps && (
                <ContextualMenu {...this.state.contextualMenuProps} />
              )}
              
              {/* This Renders the Teams Records */}
              <div className="ms-Grid-row">
                <div
                  className="ms-Grid-col ms-sm6 ms-md4 ms-lg12"
                  
                >
                  <div className={classNames.controlWrapper}>
                  
                    <TextField
                      placeholder="Search Teams"
                      className={mergeStyles({
                        width: "60vh",
                        paddingBottom: "10px",
                      })}
                      onChange={(event) =>
                        EventEmitter.dispatch(EventType.onSearch, event)
                      }
                    />
                  </div>
                  <div className="ms-DetailsList-contentWrapper">
                      <EditableGrid
                      id={1}
                      columns={this.state.columns}
                      items={this.state.workspaceItemList}
                      enableExport={true}
                      width={"140vh"}
                      enableColumnFilters={true}
                      selectionMode={SelectionMode.none}
                      />
                  </div>
                  <label id={this._labelId} className={screenReaderOnly}>
                    My sample Label
                  </label>
                  <label id={this._labelId} className={screenReaderOnly}>
                    My sample Label
                  </label>
                  <label id={this._subTextId} className={screenReaderOnly}>
                    My Sample description
                  </label>

                  <Dialog
                    hidden={this.state.hideDialog}
                    onDismiss={this._closeDialog}
                    dialogContentProps={{
                      type: DialogType.normal,
                      title: "All emails together",
                      subText:
                        "Your Inbox has changed. No longer does it include favorites, it is a singular destination for your emails.",
                    }}
                    modalProps={{
                      titleAriaId: this._labelId,
                      subtitleAriaId: this._subTextId,
                      isBlocking: false,
                      styles: { main: { maxWidth: 450 } },
                      dragOptions: this.state.isDraggable
                        ? this._dragOptions
                        : undefined,
                    }}
                  >
                    <DialogFooter>
                      <PrimaryButton onClick={this._closeDialog} text="Save" />
                      <DefaultButton
                        onClick={this._closeDialog}
                        text="Cancel"
                      />
                    </DialogFooter>
                  </Dialog>
                </div>
              </div>
              
            </div>
            {this.state.showSpinner ? this.renderSpinner("Loading",SpinnerSize.large,"right") : null}
          </div>
        ) : this.state.userIsAdmin === "false" ? (
          <div
            className="ms-Grid"
            dir="ltr"
            style={{
              display: "flex",
              justifyContent: "center",
              flexDirection: "column",
              alignItems: "center",
              height: "78vh",
            }}
          >
            <div className="ms-Grid-row" style={{ marginTop: 0 }}>

              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">

                <img height="140" width="140" src={LockIcon} alt="new" />

              </div>
            </div>

            <div className="ms-Grid-row">

              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">

                <h5 className="splash_title" style={{ margin: '0' }}> Sorry but you don't have access to this feature </h5>

              </div>

            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">
                <p className="splce_subtitle" style={{ fontFamily: "Segoe UI" }}>
                  {" "}
                  Tellus Admin is only available to administrators{" "}
                </p>
              </div>
            </div>
          </div>
        ) : <div> </div>}
      </div>
    );
  }

  // This function is used for the render the dilog box on delete Workspace
  private renderDialogDelete(item: any): JSX.Element {
    return (
      <Dialog
        hidden={this.state.dialog === "none"}
        onDismiss={() => this.closeDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: "Delete Team",
          subText: `Are you sure you want to delete this Team?`,
        }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={() => this.state.dialog === "Delete" ? this._deleteWorkspace(item) : this._archiveWorkspace(item)}
            text={'Delete'}
          />
          <DefaultButton
            onClick={() => this.closeDialog(false)}
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>
    );
  }

  // This Function is used for the render the dialog box on edit workspace
  private renderEditDialog(item: any): JSX.Element {
    <script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>
    return (
      <div className="dialogboxedit">
        <Dialog
          hidden={this.state.dialog === "none"}
          onDismiss={() => this.closeDialog(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: this.state.dialog + " Team",
          }}
        >
          <div className="dialogboxtext" >
            <label>Business Department</label>
            <TextField
              id="textTitle"
              name="Title"
              placeholder=""
            />
          </div>
          <div className="dialogboxtextfield dialogboxtext" >
            <label>Business Owner</label>
            <TextField
              id="textTitle"
              name="Title"
              placeholder=""
            />
          </div>

          <DialogFooter>
            <PrimaryButton
              text={this.state.dialog}
            />

            <DefaultButton
              onClick={() => this.closeDialog(false)}
              text="Cancel"
            />
          </DialogFooter>
        </Dialog>
      </div>
    );
  }

  // This FUnction is used for the reder loader in table and 4 boxes
  private renderSpinner(label:any, size:any, position:any): JSX.Element {
    return (
      <Spinner
        className={styles.spinner}
        label={label}
        size={size}
        labelPosition={position}
      />
    );
  }

  // This function render the dialog for Edit 
  private renderDialog(item: any): JSX.Element {
    return (
          <Dialog
            hidden={this.state.dialog === "none"}
            onDismiss={() => this.closeDialog(false)}
            dialogContentProps={{
              type: DialogType.normal,
              title: this.state.dialog + " Team",
              subText: `Are you sure you want to ${this.state.dialog.toLocaleLowerCase()} this Team?`,
            }}
          >

          <DialogFooter>
          <PrimaryButton
            onClick={() => this.state.dialog === "Delete" ? this._deleteWorkspace(item) : this._archiveWorkspace(item)}
            text={this.state.dialog}
          />
          <DefaultButton
            onClick={() => this.closeDialog(false) }
            text="Cancel"
          />
        </DialogFooter>
        </Dialog>
    );
  }

  // This Function is used for close the dialog box
  private closeDialog = (confirm: boolean): void => {
    this.onContextualMenuDismissed();
    this.setState({ dialog: "none" });
  };

// This Function is used for close the dialog box
  private _closeDialog = (): void => {
    this.onContextualMenuDismissed();
    this.setState({ hideDialog: true });
  };


  // Get all teams from the Azure Function 
  public _getAllPublicTeams = async (accessToken:string = ""): Promise<IWorkspace[]> => {
    return new Promise<any>(async (resolve, reject) => {  
      
      
      // this _getAccessToken method Genrates the access token for azure function API Call 
      accessToken = await this._getAccessToken();

      const items: IWorkspace[] = [];
      var today = this.state.today.getTime();
      
      let totalInActiveTeams = 0; // Count for the inActiveTeams In Tellus Admin.
      
      // Call Azure Function With Access Token 
      callGetPublicTeams(accessToken)
          .then((response) => response)
            .then((data: any[]) => {
              data.forEach((element) => {

                var daysSinceActivity = 0;
                
                // Logic for identity Teams is active or not 
                if (element.latestActivityDate != null) {
                  daysSinceActivity =
                    (today - new Date(element.latestActivityDate).getTime()) /
                    (1000 * 60 * 60 * 24.0);
                }
                
                // Condition true then team will Inactive Otherwise teams will Active

                if (element.status === "Active" && daysSinceActivity >= 97) {
                  
                  element.status = "Inactive";

                  totalInActiveTeams = totalInActiveTeams + 1;
                  items.push({
                    test: element.imageBlob,
                    key: element.id.toString(),
                    teamsSiteUrl: element.urlTeams,
                    sharePointSiteUrl: element.urlSharePoint,
                    name: element.title,
                    businessDepartment: element.businessDepartment,
                    status: element.status,
                    type: element.template,
                    classification: element.classification,
                    businessOwner: element.ownerName,
                    teamsExternalUser: element.teamsExternalUser,
                    teamsWithNoOwner: element.teamsOwner,
                    teamsGroupId: element.groupId,
                  });
                }
                else {
                items.push({
                  test: element.imageBlob,
                  key: element.id.toString(),
                  name: element.title,
                  teamsSiteUrl: element.urlTeams,
                  sharePointSiteUrl: element.urlSharePoint,
                  businessDepartment: element.businessDepartment,
                  status: element.status,
                  type: element.template,
                  classification: element.classification,
                  businessOwner: element.ownerName,
                  teamsExternalUser: element.teamsExternalUser,
                  teamsWithNoOwner: element.teamsOwner,
                  teamsGroupId: element.groupId,
                });
              }
              });

              // Set the count for InActive Teams 
              this.setState({
                inActiveCount: totalInActiveTeams,
                showSpinner : false
              });

              resolve(items);
            });
    });
  };


  // This function will delete the workspace from the item 
  public _deleteWorkspace = async (item: any): Promise<any> => {

    this.setState({
      dialog:"none"
    });

    return new Promise<any>( async (resolve, reject) => {

      // Get the access token for Azure Function API Call
      let accessToken = await this._getAccessToken();

      // Call Tellus Delete API using Access token and current Item. 
      deleteWorkspace(accessToken, item)
            .then(async (response: any) => {
              if (response.ok === true) {
                this._updateWorkspaces(item); // update the 4 box count teams delete
              }
            })
    });
  }

  // This function update the 4 box count and current itemList.
  public _updateWorkspaces = async (item : any) => {

    let inactiveTeamsCountTemp = this.state.inActiveCount;
    let teamsWithNoOwnerCountTemp = this.state.itemWithNoOwner;
    let teamsWithExternalUserCountTemp = this.state.teamsExternalUser;
    let teamsMissingInfoCountTemp = this.state.teamsMissingInfo;

    // Decrese the InActive count if Deleted Workspace Is InActive.
    if(item.status === "Inactive"){
      inactiveTeamsCountTemp = inactiveTeamsCountTemp - 1; 
    }
    // Decrese the MissingInfo count if Deleted Workspace Has Missing Info.
    if(item.businessOwner.trim() === "" || item.businessDepartment.trim() === "" || item.classification.trim() === "" || item.type.trim() === "" ){
      teamsMissingInfoCountTemp = teamsMissingInfoCountTemp - 1;
    }
    // Decrese the ExternalUser count if Deleted Workspace Has External User.
    if(item.teamsExternalUser > 0){
      teamsWithExternalUserCountTemp = teamsWithExternalUserCountTemp - 1;
    }
    // Decrese the TeamsOwner count if Deleted Workspace Has noOwner.
    if(item.teamsWithNoOwner === 0){
      teamsWithNoOwnerCountTemp = teamsWithNoOwnerCountTemp - 1; 
    }

    // Remove the Deleted Workspace and Set Workspaces And count. 
    var currentItemList = this.state.workspaceItemList;
    currentItemList.splice(currentItemList.indexOf(item),1);
    let updatedWorkspces = currentItemList;

    this.setState({
      workspaceItemList : []
    });

    this.setState({
            workspaceItemList: updatedWorkspces,
            dialog: "none",
            inActiveCount : inactiveTeamsCountTemp,
            teamsMissingInfo : teamsMissingInfoCountTemp,
            teamsExternalUser : teamsWithExternalUserCountTemp,
            itemWithNoOwner : teamsWithNoOwnerCountTemp,
            showSpinner : false
        },
        () => {
          this.forceUpdate();
        });
  }


  // this function will archived the workspace from the teams Item
  public _archiveWorkspace = async (item: any): Promise<any> => {
    return new Promise<any>( async (resolve, reject) => {

      // This function get the access token for the azure function API
      let accessToken = await this._getAccessToken();

      // call thr archive APi using Access Token and current Item
      archiveWorkspace(accessToken, item)
            .then(
              async (response: any) => {
                if (response.ok === true) {
                  let currentStatus : any = item.status;
                  if (currentStatus === "Archived")
                  {
                    currentStatus = "Active";
                  }
                  else
                  {
                    currentStatus = "Archived";
                  }
                  item.status = currentStatus;
                  let tempItem = item;
                  let tempWorkspaces = [...this.state.workspaceItemList];

                  tempWorkspaces.splice(tempWorkspaces.indexOf(tempItem),1);
                  tempWorkspaces.push(tempItem);

                  this.setState({
                    workspaceItemList : tempWorkspaces,
                    dialog : "none"
                  });
                }
              }
            )
            .then((data: any) => {
              resolve(data);
            })
    });
  }

  // this function will get the user role and retuen boolean expression if the response is true then user will Teams administrator 
  public _getUserRole = async (): Promise<boolean> => {
    return new Promise<boolean>( async (resolve, reject) => {
      console.log("Call The GetUserRole API :");

      let accessToken = await this._getAccessToken();
      
      canUserRestoreTeams(accessToken, this.state.currentUserEmail)
        .then((response) => response)
        .then((data: any) => {
          resolve(data);
        });
    });
  };

  // Genrates Access Token For the API 
  public _getAccessToken = async () : Promise<string> => {
    return new Promise<string>((resolve, reject) => {
      microsoftTeams.authentication.getAuthToken({
          
        successCallback: (token: string) => {

              microsoftTeams.appInitialization.notifySuccess();
              // call backEnd APi With Teams Token.    
              getClientDetails(token + "", this.state.currentUserEmail, "082a7423-5b17-4f5e-a4dc-6d2396d7edfa").then((graphToken) => {
                  
                resolve(graphToken as string);
              }).catch((err) => {
                  // console.log("Function : Error For Genrates Token :");
                  // console.log(err);
              })
          },
          
          failureCallback: (message: string) => {
              //setError(message);
              microsoftTeams.appInitialization.notifyFailure({
                  reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                  message
              });
          },
          resources:["api://ambitious-pebble-0b2637f10.1.azurestaticapps.net/b0785c01-bd69-4a12-bfe1-e558e7a4b7d1"]
        });
    });
  };
}

export default WorkspaceDetails;


// Token Decode
  //console.log("Function : Teams Token : " + token);
  //const decoded: { [key: string]: any; } = jwtDecode(token) as { [key: string]: any; };
  //setName(decoded!.name);