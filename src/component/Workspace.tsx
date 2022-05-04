import * as React  from 'react';
//import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import * as ReactIcons from '@fluentui/react-icons-mdl2';
import {  IColumn } from '@fluentui/react/lib/DetailsList'; //SelectionMode DetailsList,
import { TooltipHost , mergeStyles } from '@fluentui/react';
import { Panel } from '@fluentui/react/lib/Panel';
//import { useState } from 'react';
import { mergeStyleSets, SelectionMode, TextField } from '@fluentui/react'; //DetailsListLayoutMode, mergeStyles,DetailsListLayoutMode
import { IColumnConfig } from 'fluentui-editable-grid'; //, EventEmitter, EventType, NumberAndDateOperators, EditableGrid, EditControlType,
//import { Fabric, Checkbox } from 'office-ui-fabric-react';
import "office-ui-fabric-core/dist/css/fabric.min.css";
//import { Image } from '@fluentui/react/lib/Image';
//import WorkspaceDetailsList from '../component/editabledetailslist/gridworkspace';
import InactiveIconTeams from '../Icons/InactiveIconTeams.png';
import ExtUsersIcon from '../Icons/ExtUsersIcon.png';
import NoOwnersIcon from '../Icons/NoOwnersIcon.png';
import TeamsMissingIcon from '../Icons/TeamsMissingIcon.png';
import LockIcon from '../Icons/LockIcon.png';
import sharepointImg from '../Icons/sharepointImg.png';
import InfoIcon from '../Icons/InfoIcon.jpg';
import { EditableGrid , EventEmitter , EventType } from 'fluentui-editable-grid';
import {  Dialog } from '@fluentui/react-northstar'
// import {
//   Provider as TeamsProvider,
//   Table,
//   List,
//   TSortable
// } from "@fluentui/react-teams";

import  '../component/Pagination.scss';

//import { ContextualMenuCheckmarksExample } from '../component/ContextualMenuCheckmarksExample';
import {
  // IconButton,
  Button,
} from "office-ui-fabric-react/lib/Button";

import {
    HoverCard,
    HoverCardType,
    IPlainCardProps,
  } from "office-ui-fabric-react/lib/HoverCard";

  import {
    IContextualMenuProps,
    DirectionalHint,
    ContextualMenu,
    //IContextualMenuItem,
  } from 'office-ui-fabric-react/lib/ContextualMenu';

  //import { TextField } from '@fluentui/react/lib/TextField';
//  import { Label } from '@fluentui/react/lib/Label';
  import { loginRequest } from "../component/authConfig";
  import {callGetPublicTeams,canUserRestoreTeams}  from "../component/graph";
//  import InfiniteScroll from "react-infinite-scroll-component";
//import ReactPaginate from 'react-paginate';
import ReactTooltip from 'react-tooltip';
//import DialogExample from '../component/DialogBox/OpenDialogBox';
// import { getTsBuildInfoEmitOutputFilePath } from 'typescript';


const classNames = mergeStyleSets({
    fileIconHeaderIcon: {
      padding: 0,
      fontSize: '16px',
    },
    fileIconCell: {
      textAlign: 'center',
      selectors: {
        '&:before': {
          content: '.',
          display: 'inline-block',
          verticalAlign: 'middle',
          height: '100%',
          width: '0px',
          visibility: 'hidden',
        },
      },
    },
    fileIconImg: {
      verticalAlign: 'middle',
      maxHeight: '16px',
      maxWidth: '16px',
    },
    controlWrapper: {
      display: 'flex',
      flexWrap: 'wrap',
    },
    workspaceImage: {
      width: '36px',
      height: '36px',
    },
    exampleToggle: {
      display: 'inline-block',
      marginBottom: '10px',
      marginRight: '30px',
    },
    selectionDetails: {
      marginBottom: '20px',
    },
  });

  // const controlStyles = {
  //   root: {
  //     margin: '0 30px 20px 0',
  //     maxWidth: '300px',
  //     marginLeft : 10,
  //   },
  // };
  
  const icons = Object.keys(ReactIcons).reduce((acc: React.FC[], exportName) => {
    if ((ReactIcons as any)[exportName]?.displayName) {
      if(exportName === "MoreVerticalIcon"){
        acc.push((ReactIcons as any)[exportName] as React.FunctionComponent);
      }
    }
  
    return acc;
  }, []);

  export interface IWorkspaceExampleState {
    columns: IColumnConfig[];
    displayItems: IWorkspace[];
    serachItem: IWorkspace[];
    itemsList : IWorkspace[];
    sortItemsDetails : IWorkspace[];
    uniqueFilterValues : string[];
    //selectionDetails: string;
    sortItemCheck : boolean;
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
    userIsAdmin : boolean;
    hasMore : boolean;
    isPanelOpen : boolean;
    isPanelClose : boolean;
    itemArrayAppend : number;
    checkSearchItem : boolean;
    contextualMenuProps? : IContextualMenuProps;
    today: Date;
    inActiveCount : number;
    itemWithNoOwner : number;
    teamsMissingInfo : number;
    teamsExternalUser : number;
    Paginationdata: any;
    perPage: number;
    pages: number;
    currentItem : any;
  }

  export interface IWorkspace {
    key: string;
    test: string;
    name: string;
    businessDepartment: string;
    status: string;
    type: string;
    classification : string;
    businessOwner : string;
    teamsWithNoOwner : number;
    teamsExternalUser : number;
    teamsSiteUrl : string;
    sharePointSiteUrl : string;
    }
    

  interface IWorkspaceProps {
    instance : any;
    accounts : any;
    userIsAdmin : any;
  }
 // let userRole :any ;
  class WorkspaceDetails extends React.Component<IWorkspaceProps, IWorkspaceExampleState> {
    
    constructor(props: IWorkspaceProps, state: IWorkspaceExampleState){
        super(props);
        this.fetchMoreData = this.fetchMoreData.bind(this);
        // onscroll = (event) => {
        //   console.log(event);
        // }
        

        const columns: IColumnConfig[] = [
          {
            key: 'column1',
            name: 'test',
            text: '',
            isResizable: false,
            className: classNames.fileIconCell,
            iconClassName: classNames.fileIconHeaderIcon,
            ariaLabel: 'Column operations for File type, Press to sort on File type',
            iconName: 'Page',
            isIconOnly: true,
            fieldName: 'name',
            minWidth: 1,
            maxWidth: 1,
            // onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev), content={`${item.test} file`}
            onRender: (item: IWorkspace) => (
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg1">
                <TooltipHost key={item.key} >
                  <img src={item.test} className={classNames.fileIconImg} alt={`${item.test} file icon`} />
                </TooltipHost>
              </div>
            ),
          },
          //   {
          //     key: 'name',
          //     name: 'Name',
          //     text: 'Name',
          //     editable: true,
          //     dataType: 'string',
          //     minWidth: 150,
          //     maxWidth: 150,
          //     isResizable: true,
          //     includeColumnInExport: true,
          //     includeColumnInSearch: true,
          //     applyColumnFilter: true,
          //     onRender: (item: IWorkspace) => {
          //         return <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">  <span onClick= {() => window.open(item.teamsSiteUrl, "_blank")} > {item.name} </span> </div> ;
          //     },
          // },
          {
            key: 'name',
            name: 'Name',
            text: 'Name',
            fieldName: 'name',
            minWidth: 120,
            maxWidth: 120,
            isResizable: false,
            dataType: 'string',
            includeColumnInExport: true,
            includeColumnInSearch: true,
            onColumnClick: (ev, columns) => this._onColumnClick(columns, this.state.sortItemCheck),
            isSorted: true,
            isSortedDescending: false,
            sortAscendingAriaLabel: 'Sorted A to Z',
            sortDescendingAriaLabel: 'Sorted Z to A',
            data: 'string',
            onRender: (item: IWorkspace) => {
              return <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">  <span onClick={() => window.open(item.teamsSiteUrl, "_blank")} > {item.name} </span> </div>;
            },
            isPadded: true,
          },
          {
            key: "column3",
            text: '',
            name: "",
            fieldName: "Options",
            minWidth: 35,
            isResizable: false,
            maxWidth: 35,
            onRender: (item: IWorkspace) => {
              const plainCardProps: IPlainCardProps = {
                onRenderPlainCard: this.onRenderPlainCard,
                renderData: item,
              };
              return (
                // <div className={classNames.controlWrapper}> 
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <HoverCard
                    plainCardProps={plainCardProps}
                    instantOpenOnClick={true}
                    type={HoverCardType.plain}
                  >
                    {icons
                      .map((Icon: React.FunctionComponent<ReactIcons.ISvgIconProps>) => (
                        <Icon key={item.key} aria-label={'MoreVertical'?.replace('', '')} />
                      ))
                    }
                    {/* <IconButton
                         // className = { classNames.workspaceImage } //{styles.workspaceImage}
                          iconProps={{ iconName: "MoreVerticalIcon" }}
                          aria-label = { iconName 'MoreVerticalIcon'}
                        /> */}
                  </HoverCard>
                </div>
              );
            },
          },
          {
            key: 'businessDepartment',
            name: 'businessDepartment',
            text: 'Business Department',
            editable: true,
            dataType: 'string',
            minWidth: 160,
            maxWidth: 180,
            isResizable: true,
            includeColumnInExport: true,
            includeColumnInSearch: true,
            applyColumnFilter: true,
            onRender: (item: IWorkspace) => {
              return <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">  <span key={item.key}>{item.businessDepartment}</span> </div>;
            },
          },
          // {
          //   key: 'column4',
          //   name: 'Business Department',
          //   text:'Business Department',
          //   fieldName: 'businessDepartment',
          //   minWidth: 70,
          //   maxWidth: 90,
          //   isResizable: true,
          //   onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
          //   data: 'number',
          //   onRender: (item: IWorkspace) => {
          //     return <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg2">  <span key={item.key}>{item.businessDepartment}</span> </div>;
          //   },
          //   isPadded: true,
          // },
          {
            key: 'businessOwner',
            name: 'businessOwner',
            text: 'Business Owner',
            editable: true,
            dataType: 'string',
            minWidth: 140,
            maxWidth: 140,
            isResizable: true,
            includeColumnInExport: true,
            includeColumnInSearch: true,
            applyColumnFilter: true
          },
          // {
          //   key: 'businessOwner',
          //   name: 'businessOwner',
          //   text: 'Business Owner',
          //   editable: true,
          //   dataType: 'string',
          //   minWidth: 100,
          //   maxWidth: 100,
          //   isResizable: true,
          //   includeColumnInExport: true,
          //   includeColumnInSearch: true,
          //   //inputType: EditControlType.MultilineTextField,
          //   applyColumnFilter: true
          // },
          {
            key: 'status',
            name: 'status',
            text: 'Status',
            editable: true,
            dataType: 'string',
            minWidth: 110,
            maxWidth: 110,
            isResizable: true,
            includeColumnInExport: true,
            includeColumnInSearch: true,
            applyColumnFilter: true
          },
          {
            key: 'type',
            name: 'type',
            text: 'Type',
            editable: true,
            dataType: 'string',
            minWidth: 110,
            maxWidth: 110,
            isResizable: true,
            includeColumnInExport: true,
            includeColumnInSearch: true,
            onRender: (item: IWorkspace) => {
              return <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">  <span key={item.key}>{item.type}</span> </div>;
            },
          },
          // {
          //   key: 'column7',
          //   name: 'Type',
          //   text:'Type',
          //   fieldName: 'type',
          //   minWidth: 70,
          //   maxWidth: 90,
          //   isResizable: true,
          //   isCollapsible: true,
          //   data: 'number',
          //   onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
          //   onRender: (item: IWorkspace) => {
          //     return <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg1">  <span key={item.key}>{item.type}</span> </div>;
          //   },
          // },
          {
            key: 'classification',
            name: 'classification',
            text: 'Classification',
            editable: true,
            dataType: 'string',
            minWidth: 150,
            maxWidth: 150,
            isResizable: true,
            includeColumnInExport: true,
            includeColumnInSearch: true,
            onRender: (item: IWorkspace) => {
              return <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg1">  <span key={item.key}>{item.classification}</span> </div>;
            },
          },
          // {
          //   key: 'column8',
          //   name: 'Classification',
          //   text:'Classification',
          //   fieldName: 'classification',
          //   minWidth: 70,
          //   maxWidth: 90,
          //   isResizable: true,
          //   isCollapsible: true,
          //   data: 'number',
          //   onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
          //   onRender: (item: IWorkspace) => {
          //     return <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg1">  <span key={item.key}>{item.classification}</span> </div>;
          //   },
          // },
          {
            key: 'column9',
            name: 'test',
            text: '',
            className: classNames.fileIconCell,
            iconClassName: classNames.fileIconHeaderIcon,
            ariaLabel: 'Column operations for File type, Press to sort on File type',
            iconName: 'Page',
            isIconOnly: true,
            fieldName: 'name',
            minWidth: 16,
            maxWidth: 16,
            // onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
            onRender: (item: IWorkspace) => (
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg1">
                <TooltipHost key={item.key} >
                  <img onClick={() => window.open(item.sharePointSiteUrl, "_blank")} src={sharepointImg} className={classNames.fileIconImg} alt={`${item.test} file icon`} />
                </TooltipHost>
              </div>
            ),
          },
        ];

        let today = new Date();
        this.state = {
          displayItems: [],
          serachItem : [],
          itemsList : [],
          sortItemsDetails : [],
          columns: columns,
          contextualMenuProps:undefined,
          sortItemCheck : true,
          uniqueFilterValues : [],
          // selectionDetails: this._getSelectionDetails(),
          isModalSelection: false,
          isCompactMode: false,
          announcedMessage: undefined,
          userIsAdmin : this.props.userIsAdmin,
          hasMore : true,
          today :today ,
          isPanelOpen : false,
          isPanelClose : true,
          checkSearchItem : false,
          itemArrayAppend : 20,
          inActiveCount : 0,
          itemWithNoOwner :0,
          teamsMissingInfo : 0,
          teamsExternalUser : 0,
          Paginationdata : [],
          perPage : 8,
          pages : 0,
          currentItem : {},
        };
    }

    private _onColumnContextMenu = (column: IColumn, ev: React.MouseEvent<HTMLElement>): void => {
      this.setState({
        contextualMenuProps: this._getContextualMenuProps(ev, column),
      });
    };

    

    private _getContextualMenuProps(ev: React.MouseEvent<HTMLElement>, column: IColumn): IContextualMenuProps  {
      

      // var uniqueVals = [], enabledVals = [];
      //var workspacesUnfiltered :any , workspaces;
      
      // workspaces = this.state.itemsList;
      // workspacesUnfiltered = this.state.columns;
      
      // let namesArray = workspaces.map(elem => elem.businessDepartment);
      // let namesTraversed : any = [];
      // let currentCountOfName = 1;
      // let len = 0;
      let itemForCheckbox = this.state.itemsList;
      
      let uniqueValues  = itemForCheckbox.filter( (ele, ind) => ind === itemForCheckbox.findIndex( elem => elem.businessDepartment.trim() !== "" ?  elem.businessDepartment.trim() === ele.businessDepartment.trim() : undefined ) );
      
      let uniqueString : string [] = [];
      
      uniqueValues.forEach((element) => uniqueString.push(element.businessDepartment));
      
      this.setState({
        uniqueFilterValues : uniqueString
      });

      // let ItemsForCheckBox ;

      // const items = [
      //   { key: uniqueString[0], text: uniqueString[0], canCheck: true  },
      //   { key: uniqueString[1], text: uniqueString[1], canCheck: true },
      //   { key: uniqueString[1], text: uniqueString[0], canCheck: true },
      // ];
      
      const items = [
        {
          key: uniqueString[0],
          name: uniqueString[0],
          iconProps: { iconName: 'SortUp' },
          canCheck: true,
          checked: column.isSorted && !column.isSortedDescending,
          isChecked :  uniqueString[0],
        },
        {
          key: uniqueString[1],
          name: uniqueString[1],
          iconProps: { iconName: 'SortDown' },
          canCheck: true,
          checked: column.isSorted && column.isSortedDescending,
          isChecked : uniqueString[1] ,
        },
      ];

      return {
        items : items,
        target: ev.currentTarget as HTMLElement,
        directionalHint: DirectionalHint.bottomLeftEdge,
        gapSpace: 10,
        isBeakVisible: true,
        onDismiss: this.onContextualMenuDismissed,
      };
    }

  private onContextualMenuDismissed = (): void => {
    this.setState({
      contextualMenuProps: undefined,
    });
  }

    private _onColumnClick = (column: IColumn , checkOrder : boolean): void => {
      const { columns, sortItemsDetails } = this.state;
      const newColumns: IColumnConfig[] = columns.slice();
      const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
      newColumns.forEach((newCol: IColumn) => {
        if (newCol === currColumn) {
          currColumn.isSortedDescending = !currColumn.isSortedDescending;
          currColumn.isSorted = true;
          this.setState({
            announcedMessage: `${currColumn.name} is sorted ${
              currColumn.isSortedDescending ? 'descending' : 'ascending'
            }`,
          });
        } else {
          newCol.isSorted = false;
          newCol.isSortedDescending = true;
        }
      });
      const newItems = _copyAndSort(sortItemsDetails, currColumn.fieldName!, checkOrder);
      let itemsCount = 20;
      // this.setState({
      //   itemsList:newItems,
      //   itemArrayAppend : itemsCount
      // });
      let getItemsbyScroll = newItems.slice(0, itemsCount);
      this.setState({
        itemsList:newItems,
        itemArrayAppend : itemsCount,
        columns: newColumns,
        displayItems: getItemsbyScroll,
        sortItemCheck: !checkOrder,
      });
    };
    

    public onRenderPlainCard(){
      return(
        <div className='elliptical-menu'>
          <Dialog
          cancelButton="Program protocol"
          confirmButton="Hack matrix"
          content="Quantify array"
          header="Back up application"
          headerAction="Reboot port"
          trigger={<Button text="Edit" />}
        />
        <br />
        <Button
          text="Archived"
          //className= {styles.createNewButton}
          //onClick={() => this.setState({ currentItem: item, dialog: "Delete" })}
        />
        <br />
        <Button
          text="Delete"
          //className= {styles.createNewButton}
          //onClick={() => this.setState({ currentItem: item, dialog: "Delete" })}
        />
        <br />
        </div>
      );
    }

    public OpenDialogBox(){
      return (
        <Dialog
          cancelButton="Program protocol"
          confirmButton="Hack matrix"
          content="Quantify array"
          header="Back up application"
          headerAction="Reboot port"
          trigger={<Button content="A trigger" />}
        />
      );
    }

    public async componentDidMount(){

      await this._getUserRole().then((teamsUserRoleStatus:boolean)  => {
        if(teamsUserRoleStatus === true){
          //userRole = teamsUserRoleStatus;
          this.setState({
            userIsAdmin : teamsUserRoleStatus // true
          })
          console.log("Teams User Role status : " + this.state.userIsAdmin );
        }
        else {
          //userRole = teamsUserRoleStatus;
          this.setState({
            userIsAdmin : teamsUserRoleStatus
          });
        }
      });

      await this._getInActiveTeams().then((ActiveTeams : any[]) => {
        console.log("Component Teams Log =-=-=-=-= " + ActiveTeams );
      })

      await this._getAllPublicTeams().then((teamsDetails : any[]) => {
        console.log("Component Teams Log" + teamsDetails );
        //if(teamsDetails.status === ''){}
        // this._allItems = teamsDetails;

        let countNumber = 0;
        let countMissiongInformation = 0; 
        let countExternalUser = 0;
        for(let i=0; i< teamsDetails.length ; i++) {
          if(teamsDetails[i].teamsWithNoOwner === 0){
            countNumber = countNumber + 1;
          }
          if(teamsDetails[i].teamsExternalUser > 0) {
            countExternalUser = countExternalUser + 1;
          }
          if(teamsDetails[i].businessOwner ===''  || teamsDetails[i].businessDepartment === ''  || teamsDetails[i].classification === '' || teamsDetails[i].type === '' ) {
            countMissiongInformation = countMissiongInformation + 1;
          }
        }

        this.setState({
          displayItems: teamsDetails.slice(0,this.state.itemArrayAppend),
          serachItem : teamsDetails,
          itemsList : teamsDetails,
          sortItemsDetails : teamsDetails,
          itemWithNoOwner : countNumber,
          teamsMissingInfo : countMissiongInformation,
          teamsExternalUser : countExternalUser,
        });
        var exp: any = document.getElementById('export');
      document.getElementsByClassName('ms-TextField-wrapper')[0].appendChild(exp)
      });
      document.getElementsByClassName('ms-TextField-field')[1].setAttribute('placeholder', 'Search');



      this.setState({ pages: Math.round(this.state.itemsList.length / this.state.perPage) })
        let page = 0;
        let itemsPagination = this.state.itemsList.slice(page * this.state.perPage, (page + 1) * this.state.perPage);

        this.setState({ Paginationdata: itemsPagination});
    }

    public  handlePageClick = (event:any) => {
      let page = event.selected;

      //Pagination

      let items = this.state.itemsList.slice(page * this.state.perPage, (page + 1) * this.state.perPage);

      this.setState({ Paginationdata: items });

  }

    public updateMoreData = () => {
      this.setState({
        displayItems : this.state.displayItems
      });
    };

    public fetchMoreData = () => {

      let tempAllItems = this.state.itemsList;

      this.setState({
        itemArrayAppend : this.state.itemArrayAppend + 20
      });

      // if(this.state.displayItems.length == this.state.itemsList.length){
      //   this.setState({ hasMore: false });
      //   return;
      // }

      
      // a fake async api call like which sends
      // 20 more records in .5 secs

      if(this.state.itemsList.length > 0){
        if (this.state.displayItems.length === this.state.itemsList.length) {
          this.setState({ hasMore: false });
          return;
        }
        setTimeout(() => {
          this.setState({
            displayItems: this.state.itemsList.slice(0, this.state.itemArrayAppend)
          });
        }, 1500);
      }
      else {
        if (this.state.displayItems.length === tempAllItems.length) {
          this.setState({ hasMore: false });
          return;
        }
        setTimeout(() => {
          this.setState({
            displayItems: tempAllItems.slice(0, this.state.itemArrayAppend)
          });
        }, 1500);
      }
    };
    
    render(){

      return(
        <div className='container-custom'>
        {
          this.state.userIsAdmin ? <div className="ms-Grid" dir="ltr">
            {/* style= {{ height : '40px' }} */}
            <div className="ms-Grid-row" style={{ height: '40px' }}>
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">
                {/* style = {{ textAlign: 'left', marginLeft:'10px'}} */}
                <h3 style={{ textAlign: 'left', marginLeft: '10px' }}> Manage Teams </h3>
              </div>
            </div>
            <div className="ms-Grid-row">
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                <div className='white-wrapper' style={{
                  backgroundColor: '#FFFFFF',
                  //margin: '0px 10px 0px 0px ',
                  padding: 2,
                }}>
                  <div className="title_wrapper">
                    <h6 style={{ textAlign: 'left', margin: '10px 15px 0px 15px', fontFamily: 'Segoe UI' }} > Inactive Teams
                    </h6>
                    <div style={{ marginLeft: '0px' }} data-tip="Total Teams With InActive Status">
                      <img
                      height="10.0"
                      width="10.0"
                        src={InfoIcon}
                        alt="new"
                      />
                      <ReactTooltip></ReactTooltip>
                    </div>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <h3 style={{ textAlign: 'left', margin: '10px 15px 0px 13px', fontSize: 36 }}> {this.state.inActiveCount} </h3>
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <img
                        src={InactiveIconTeams}
                        alt="new"
                      />
                    </div>
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                <div className='white-wrapper' style={{
                  backgroundColor: '#FFFFFF',
                  margin: '0px 10px 0px 0px ',
                  padding: 2,
                }}>
                  <div className="title_wrapper">
                    <h6 style={{ textAlign: 'left', margin: '10px 15px 0px 15px', fontFamily: 'Segoe UI' }} > Teams With No Owner</h6>
                    <div data-tip="Total Teams With No Owner">
                      <img
                        height="10.0"
                        width="10.0"
                        src={InfoIcon}
                        alt="new"
                      />
                    </div>
                    <ReactTooltip></ReactTooltip>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <h3 style={{ textAlign: 'left', margin: '10px 15px 0px 13px', fontSize: 36 }}> {this.state.itemWithNoOwner} </h3>
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <img
                        
                        src={NoOwnersIcon}
                        alt="new"
                      />
                    </div>
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                <div className='white-wrapper' style={{
                  backgroundColor: '#FFFFFF',
                  margin: '0px 10px 0px 0px ',
                  padding: 2,
                }}>
                  <div className="title_wrapper">
                    <h6 style={{ textAlign: 'left', margin: '10px 15px 0px 15px', fontFamily: 'Segoe UI' }}> Teams With External User</h6>
                    <div data-tip="Total Teams With External User">
                      <img
                        height="10.0"
                        width="10.0"
                        src={InfoIcon}
                        alt="new"
                      />
                    </div>
                    <ReactTooltip></ReactTooltip>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <h3 style={{ textAlign: 'left', margin: '10px 15px 0px 13px', fontSize: 36 }}> {this.state.teamsExternalUser} </h3>
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <img
                       
                        src={ExtUsersIcon}
                        alt="new"
                      />
                    </div>
                  </div>
                </div>
              </div>
              <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg3">
                <div className='white-wrapper' style={{
                  backgroundColor: '#FFFFFF',
                  margin: '0px 0px 0px 0px ',
                  padding: 2,
                }}>
                  <div className="title_wrapper">
                    <h6 style={{ textAlign: 'left', margin: '10px 15px 0px 15px', fontFamily: 'Segoe UI' }}> Teams Missing Information</h6>
                    <div data-tip="Total Teams With Missing Information">
                      <img
                        height="10.0"
                        width="10.0"
                        src={InfoIcon}
                        alt="new"
                      />
                    </div>
                    <ReactTooltip></ReactTooltip>
                  </div>
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <h3 style={{ textAlign: 'left', margin: '10px 15px 0px 13px', fontSize: 36 }}> {this.state.teamsMissingInfo} </h3>
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg6">
                      <img
                        src={TeamsMissingIcon}
                        alt="new"
                      />
                    </div>
                  </div>
                </div>
              </div>
            </div>
            {/* Render table */}
            <div className="ms-Grid" style={{ margin: '15px 0', backgroundColor: '#FFFFFF', boxShadow: '1px 2px 7px #0000000f', borderRadius: '5px' }}>
              {/* region Showing the All Teams Section */}
              <div className="ms-Grid-row" style={{ marginTop: 20 }}>
                {/* <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">
                <h5 className='section-heading' style={{ textAlign: 'left', marginLeft: 15 }}> All Teams </h5>
                </div> */}
              </div>
              {/* showing the search teams section */}
              <div>
                <div className="ms-Grid-row" >
                  <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">
                  
                    {/* <TextField placeholder="Search For a Team" onScroll={this.fetchMoreData} onChange={(event: any) => this._onChangeText(event)}
                      styles={controlStyles} /> */}
                  </div>
                </div>
              </div>
              {/* {this.state.contextualMenuProps && <ContextualMenu {... } />} */}
              {this.state.contextualMenuProps && <ContextualMenu {...this.state.contextualMenuProps} />}
              {/* {this.state.uniqueFilterValues.map((value) => {
return <Checkbox
label={value ? value : "(No value)"}
//className={styles.checkbox}
// disabled={this.state.enabledValues.indexOf(value) == -1}
// defaultChecked={this.state.checkedFilterValues.indexOf(value) !== -1}
// onChange={(ev: any, checked: boolean) => {
// if (checked) {
// this.state.checkedFilterValues.push(value);
// } else {
// let index = this.state.checkedFilterValues.indexOf(value);
// if (index !== -1) {
// this.state.checkedFilterValues.splice(index, 1);
// }
// }
// this.setState({ checkedFilterValues: this.state.checkedFilterValues });
// }}
/>;
})} */}
              {/* This Renders the Teams Records */}
              <div className="ms-Grid-row" >
                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12" style ={{ padding: '0px' }}>
                  {/* <InfiniteScroll
dataLength={this.state.displayItems.length}
next={this.fetchMoreData}
hasMore={this.state.hasMore}
loader={<h4>Loading...</h4>}
// onScroll = { this.updateMoreData }
endMessage={
<p style={{ textAlign: "center" }}>
<b>Yay! You have seen it all</b>
</p>
}
> */}
                  <div className={classNames.controlWrapper}>
                      <TextField placeholder='Search For a Team' className={mergeStyles({ width: '60vh', paddingBottom:'10px' })} onChange={(event) => EventEmitter.dispatch(EventType.onSearch, event)}/>
                  </div>
                  <EditableGrid
                    id={1}
                    columns={this.state.columns}
                    items={this.state.itemsList}
                    //enableCellEdit={true}
                    enableExport={true}
                    // enableTextFieldEditMode={true}
                    // enableTextFieldEditModeCancel={true}
                    // enableGridRowsDelete={true}
                    // enableGridRowsAdd={true}
                    //height={'40vh'}
                    width={'140vh'}
                    //position={'relative'}
                    // enableUnsavedEditIndicator={true}
                    //onGridSave={onGridSave}
                    //enableGridReset={true}
                    enableColumnFilters={true}
                    //enableColumnFilterRules={true}
                    // enableRowAddWithValues={{enable : true, enableRowsCounterInPanel : true}}
                    //layoutMode={DetailsListLayoutMode.justified}
                    selectionMode={SelectionMode.none}
                  // enableRowEdit={true}
                  // enableRowEditCancel={true}
                  // enableBulkEdit={true}
                  // enableColumnEdit={true}
                  // enableSave={true}
                  />
                   {/* <ReactPaginate
                      previousLabel={'<'}
                      nextLabel={'>'}
                      pageCount={this.state.pages}
                      onPageChange={this.handlePageClick}
                      containerClassName="pagination"
                      activeClassName="active"
                  /> */}

                  {/* <DetailsList
items= {[ ...this.state.displayItems]}
compact={this.state.isCompactMode}
columns={this.state.columns}
selectionMode={SelectionMode.none}
getKey={this._getKey}
setKey="set"
// layoutMode={DetailsListLayoutMode.justified}
// isHeaderVisible={true}
// data-is-scrollable="true"
// onItemInvoked={this._onItemInvoked}
/> */}
                  {/* <DetailsList
items= {[ ...this.state.displayItems]}
compact={this.state.isCompactMode}
columns={this.state.columns}
selectionMode={SelectionMode.none}
onRenderItemColumn = { (item) => {
return item
} }
// getKey={this._getKey}
// setKey="set"
// layoutMode={DetailsListLayoutMode.justified}
// isHeaderVisible={true}
// data-is-scrollable="true"
// onItemInvoked={this._onItemInvoked}
/> */}
                  {/* </InfiniteScroll> */}
                </div>
              </div>
              {/* {console.log("Item Count :- " + this.state.items.length)}
<DetailsList
items={this.state.items}
columns={this.state.columns}
getKey={this._getKey}
/> */}
              <Panel
                headerText="Sample panel"
                isOpen={this.state.isPanelOpen}
                onDismiss={() => { this.setState({ isPanelOpen: false }); }}
                // You MUST provide this prop! Otherwise screen readers will just say "button" with no label.
                closeButtonAriaLabel="Close"
              >
                <p>Content goes here.</p>
              </Panel>
            </div>
          </div>
            :
            <div className="ms-Grid" dir="ltr" style={{display: 'flex', justifyContent: 'center',

            flexDirection: 'column',

            alignItems: 'center',

            height: '100vh'}}>
              <div className="ms-Grid-row" style={{ marginTop: 300 }}>
                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">
                  <img
                    height="140"
                    width="140"
                    src={LockIcon}
                    alt="new"
                  /></div>
              </div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">
                  <h5> Sorry but you don't have access to this feature </h5>
                </div>
              </div>
              <div className="ms-Grid-row">
                <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg12">
                  <p style={{ fontFamily: "Segoe UI" }}> Tellus Admin is only available to administrators </p>
                </div>
              </div>
            </div>
        }
      </div>
      )
    }
    
    public _closePanel(){
      this.setState({
        isPanelOpen : false,
        isPanelClose : false,
      });
    }

    private _getKey(item: any, index?: number): string {
      return item.key ;
    }

    public _onChangeText = (ev: any): void => {

      let testData = this.state.serachItem ;
      let searchData = ev.target.value !== ""  ?  testData.filter(i => i.name.toLowerCase().startsWith(ev.target.value.toLowerCase())) : testData;
      

      if(searchData.length < 20)
      {
        this.setState({
          displayItems: searchData.slice(0, this.state.itemArrayAppend),
          itemsList : searchData,
          checkSearchItem : true,
          hasMore : false,
         },
          () => console.log(this.state.displayItems));
      }
      else {
        this.setState({
          displayItems: searchData.slice(0, this.state.itemArrayAppend),
          itemsList : searchData,
          checkSearchItem : true,
          hasMore : true,
         },
          () => console.log(this.state.displayItems));
      }
     };

    public _getAllPublicTeams = async () : Promise<IWorkspace[]>  =>  {
      
      return new Promise<any>((resolve, reject) => 
      {
        const items: IWorkspace[] = [];
        var today = this.state.today.getTime();
        let totalInActiveTeams = 0;
        // let countItem = 0;

        this.props.instance.acquireTokenSilent({
          ...loginRequest,
          account: this.props.accounts[0]
            }).then((response:any) => {
          callGetPublicTeams(response.accessToken).then(response => response).then((data:any[]) =>
          {
            data.forEach(element => {
              var daysSinceActivity = 0;

                if (element.latestActivityDate != null) {
                  daysSinceActivity =
                    (today - new Date(element.latestActivityDate).getTime()) /
                    (1000 * 60 * 60 * 24.0);
                }
                if (element.latestActivityDate != null) {
                  daysSinceActivity =
                    (today - new Date(element.latestActivityDate).getTime()) /
                    (1000 * 60 * 60 * 24.0);
                }
                if (element.status === "Active" && daysSinceActivity >= 97) {
                  element.status = "Inactive";
                  totalInActiveTeams = totalInActiveTeams + 1; 
                  items.push({
                    test: element.imageBlob,
                    key: element.id.toString(),
                    teamsSiteUrl : element.urlTeams,
                    sharePointSiteUrl : element.urlSharePoint,
                    name: element.title,
                    businessDepartment: element.businessDepartment,
                    status: element.status,
                    type: element.template,
                    classification: element.classification,
                    businessOwner : element.ownerName,
                    teamsExternalUser : element.teamsExternalUser,
                    teamsWithNoOwner : element.teamsOwner
                  });
                }
                else {
                  items.push({
                    test: element.imageBlob,
                    key: element.id.toString(),
                    name: element.title,
                    teamsSiteUrl : element.urlTeams,
                    sharePointSiteUrl : element.urlSharePoint,
                    businessDepartment: element.businessDepartment,
                    status: element.status,
                    type: element.template,
                    classification: element.classification,
                    businessOwner : element.ownerName,
                    teamsExternalUser : element.teamsExternalUser,
                    teamsWithNoOwner : element.teamsOwner
                  });
                }
            });
            // let countWithnoOwner = items.map(x => x.businessOwner == null || "" ? true : false).length;
            
            this.setState({
              inActiveCount : totalInActiveTeams,
              // itemWithNoOwner : countWithnoOwner,
            });
            
            resolve(items);
          });
        });
      }
      );
    }

    public _getInActiveTeams = async () : Promise<IWorkspace[]>  =>  {

        return new Promise<any>((resolve, reject) => 
        {
          const items: IWorkspace[] = [];
          this.props.instance.acquireTokenSilent({
            ...loginRequest,
            account: this.props.accounts[0]
              }).then((response:any) => {
            callGetPublicTeams(response.accessToken).then(response => response).then((data:any[]) =>
            {

              var today = this.state.today.getTime();
              data.forEach(element => {
                var daysSinceActivity = 0;
                
                if (element.latestActivityDate != null) {
                  daysSinceActivity =
                    (today - new Date(element.latestActivityDate).getTime()) /
                    (1000 * 60 * 60 * 24.0);
                }
                if (element.status === "Active" && daysSinceActivity >= 97) {
                  element.status = "Inactive";
                  items.push({
                    test: element.imageBlob,
                    key: element.id.toString(),
                    teamsSiteUrl : element.UrlTeams,
                    sharePointSiteUrl : element.UrlSharePoint,
                    name: element.title,
                    businessDepartment: element.businessDepartment,
                    status: element.status,
                    type: element.template,
                    classification: element.classification,
                    businessOwner : element.ownerName,
                    teamsExternalUser : element.teamsExternalUser,
                    teamsWithNoOwner : element.teamsOwner
                  });
                }
                else {
                  items.push({
                    test: element.imageBlob,
                    key: element.id.toString(),
                    teamsSiteUrl : element.UrlTeams,
                    sharePointSiteUrl : element.UrlSharePoint,
                    name: element.title,
                    businessDepartment: element.businessDepartment,
                    status: element.status,
                    type: element.template,
                    classification: element.classification,
                    businessOwner : element.ownerName,
                    teamsExternalUser : element.teamsExternalUser,
                    teamsWithNoOwner : element.teamsOwner
                  });
                }
              })
              resolve(items);
            });
          });
        }
      );
    }

    public _getUserRole = async () : Promise<boolean> => {
      return new Promise<boolean>((resolve, reject) =>{
        this.props.instance.acquireTokenSilent({
          ...loginRequest,
        account: this.props.accounts[0]
        }).then((response:any) => {
          canUserRestoreTeams(response.accessToken,this.props.accounts[0].username).then(response => response).then((data:any) =>{
              resolve(data);
          })
        })
      })
    }
  }

  function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
    
    let key = columnKey as keyof T;
    let sortedItems = items.slice(0).sort((a :any , b:any) => ( //a[key] === null ? 1 : b[key] === null ? -1 :
      (a[key].toString().toLowerCase() === b[key].toString().toLowerCase() ? 0 : isSortedDescending ? a[key].toString().toLowerCase() < b[key].toString().toLowerCase() : a[key].toString().toLowerCase() > b[key].toString().toLowerCase()) ? 1 : -1)
    );
    return sortedItems;
  }

  // const DialogExample = () => {
  //   return (
  //     <>
  //         <Dialog
  //           cancelButton="Cancel"
  //           confirmButton="Confirm"
  //           header="Action confirmation"
  //           trigger={<Button content="Open a dialog" />}
  //         />
  //     </>
  //   )
  // }



  // function getContextualMenuDetails(){
  //   const [Selection, SetSelection] = React.useState<{ [key: string]: boolean }>({});
  //   // const menuProps: IContextualMenuProps = React.useMemo(
  //   //   () => ({
  //   //     shouldFocusOnMount: true,
  //   //     items: [
  //   //       { key: "HR", text: 'New', canCheck: true, isChecked: selection["HR"] },
  //   //       { key: "Developer", text: 'Share', canCheck: true, isChecked: selection["Developer"]},
  //   //       { key: "Infra", text: 'Mobile', canCheck: true, isChecked: selection["Infra"]},
  //   //     ],
  //   //   }),
  //   //   [selection],
  //   // );
  // }

  export default WorkspaceDetails;
  