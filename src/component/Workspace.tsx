import * as React  from 'react';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import * as ReactIcons from '@fluentui/react-icons-mdl2';
import { DetailsList, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { TooltipHost } from '@fluentui/react';
import { Panel } from '@fluentui/react/lib/Panel';
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
  } from 'office-ui-fabric-react/lib/ContextualMenu';

  import { TextField } from '@fluentui/react/lib/TextField';
  import { Label } from '@fluentui/react/lib/Label';
  import { loginRequest } from "../component/authConfig";
  import {callAllTeamsRequest,canUserRestoreTeams}  from "../component/graph";
  import InfiniteScroll from "react-infinite-scroll-component";


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

  const controlStyles = {
    root: {
      margin: '0 30px 20px 0',
      maxWidth: '300px',
    },
  };
  
  const icons = Object.keys(ReactIcons).reduce((acc: React.FC[], exportName) => {
    if ((ReactIcons as any)[exportName]?.displayName) {
      if(exportName === "MoreVerticalIcon"){
        acc.push((ReactIcons as any)[exportName] as React.FunctionComponent);
      }
    }
  
    return acc;
  }, []);

  export interface IWorkspaceExampleState {
    columns: IColumn[];
    displayItems: IWorkspace[];
    serachItem: IWorkspace[];
    itemsList : IWorkspace[];
    sortItemsDetails : IWorkspace[];
    //selectionDetails: string;
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
  }

  interface IWorkspaceProps {
    instance : any;
    accounts : any;
  }

  class WorkspaceDetails extends React.Component<IWorkspaceProps, IWorkspaceExampleState> {
    
    constructor(props: IWorkspaceProps, state: IWorkspaceExampleState){
        super(props);
        this.fetchMoreData = this.fetchMoreData.bind(this);
        // onscroll = (event) => {
        //   console.log(event);
        // }

        const columns : IColumn[] = [
            {
                key: 'column1',
                name: 'test',
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
                  <TooltipHost key={item.key} content={`${item.test} file`}>
                    <img src={item.test} className={classNames.fileIconImg} alt={`${item.test} file icon`} /> 
                  </TooltipHost>
                ),
              },
              {
                key: 'column2',
                name: 'Name',
                fieldName: 'name',
                minWidth: 210,
                maxWidth: 350,
                isResizable: true,
                onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),

                // onColumnClick: (ev, column) => {
                //   this.onColumnClick(column, ev);
                // },
                // onColumnContextMenu: (column, ev) => {
                //   this.onColumnClick(column, ev);
                // },
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                data: 'number',
                onRender: (item: IWorkspace) => {
                  return <span > {item.name}</span>;
                },
                isPadded: true,
              },
              {
                key: "column3",
                name: "",
                fieldName: "Options",
                minWidth: 10,
                maxWidth: 10,
                onRender:(item: IWorkspace) => {
                  const plainCardProps: IPlainCardProps = {
                    onRenderPlainCard: this.onRenderPlainCard,
                    renderData: item,
                  };
                  return (
                    // <div className={classNames.controlWrapper}> 
                    <HoverCard
                      plainCardProps={plainCardProps}
                      instantOpenOnClick={true}
                      type={HoverCardType.plain}
                    >
                {icons
                  .map((Icon: React.FunctionComponent<ReactIcons.ISvgIconProps>)  => (
                      <Icon key={item.key} aria-label={ 'MoreVertical'?.replace('', '') }  />
                  ))
                }
                   
                    {/* <IconButton
                     // className = { classNames.workspaceImage } //{styles.workspaceImage}
                      iconProps={{ iconName: "MoreVerticalIcon" }}
                      aria-label = { iconName 'MoreVerticalIcon'}
                    /> */}
                      </HoverCard>
                      // </div>
                     );
                },
              },
              {
                key: 'column4',
                name: 'Business Department',
                fieldName: 'businessDepartment',
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
                data: 'number',
                onRender: (item: IWorkspace) => {
                  return <span key={item.key}>{item.businessDepartment}</span>;
                },
                isPadded: true,
              },
              {
                key: 'column5',
                name: 'Business Owner',
                fieldName: 'businessOwner',
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
                data: 'number',
                onRender: (item: IWorkspace) => {
                  return <span >{item.businessOwner}</span>;
                },
                isPadded: true,
              },
              {
                key: 'column6',
                name: 'Status',
                fieldName: 'status',
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
                onRender: (item: IWorkspace) => {
                  return <span>{item.status}</span>;
                },
                isPadded: true,
              },
              {
                key: 'column7',
                name: 'Type',
                fieldName: 'type',
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                isCollapsible: true,
                data: 'number',
                onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
                onRender: (item: IWorkspace) => {
                  return <span key={item.key}>{item.type}</span>;
                },
              },
              {
                key: 'column8',
                name: 'Classification',
                fieldName: 'classification',
                minWidth: 70,
                maxWidth: 90,
                isResizable: true,
                isCollapsible: true,
                data: 'number',
                onColumnClick: (ev, columns) =>  this._onColumnContextMenu(columns, ev),
                onRender: (item: IWorkspace) => {
                  return <span key={item.key}>{item.classification}</span>;
                },
              },
        ]; 
        
        this.state = {
          displayItems: [],
          serachItem : [],
          itemsList : [],
          sortItemsDetails : [],
          columns: columns,
          contextualMenuProps:undefined,
          // selectionDetails: this._getSelectionDetails(),
          isModalSelection: false,
          isCompactMode: false,
          announcedMessage: undefined,
          userIsAdmin : false,
          hasMore : true,
          isPanelOpen : false,
          isPanelClose : true,
          checkSearchItem : false,
          itemArrayAppend : 20,
        };
    }

    private _onColumnContextMenu = (column: IColumn, ev: React.MouseEvent<HTMLElement>): void => {
      this.setState({
          contextualMenuProps: this._getContextualMenuProps(ev, column),
        });
    };

    private _getContextualMenuProps(ev: React.MouseEvent<HTMLElement>, column: IColumn): IContextualMenuProps {
      const items = [
        {
          key: 'aToZ',
          name: 'A to Z',
          iconProps: { iconName: 'SortUp' },
          canCheck: true,
          checked: column.isSorted && !column.isSortedDescending,
          onClick: () => this._onColumnClick(column, false),
        },
        {
          key: 'zToA',
          name: 'Z to A',
          iconProps: { iconName: 'SortDown' },
          canCheck: true,
          checked: column.isSorted && column.isSortedDescending,
          onClick: () => this._onColumnClick(column,true ),
        },
        
      ];
      if(column.name !== 'Name'){
        items.push({
          key: 'filter',
          name: 'Filter by',
          iconProps: { iconName: 'Filter' },
          canCheck: true,
          checked: column.isFiltered,
          onClick: () => this.onFilterColumn(column),
        });
      }
      return {
        items: items,
        target: ev.currentTarget as HTMLElement,
        directionalHint: DirectionalHint.bottomLeftEdge,
        gapSpace: 10,
        isBeakVisible: true,
        onDismiss: this.onContextualMenuDismissed,
      };
    }

    private onFilterColumn = (column: IColumn): void => {

      // var uniqueVals = [], enabledVals = [];
      // var workspacesUnfiltered, workspaces;
      
  
      this.setState({
        isPanelOpen: true,
        isPanelClose : true,
      });
    }

  private onContextualMenuDismissed = (): void => {
    this.setState({
      contextualMenuProps: undefined,
    });
  }

    private _onColumnClick = (column: IColumn , checkOrder : boolean): void => {
      const { columns, sortItemsDetails } = this.state;
      const newColumns: IColumn[] = columns.slice();
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
      });
    };
    

    public onRenderPlainCard(){
      return(
        <div>
          <Button
          text="Edit"
          //className= {styles.createNewButton}
          //onClick={() => this.setState({ currentItem: item, dialog: "Delete" })}
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

    public async componentDidMount(){

      await this._getUserRole().then((teamsUserRoleStatus:boolean)  => {
        if(teamsUserRoleStatus === true){
          this.setState({
            userIsAdmin : true
          })
          console.log("Teams User Role status : " + this.state.userIsAdmin );
        }
        else {
          this.setState({
            userIsAdmin : false
          });
        }
      });

      await this._getAllPublicTeams().then((teamsDetails : any[]) => {
        console.log("Component Teams Log" + teamsDetails );
        //if(teamsDetails.status === ''){}
        // this._allItems = teamsDetails;
        this.setState({
          displayItems: teamsDetails.slice(0,this.state.itemArrayAppend),
          serachItem : teamsDetails,
          itemsList : teamsDetails,
          sortItemsDetails : teamsDetails,
        });
      });
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
        <div>
            {
              this.state.userIsAdmin ? <div>
                  <div>
                  <Label style={{fontWeight:"bold"}}>Teams</Label>

                  <TextField label="Search Teams Here :" onScroll  = {this.fetchMoreData} onChange={ (event: any) => this._onChangeText(event)}
                   styles={controlStyles} />
                   {this.state.contextualMenuProps && <ContextualMenu {...this.state.contextualMenuProps} />}
                  </div>
                  <InfiniteScroll
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
                  >

                    <DetailsList
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
                          />

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
                  </InfiniteScroll>
                    {/* {console.log("Item Count :- " + this.state.items.length)}
                    <DetailsList
                        items={this.state.items}
                        columns={this.state.columns}
                        getKey={this._getKey}
                    />  */}

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
                :
              <div>
                  User Unauthorized 
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
        this.props.instance.acquireTokenSilent({
          ...loginRequest,
          account: this.props.accounts[0]
            }).then((response:any) => {
          callAllTeamsRequest(response.accessToken).then(response => response).then((data:any[]) =>
          {
            data.forEach(element => {
              if(element.status === 'Active' || element.status === 'Inactive'){
                items.push({
                  test: element.imageBlob,
                  key: element.id.toString(),
                  name: element.title,
                  businessDepartment: element.businessDepartment,
                  status: element.status,
                  type: '1',
                  classification: element.classification,
                  businessOwner : '1',
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

  export default WorkspaceDetails;
