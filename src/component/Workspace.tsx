import * as React  from 'react';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import * as ReactIcons from '@fluentui/react-icons-mdl2';
import { DetailsList, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { TooltipHost } from '@fluentui/react';
import {
  // IconButton,
  Button,
} from "office-ui-fabric-react/lib/Button";
import {
    HoverCard,
    HoverCardType,
    IPlainCardProps,
    
  } from "office-ui-fabric-react/lib/HoverCard";
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
    //selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
    userIsAdmin : boolean;
    hasMore : boolean;
    itemArrayAppend : number;
    checkSearchItem : boolean;
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
                // onColumnClick: this._onColumnClick,
                onRender: (item: IWorkspace) => (
                  <TooltipHost content={`${item.test} file`}>
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
                // onColumnClick: this._onColumnClick,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                data: 'number',
                onRender: (item: IWorkspace) => {
                  return <span> {item.name}</span>;
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
                      <Icon aria-label={ 'MoreVertical'?.replace('', '') }  />
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
                // onColumnClick: this._onColumnClick,
                data: 'number',
                onRender: (item: IWorkspace) => {
                  return <span>{item.businessDepartment}</span>;
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
                // onColumnClick: this._onColumnClick,
                data: 'number',
                onRender: (item: IWorkspace) => {
                  return <span>{item.businessOwner}</span>;
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
                // onColumnClick: this._onColumnClick,
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
                // onColumnClick: this._onColumnClick,
                onRender: (item: IWorkspace) => {
                  return <span>{item.type}</span>;
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
                // onColumnClick: this._onColumnClick,
                onRender: (item: IWorkspace) => {
                  return <span>{item.classification}</span>;
                },
              },
        ]; 
        
        this.state = {
          displayItems: [],
          serachItem : [],
          itemsList : [],
          columns: columns,
          // selectionDetails: this._getSelectionDetails(),
          isModalSelection: false,
          isCompactMode: false,
          announcedMessage: undefined,
          userIsAdmin : false,
          hasMore : true,
          checkSearchItem : false,
          itemArrayAppend : 20,
        };
    }
    

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
          })
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
                            onRenderItemColumn = { (item) => {
                                return item
                            } }
                            // getKey={this._getKey}
                            // setKey="set"
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
            </div>
                :
              <div>
                  User Unauthorized 
              </div>
            }
          </div>
      )
    }
    

    private _getKey(item: any, index?: number): string {
      return item.key ;
    }

    public _onChangeText = (ev: any): void => {

      let testData = this.state.serachItem ;
      let searchData = ev.target.value !== ""  ?  testData.filter(i => i.name.toLowerCase().startsWith(ev.target.value.toLowerCase())) : testData;
      
       this.setState({
        displayItems: searchData.slice(0, this.state.itemArrayAppend),
        itemsList : searchData,
        checkSearchItem : true,
        hasMore : true,
       },
        () => console.log(this.state.displayItems));
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
  export default WorkspaceDetails;
