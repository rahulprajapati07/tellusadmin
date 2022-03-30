import * as React  from 'react';
//import { useMsal } from "@azure/msal-react";
import { TextField } from '@fluentui/react/lib/TextField';
//import { Toggle } from '@fluentui/react/lib/Toggle';
// import { Announced } from '@fluentui/react/lib/Announced';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn } from '@fluentui/react/lib/DetailsList';
//import { MarqueeSelection } from '@fluentui/react/lib/MarqueeSelection';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';
import { TooltipHost } from '@fluentui/react';
import { loginRequest } from "./component/authConfig";
import {callAllTeamsRequest,canUserRestoreTeams}  from "./component/graph";//{ callMsGraph,callMsGraphGroup,callAllTeamsRequest,callGetPublicTeams }
//import Button from "react-bootstrap/Button";
import { Label } from '@fluentui/react/lib/Label';
//import { Icon } from '@fluentui/react/lib/Icon';
import * as ReactIcons from '@fluentui/react-icons-mdl2';
//import InfiniteScroll from "react-infinite-scroll-component";


import {
  HoverCard,
  HoverCardType,
  IPlainCardProps,
  
} from "office-ui-fabric-react/lib/HoverCard";
import {
  // IconButton,
  Button,
} from "office-ui-fabric-react/lib/Button";
// import { GetMyPublicTeams } from './component/BackendService';


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

  // const classes = mergeStyleSets({
  //   cell: {
  //     display: 'flex',
  //     flexDirection: 'column',
  //     alignItems: 'center',
  //     margin: '80px',
  //     float: 'left',
  //     height: '50px',
  //     width: '50px',
  //   },
  //   icon: {
  //     fontSize: '50px',
  //   },
  //   code: {
  //     background: '#f2f2f2',
  //     borderRadius: '4px',
  //     padding: '4px',
  //   },
  //   navigationText: {
  //     width: 100,
  //     margin: '0 5px',
  //   },
  // });

  const icons = Object.keys(ReactIcons).reduce((acc: React.FC[], exportName) => {
    if ((ReactIcons as any)[exportName]?.displayName) {
      if(exportName === "MoreVerticalIcon"){
        acc.push((ReactIcons as any)[exportName] as React.FunctionComponent);
      }
    }
  
    return acc;
  }, []);
  
//let allItems : IDocument[] = [];

  export interface IDetailsListDocumentsExampleState {
    columns: IColumn[];
    items: IDocument[];
    serachItem: IDocument[];
    itemsList : IDocument[];
    //selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
    announcedMessage?: string;
    userIsAdmin : boolean;
    hasMore : boolean;
  }
  
  export interface IDocument {
    key: number;
    test: string;
    name: string;
    businessDepartment: string;
    status: string;
    type: string;
    classification : string;
    businessOwner : string;
  }

  export interface IPublicTeams {
    name: string;
    businessDepartment: string;
    status: string;
    type: string;
    classification : string;
    businessOwner : string;
  }

  interface IDetailsListDocumentsExampleProps {
    instance : any;
    accounts : any;
  }
  
 

  class DetailsListDemo extends React.Component<IDetailsListDocumentsExampleProps, IDetailsListDocumentsExampleState> {
    //private _selection: Selection;
    private _allItems: IDocument[];

    //private _allTeamsData : IDocument[];
    //private _allpublicTeams : IPublicTeams[];
    
    constructor(props: IDetailsListDocumentsExampleProps, state: IDetailsListDocumentsExampleState) {
      super(props);
      //getAllMyPublicTeams(this.props.instance,this.props.accounts);
      //this._allItems = _generateDocuments();
      //this._allItems = getAllMyPublicTeams(this.props.instance,this.props.accounts);
      this._allItems = [];
      
      
      //this._onChangeText = this._onChangeText.bind(this);
      const columns: IColumn[] = [
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
          onRender: (item: IDocument) => (
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
          onRender: (item: IDocument) => {
            return <span>{item.name}</span>;
          },
          isPadded: true,
        },
        {
          key: "column3",
          name: "",
          fieldName: "Options",
          minWidth: 10,
          maxWidth: 10,
          onRender:(item: IDocument) => {
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
          onRender: (item: IDocument) => {
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
          onRender: (item: IDocument) => {
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
          onRender: (item: IDocument) => {
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
          onRender: (item: IDocument) => {
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
          onRender: (item: IDocument) => {
            return <span>{item.classification}</span>;
          },
        },
        
      ];
      
      // this._selection = new Selection({
      //   onSelectionChanged: () => {
      //     this.setState({
      //       selectionDetails: this._getSelectionDetails(),
      //     });
      //   },
      // });
      
      this.state = {
        items: [],
        serachItem : [],
        itemsList : Array.from({ length: 20 }),
        columns: columns,
        // selectionDetails: this._getSelectionDetails(),
        isModalSelection: false,
        isCompactMode: false,
        announcedMessage: undefined,
        userIsAdmin : false,
        hasMore : true
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
        this._allItems = teamsDetails;

        this.setState({
          items: teamsDetails,
          serachItem : teamsDetails
        });
      });
    }

    // public fetchMoreData = () => {

    //   if (this.state.items.length >= 1000) {
    //     this.setState({ hasMore: false });
    //     return;
    //   }
    //   let numberOfItem = 20;

    //   // for(var i=0;i<20;i++){
    //   //   numberOfItem = numberOfItem;
    //   //   //let itemsFromArray = this.state.items.slice(0,numberOfItem);
        
    //   //     this.setState({
    //   //       itemsList: this.state.items.slice(0,numberOfItem)
    //   //     });
        
    //   //   numberOfItem ++;
    //   // }

    //   // a fake async api call like which sends
    //   // 20 more records in 1.5 secs
    //   setTimeout(() => {
    //     this.setState({
    //       itemsList: this.state.items.slice(0,numberOfItem)
    //     });
    //   }, 1500);
    //   numberOfItem = numberOfItem + 20;
    // };

    public render(){
      //const { columns, isCompactMode, items, selectionDetails, isModalSelection, announcedMessage , userIsAdmin } = this.state;
      
      console.log("Teams User Role :- " + this.state.userIsAdmin);
        return (
          <div>
            {
              this.state.userIsAdmin ? <div>
                  <div>
                  <Label style={{fontWeight:"bold"}}>Teams</Label>
                  <TextField label="Filter by name:" onChange={ (event: any) => this._onChangeText(event)} styles={controlStyles} />
                  
                  </div>
                  
                   
                    {console.log("Item Count :- " + this.state.items.length)}
                    <DetailsList
                        items={this.state.items}
                        compact={this.state.isCompactMode}
                        columns={this.state.columns}
                        selectionMode={SelectionMode.multiple}
                        getKey={this._getKey}
                        setKey="multiple"
                        layoutMode={DetailsListLayoutMode.justified}
                        isHeaderVisible={true}
                        //selection={this._selection}
                        selectionPreservedOnEmptyClick={true}
                        onItemInvoked= {this._onItemInvoked}
                        enterModalSelectionOnTouch={true}
                        ariaLabelForSelectionColumn="Toggle selection"
                        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                        checkButtonAriaLabel="select row"
            /> 
            </div>
                :
              <div>
                  User Unauthorized 
              </div>
            }
          </div> 
        );
    }
  
    // public componentDidUpdate(previousProps: any, previousState: IDetailsListDocumentsExampleState) {
    //   if (previousState.isModalSelection !== this.state.isModalSelection && !this.state.isModalSelection) {
    //     this._selection.setAllSelected(false);
        
    //   }
    // }
  
    private _getKey(item: any, index?: number): string {
      console.log("Item Key :- " + item.key);
      return item.key;
    }
  
    // private _onChangeCompactMode = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
    //   this.setState({ isCompactMode: checked });
    // };
  
    // private _onChangeModalSelection = (ev: React.MouseEvent<HTMLElement>, checked: boolean): void => {
    //   this.setState({ isModalSelection: checked });
    // };
  
    public _onChangeText = (ev: any): void => {

     let testData = this._allItems ;
     let searchData = ev.target.value !== ""  ?  testData.filter(i => i.name.toLowerCase().startsWith(ev.target.value.toLowerCase())) : testData;

      this.setState({
        items: searchData
      },
       () => console.log(this.state.items));
    };
  
    private _onItemInvoked(item: any): void {
      console.log("Item invoked " + item);
      alert(`Item invoked: ${item.name}`);
    }
  
    // private _getSelectionDetails(): string {
    //   const selectionCount = this._selection.getSelectedCount();
  
    //   switch (selectionCount) {
    //     case 0:
    //       return 'No items selected';
    //     case 1:
    //       return '1 item selected: ' + (this._selection.getSelection()[0] as IDocument).name;
    //     default:
    //       return `${selectionCount} items selected`;
    //   }
    // }
  
    // private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    //   const { columns, items } = this.state;
    //   const newColumns: IColumn[] = columns.slice();
    //   const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
    //   newColumns.forEach((newCol: IColumn) => {
    //     if (newCol === currColumn) {
    //       currColumn.isSortedDescending = !currColumn.isSortedDescending;
    //       currColumn.isSorted = true;
    //       this.setState({
    //         announcedMessage: `${currColumn.name} is sorted ${
    //           currColumn.isSortedDescending ? 'descending' : 'ascending'
    //         }`,
    //       });
    //     } else {
    //       newCol.isSorted = false;
    //       newCol.isSortedDescending = true;
    //     }
    //   });
    //   const newItems = _copyAndSort(items, currColumn.fieldName!, currColumn.isSortedDescending);
    //   this.setState({
    //     columns: newColumns,
    //     items: newItems,
    //   });
    // };

    public _getAllPublicTeams = async () : Promise<IDocument[]>  =>  {
      
      return new Promise<any>((resolve, reject) => 
      {
        const items: IDocument[] = [];
        this.props.instance.acquireTokenSilent({
          ...loginRequest,
          account: this.props.accounts[0]
            }).then((response:any) => {
          callAllTeamsRequest(response.accessToken).then(response => response).then((data:any[]) =>
          {
            data.forEach(element => {
              if(element.status === 'Active' || element.status === 'Inactive'){
                items.push({
                  key: element.id,
                  test: element.imageBlob,
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

  export default DetailsListDemo;
  
  // function getAllMyPublicTeams(instance :any, accounts :any) {

  //   const items: IDocument[] = [];
  //     instance.acquireTokenSilent({
  //       ...loginRequest,
  //       account: accounts[0]
  //         }).then((response:any) => {
  //       callAllTeamsRequest(response.accessToken).then(response => response).then((data:any[]) =>
  //       {
  //         data.forEach(element => {
  //           items.push({
  //             key: element.imageBlob,
  //             name: element.title,
  //             businessDepartment: element.businessDepartment,
  //             status: element.status,
  //             type: element.visibility,
  //             classification: element.classification,
  //             businessOwner : element.ownerName,
  //           });
  //         });
  //       });;
  //     }); 
  //     return items;
  //   }
  
  // function _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
  //   const key = columnKey as keyof T;
  //   return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
  // }
  

  // function _generateDocuments() {
  //   const items: IDocument[] = [];

    
  //   for (let i = 0; i < 500; i++) {
  //     const randomDate = _randomDate(new Date(2012, 0, 1), new Date());
  //     const randomFileSize = _randomFileSize();
  //     const randomFileType = _randomFileIcon();
  //     let fileName = _lorem(2);
  //     fileName = fileName.charAt(0).toUpperCase() + fileName.slice(1).concat(`.${randomFileType.docType}`);
  //     let userName = _lorem(2);
  //     userName = userName
  //       .split(' ')
  //       .map((name: string) => name.charAt(0).toUpperCase() + name.slice(1))
  //       .join(' ');
  //     items.push({
  //       key: i.toString(),
  //       name: fileName,
  //       value: fileName,
  //       iconName: randomFileType.url,
  //       fileType: randomFileType.docType,
  //       modifiedBy: userName,
  //       dateModified: randomDate.dateFormatted,
  //       dateModifiedValue: randomDate.value,
  //       fileSize: randomFileSize.value,
  //       fileSizeRaw: randomFileSize.rawSize,
  //     });
  //   }
  //   return items;
  // }
  
  // const allTeams = () => {
  //   const { instance, accounts } = useMsal();
  //   function getAllMyPublicTeams()
  //     {
  //       instance.acquireTokenSilent({
  //           ...loginRequest,
  //           account: accounts[0]
  //       }).then((response) => {
  //           callAllTeamsRequest(response.accessToken).then(response => response).then((data:any[]) =>
  //           {
  //               console.log("Request all Teams data",data);
  //           });;
  //       });
  //     }
      
  // }
  
  
  // function _randomDate(start: Date, end: Date): { value: number; dateFormatted: string } {
  //   const date: Date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
  //   return {
  //     value: date.valueOf(),
  //     dateFormatted: date.toLocaleDateString(),
  //   };
  // }
  
  // const FILE_ICONS: { name: string }[] = [
  //   { name: 'accdb' },
  //   { name: 'audio' },
  //   { name: 'code' },
  //   { name: 'csv' },
  //   { name: 'docx' },
  //   { name: 'dotx' },
  //   { name: 'mpp' },
  //   { name: 'mpt' },
  //   { name: 'model' },
  //   { name: 'one' },
  //   { name: 'onetoc' },
  //   { name: 'potx' },
  //   { name: 'ppsx' },
  //   { name: 'pdf' },
  //   { name: 'photo' },
  //   { name: 'pptx' },
  //   { name: 'presentation' },
  //   { name: 'potx' },
  //   { name: 'pub' },
  //   { name: 'rtf' },
  //   { name: 'spreadsheet' },
  //   { name: 'txt' },
  //   { name: 'vector' },
  //   { name: 'vsdx' },
  //   { name: 'vssx' },
  //   { name: 'vstx' },
  //   { name: 'xlsx' },
  //   { name: 'xltx' },
  //   { name: 'xsn' },
  // ];
  
  // function _randomFileIcon(): { docType: string; url: string } {
  //   const docType: string = FILE_ICONS[Math.floor(Math.random() * FILE_ICONS.length)].name;
  //   return {
  //     docType,
  //     url: `https://static2.sharepointonline.com/files/fabric/assets/item-types/16/${docType}.svg`,
  //   };
  // }
  
  // function _randomFileSize(): { value: string; rawSize: number } {
  //   const fileSize: number = Math.floor(Math.random() * 100) + 30;
  //   return {
  //     value: `${fileSize} KB`,
  //     rawSize: fileSize,
  //   };
  // }
  
  // const LOREM_IPSUM = (
  //   'lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut ' +
  //   'labore et dolore magna aliqua ut enim ad minim veniam quis nostrud exercitation ullamco laboris nisi ut ' +
  //   'aliquip ex ea commodo consequat duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore ' +
  //   'eu fugiat nulla pariatur excepteur sint occaecat cupidatat non proident sunt in culpa qui officia deserunt '
  // ).split(' ');
  // let loremIndex = 0;
  // function _lorem(wordCount: number): string {
  //   const startIndex = loremIndex + wordCount > LOREM_IPSUM.length ? 0 : loremIndex;
  //   loremIndex = startIndex + wordCount;
  //   return LOREM_IPSUM.slice(startIndex, loremIndex).join(' ');
  // }