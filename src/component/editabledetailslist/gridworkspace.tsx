import { DetailsListLayoutMode, mergeStyles, mergeStyleSets, SelectionMode, TextField } from '@fluentui/react';
import { render } from '@testing-library/react';
import { EditableGrid, EditControlType, IColumnConfig, EventEmitter, EventType, NumberAndDateOperators } from 'fluentui-editable-grid';
import { Fabric } from 'office-ui-fabric-react';
import * as React from 'react';
import { useState } from 'react';
import { TooltipHost } from '@fluentui/react';

interface IWorkspaceprops {
    workspaceProps:any;
    workspaceState: any;
}

interface IWorkspaceColumns {
    workspaceColumns : IColumnConfig[];
}

const classNames = mergeStyleSets({
    controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap',
    }
});

interface IWorkspace {
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
    }

const classNamesForIcon = mergeStyleSets({
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


class WorkspaceDetailsList extends React.Component<IWorkspaceprops,IWorkspaceColumns> {
    
    constructor(props: IWorkspaceprops){
        super(props);
        const columnsList: IColumnConfig[] = [
            {
                key:"test",
                text:"",
                name:"test",
                className: classNamesForIcon.fileIconCell,
                iconClassName: classNamesForIcon.fileIconHeaderIcon,
                iconName: 'Page',
                isIconOnly: true,
                fieldName: 'test',
                minWidth: 20,
                maxWidth: 20,
                dataType:'icon',
                onRender: (item: IWorkspace ) => (
                    <div className="ms-Grid-col ms-sm6 ms-md4 ms-lg1"> 
                      <TooltipHost key={item.key} content={`${item.test} file`}>
                        <img src={item.test} className={classNamesForIcon.fileIconImg} alt={`${item.test} file icon`} /> 
                      </TooltipHost>
                    </div>
                  ),
            },
            // {
            //     key: 'customerhovercol',
            //     name: 'Custom Hover Column',
            //     text: 'Custom Hover Column',
            //     editable: true,
            //     dataType: 'string',
            //     minWidth: 100,
            //     maxWidth: 100,
            //     isResizable: true,
            //     includeColumnInExport: false,
            //     includeColumnInSearch: false,
            //     applyColumnFilter: false,
            //     disableSort: true,
            //     //hoverComponentOptions: { enable:true, hoverChildComponent: <CellHover customProps={{ someProp: '' }} /> }
            // },
            {
                key: 'name',
                name: 'Name',
                text: 'Name',
                editable: true,
                dataType: 'string',
                minWidth: 100,
                maxWidth: 100,
                isResizable: true,
                includeColumnInExport: true,
                includeColumnInSearch: true,
                applyColumnFilter: true
            },
            {
                key: 'businessDepartment',
                name: 'businessDepartment',
                text: 'Business Department',
                editable: true,
                dataType: 'string',
                minWidth: 100,
                maxWidth: 100,
                isResizable: true,
                includeColumnInExport: true,
                includeColumnInSearch: true,
                applyColumnFilter: true
            },
            {
                key: 'businessOwner',
                name: 'businessOwner',
                text: 'Business Owner',
                editable: true,
                dataType: 'string',
                minWidth: 100,
                maxWidth: 100,
                isResizable: true,
                includeColumnInExport: true,
                includeColumnInSearch: true,
                inputType: EditControlType.MultilineTextField,
                applyColumnFilter: true
            },
            {
                key: 'status',
                name: 'status',
                text: 'Status',
                editable: true,
                dataType: 'string',
                minWidth: 100,
                maxWidth: 100,
                isResizable: true,
                includeColumnInExport: false,
                includeColumnInSearch: true,
                maxLength:5,
                applyColumnFilter: true,
                cellStyleRule: { 
                    enable: true, 
                    rule: { 
                        operator : NumberAndDateOperators.LESSTHAN, 
                        value: 50000 
                    }, 
                    whenTrue: { textColor: '#EF5350', fontWeight: 'bold' },
                    whenFalse: { textColor: '#9CCC65' }
                }
            },
            {
                key: 'type',
                name: 'type',
                text: 'Type',
                editable: true,
                dataType: 'string',
                minWidth: 150,
                maxWidth: 150,
                isResizable: true,
                includeColumnInExport: true,
                includeColumnInSearch: true,
                inputType: EditControlType.Date
            },
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
                inputType: EditControlType.DropDown,
                dropdownValues: [
                    { key: 'weekly', text: 'Weekly' },
                    { key: 'biweekly', text: 'Bi-Weekly' },
                    { key: 'monthly', text: 'Monthly' }
                ]
            }
            // ,{
            //     key: 'employmenttype',
            //     name: 'Employment Type',
            //     text: 'Employment Type',
            //     editable: true,
            //     dataType: 'string',
            //     minWidth: 200,
            //     maxWidth: 200,
            //     isResizable: true,
            //     includeColumnInExport: true,
            //     includeColumnInSearch: true,
            //     inputType: EditControlType.Picker,
            //     pickerOptions: {
            //         pickerTags: ['Employment Type1', 'Employment Type2', 'Employment Type3', 'Employment Type4', 'Employment Type5', 'Employment Type6', 'Employment Type7', 'Employment Type8', 'Employment Type9', 'Employment Type10', 'Employment Type11', 'Employment Type12'],
            //         minCharLimitForSuggestions: 2,
            //         tagsLimit: 1,
            //         pickerDescriptionOptions: { 
            //             enabled: true, 
            //             values: [
            //                 { key: 'Employment Type1', description: 'Employment Type1 Description'},
            //                 { key: 'Employment Type2', description: 'Employment Type2 Description'},
            //                 { key: 'Employment Type3', description: 'Employment Type3 Description'},
            //                 { key: 'Employment Type4', description: 'Employment Type4 Description'},
            //                 { key: 'Employment Type5', description: 'Employment Type5 Description'},
            //                 { key: 'Employment Type6', description: 'Employment Type6 Description'},
            //                 { key: 'Employment Type7', description: 'Employment Type7 Description'},
            //                 { key: 'Employment Type8', description: 'Employment Type8 Description'},
            //                 { key: 'Employment Type9', description: 'Employment Type9 Description'},
            //                 { key: 'Employment Type10', description: 'Employment Type10 Description'},
            //                 { key: 'Employment Type11', description: 'Employment Type11 Description'},
            //                 { key: 'Employment Type12', description: 'Employment Type12 Description'},
            //         ] },
            //         //suggestionsRule: StringOperators.STARTSWITH
            //     }
            // }
        ];
        this.state = {
            workspaceColumns : columnsList
        }
    }
    

    public async componentDidMount() {

    }

    render(){
        return(
            <Fabric>
                    <div className={classNames.controlWrapper}>
                        <TextField placeholder='Search Grid' className={mergeStyles({ width: '60vh', paddingBottom:'10px' })} onChange={(event) => EventEmitter.dispatch(EventType.onSearch, event)}/>
                    </div>
                    <EditableGrid
                        id={1}
                        columns={this.state.workspaceColumns}
                        items={this.props.workspaceState.itemsList}
                        // enableCellEdit={true}
                        enableExport={true}
                        //enableTextFieldEditMode={true}
                        //enableTextFieldEditModeCancel={true}
                        //enableGridRowsDelete={true}
                        //enableGridRowsAdd={true}
                        //height={'70vh'}
                        //width={'140vh'}
                        //position={'relative'}
                        //enableUnsavedEditIndicator={true}
                        //onGridSave={onGridSave}
                        // enableGridReset={true}
                        enableColumnFilters={true}
                        //enableColumnFilterRules={true}
                        //enableRowAddWithValues={{enable : true, enableRowsCounterInPanel : true}}
                        //layoutMode={DetailsListLayoutMode.justified}
                        //selectionMode={SelectionMode.multiple}
                        //enableRowEdit={true}
                        //enableRowEditCancel={true}
                        //enableBulkEdit={true}
                        //enableColumnEdit={true}
                        //enableSave={true}
                    />
            </Fabric>
        )
    }
}

const Consumer = () => {
    const classNames = mergeStyleSets({
        controlWrapper: {
        display: 'flex',
        flexWrap: 'wrap',
        }
    });



const [items, setItems] = useState<any[]>([]);
const columns: IColumnConfig[] = [
    {
        key: 'id',
        name: 'ID',
        text: 'ID',
        editable: false,
        dataType: 'number',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        applyColumnFilter: true,
        disableSort: true
    },
    {
        key: 'customerhovercol',
        name: 'Custom Hover Column',
        text: 'Custom Hover Column',
        editable: true,
        dataType: 'string',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        includeColumnInExport: false,
        includeColumnInSearch: false,
        applyColumnFilter: false,
        disableSort: true,
        //hoverComponentOptions: { enable:true, hoverChildComponent: <CellHover customProps={{ someProp: '' }} /> }
    },
    {
        key: 'name',
        name: 'Name',
        text: 'Name',
        editable: true,
        dataType: 'string',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        applyColumnFilter: true
    },
    {
        key: 'age',
        name: 'Age',
        text: 'Age',
        editable: true,
        dataType: 'number',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        applyColumnFilter: true
    },
    {
        key: 'designation',
        name: 'Designation',
        text: 'Designation',
        editable: true,
        dataType: 'string',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        inputType: EditControlType.MultilineTextField,
        applyColumnFilter: true
    },
    {
        key: 'salary',
        name: 'Salary',
        text: 'Salary',
        editable: true,
        dataType: 'number',
        minWidth: 100,
        maxWidth: 100,
        isResizable: true,
        includeColumnInExport: false,
        includeColumnInSearch: true,
        maxLength:5,
        applyColumnFilter: true,
        cellStyleRule: { 
            enable: true, 
            rule: { 
                operator : NumberAndDateOperators.LESSTHAN, 
                value: 50000 
            }, 
            whenTrue: { textColor: '#EF5350', fontWeight: 'bold' },
            whenFalse: { textColor: '#9CCC65' }
        }
    },
    {
        key: 'dateofjoining',
        name: 'Date of Joining',
        text: 'Date of Joining',
        editable: true,
        dataType: 'date',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        inputType: EditControlType.Date
    },
    {
        key: 'payrolltype',
        name: 'Payroll Type',
        text: 'Payroll Type',
        editable: true,
        dataType: 'string',
        minWidth: 150,
        maxWidth: 150,
        isResizable: true,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        inputType: EditControlType.DropDown,
        dropdownValues: [
            { key: 'weekly', text: 'Weekly' },
            { key: 'biweekly', text: 'Bi-Weekly' },
            { key: 'monthly', text: 'Monthly' }
        ]
    },
    {
        key: 'employmenttype',
        name: 'Employment Type',
        text: 'Employment Type',
        editable: true,
        dataType: 'string',
        minWidth: 200,
        maxWidth: 200,
        isResizable: true,
        includeColumnInExport: true,
        includeColumnInSearch: true,
        inputType: EditControlType.Picker,
        pickerOptions: {
            pickerTags: ['Employment Type1', 'Employment Type2', 'Employment Type3', 'Employment Type4', 'Employment Type5', 'Employment Type6', 'Employment Type7', 'Employment Type8', 'Employment Type9', 'Employment Type10', 'Employment Type11', 'Employment Type12'],
            minCharLimitForSuggestions: 2,
            tagsLimit: 1,
            pickerDescriptionOptions: { 
                enabled: true, 
                values: [
                    { key: 'Employment Type1', description: 'Employment Type1 Description'},
                    { key: 'Employment Type2', description: 'Employment Type2 Description'},
                    { key: 'Employment Type3', description: 'Employment Type3 Description'},
                    { key: 'Employment Type4', description: 'Employment Type4 Description'},
                    { key: 'Employment Type5', description: 'Employment Type5 Description'},
                    { key: 'Employment Type6', description: 'Employment Type6 Description'},
                    { key: 'Employment Type7', description: 'Employment Type7 Description'},
                    { key: 'Employment Type8', description: 'Employment Type8 Description'},
                    { key: 'Employment Type9', description: 'Employment Type9 Description'},
                    { key: 'Employment Type10', description: 'Employment Type10 Description'},
                    { key: 'Employment Type11', description: 'Employment Type11 Description'},
                    { key: 'Employment Type12', description: 'Employment Type12 Description'},
            ] },
            //suggestionsRule: StringOperators.STARTSWITH
        }
    }
];

const SetDummyData = () : void => {
    const dummyData = [
        {
            id: "1",
            customerhovercol: 'Hover Me',
            name: "Name1",
            age:32,
            designation:'Designation1',
            salary:57000,
            dateofjoining:'2010-04-01T14:57:10',
            payrolltype: 'Weekly',
            employmenttype: 'Employment Type11'
        },
        {
            id: "2",
            customerhovercol: 'Hover Me',
            name: "Name2",
            age:27,
            designation:'Designation2',
            salary:42000,
            dateofjoining:'2014-06-09T14:57:10',
            payrolltype: 'Monthly',
            employmenttype: 'Employment Type4'
        },
        {
            id: "3",
            customerhovercol: 'Hover Me',
            name: "Name3",
            age:35,
            designation:'Designation3',
            salary:75000,
            dateofjoining:'2005-07-02T14:57:10',
            payrolltype: 'Weekly',
            employmenttype: 'Employment Type7'
        },
        {
            id: "4",
            customerhovercol: 'Hover Me',
            name: "Name4",
            age:30,
            designation:'Designation4',
            salary:49000,
            dateofjoining:'2019-04-01T14:57:10',
            payrolltype: 'Bi-Weekly',
            employmenttype: 'Employment Type2'
        }
    ];
    setItems(dummyData);
}

React.useEffect(() => {
    SetDummyData();
}, []);

return (
    <Fabric>
        <div className={classNames.controlWrapper}>
            <TextField placeholder='Search Grid' className={mergeStyles({ width: '60vh', paddingBottom:'10px' })} onChange={(event) => EventEmitter.dispatch(EventType.onSearch, event)}/>
        </div>
        <EditableGrid
            id={1}
            columns={columns}
            items={items}
            enableCellEdit={true}
            enableExport={true}
            enableTextFieldEditMode={true}
            enableTextFieldEditModeCancel={true}
            enableGridRowsDelete={true}
            enableGridRowsAdd={true}
            height={'70vh'}
            width={'140vh'}
            position={'relative'}
            enableUnsavedEditIndicator={true}
            //onGridSave={onGridSave}
            enableGridReset={true}
            enableColumnFilters={true}
            enableColumnFilterRules={true}
            enableRowAddWithValues={{enable : true, enableRowsCounterInPanel : true}}
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.multiple}
            enableRowEdit={true}
            enableRowEditCancel={true}
            enableBulkEdit={true}
            enableColumnEdit={true}
            enableSave={true}
        />
    </Fabric>
);
};
export default WorkspaceDetailsList;
// export default Consumer;