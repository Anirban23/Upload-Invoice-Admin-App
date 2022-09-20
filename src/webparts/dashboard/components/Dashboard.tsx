import * as React from 'react';
import styles from './Dashboard.module.scss';
import { IDashboardProps } from './IDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import 'office-ui-fabric-react/dist/css/fabric.css';
import { DefaultButton, Icon, IIconProps, personaPresenceSize, Pivot, PivotItem, TextField } from 'office-ui-fabric-react';

import { AgGridColumn, AgGridReact } from 'ag-grid-react';
import 'ag-grid-community/dist/styles/ag-grid.css';
import 'ag-grid-community/dist/styles/ag-theme-alpine.css';

import { GridApi, SelectionChangedEvent } from 'ag-grid-community';
import * as pnp from 'sp-pnp-js'

export default class Dashboard extends React.Component<IDashboardProps, any> {
  [x: string]: any;
  public constructor(props: IDashboardProps) {
    super(props);
    this.state = {
      AllAssignmentData: [],
      AllPaymentData: [],
      DeleteList: '',
      DeleteID: '',
      columnDefAssignment: [
        {
          headerName: "",
          field: "ID", width: 50,
          cellRendererFramework: (item: any) => <div>
            <Icon iconName='Delete' onClick={() => this.onDeleteButtonClick(item, "Assignment")} style={{paddingRight:'5px'}}/>
            <Icon iconName='Edit' onClick={() => this.onDeleteButtonClick(item, "Assignment")}  style={{paddingLeft:'5px'}}/>
          </div>
        },
        { headerName: "Assignment Id", field: "AssignmentId", hide: false, width: 150, },
        { headerName: "Consultant ID", field: "ConsultantID", hide: false, width: 150, },
        { headerName: "Project Id", field: "ProjectId", hide: false, width: 150, },
        { headerName: "Approver", field: "Approver", hide: false, width: 150, },
        { headerName: "Consultant Name", field: "ConsultantName", hide: false, width: 150, },
      ],
      columnDefPayment: [
        {
          headerName: "",
          field: "ID", width: 50,
          cellRendererFramework: (item: any) => <div>
            <Icon iconName='Delete' onClick={() => this.onDeleteButtonClick(item, "PaymentTerms")} />
            <Icon iconName='Edit' onClick={() => this.onDeleteButtonClick(item, "PaymentTerms")} />
          </div>
        },
        { headerName: "Consultant Name", field: "ConsultantName", hide: false, width: 150, },
        { headerName: "Payment Terms", field: "PaymentTerms", hide: false, width: 150, },
        { headerName: "Finance Representative", field: "FinanceRepresentative", hide: false, width: 150, },
        { headerName: "Assignment Id", field: "AssignmentID", hide: false, width: 150, },
      ],
      defaultColDef: {
        resizable: true,
        sortable: true,
        filter: true,
        //floatingFilter: true,
        //editable: true,
        unSortIcon: true,
        //suppressColumnMoveAnimation: true,
        wrapText: false,
        //autoHeight: true,
        //cellStyle: { 'white-space': 'normal', fontSize: '11px' },
        cellStyle: { fontSize: '10px', paddingLeft: '8px', paddingRight: '0px' },
        //cellStyle: (params: any) => { 
        // if (params.node.rowIndex % 2 === 1){
        //   return{backgroundColor: '#fff',fontSize: '10px', paddingLeft: '8px', paddingRight: '0px'};
        // }
        // else{
        //   return{backgroundColor: 'rgb(175 177 231)', fontSize: '10px', paddingLeft: '8px', paddingRight: '0px'};
        // }

        // },
        cellHeaderStyle: { fontSize: '11px' },
        wrapHeaderText: true,
        autoHeaderHeight: true,
        headerClass: { 'white-space': 'normal', fontSize: '110px' },
        pagination: true,
        paginationPageSize: 30,


      }
    }
  }

  public componentDidMount() {
    this.loadUploadInvoiceAdmin();
  }

  onGridReady = (params: any) => {
    this.gridApi = params.api;
    this.gridColumnApi = params.columnApi;
  };

  private onDeleteButtonClick = (e: any, listName: string) => {
    console.log(e.data);
    this.setState({DeleteList: listName, DeleteID: e.data.ID});
    var modal = document.getElementById("DeleteModel");
    modal.style.display = "block";
  }

  private async onDeleteRecord() {
    var res = await pnp.sp.web.lists.getByTitle(this.state.DeleteList).items.getById(this.state.DeleteID).delete();
    console.log(res);

    let selectedData = this.gridApi.getSelectedRows();
    this.gridApi.applyTransaction({ remove: [selectedData[0]] });

    this.DeleteRecordSpanClose();
  }

  private DeleteRecordSpanClose() {
    var modal = document.getElementById("DeleteModel");
    modal.style.display = "none";
  }

  private async loadUploadInvoiceAdmin() {
    const itemAssignment: any = await pnp.sp.web.lists.getByTitle("Assignment").items.select("AssignmentId", "ConsultantID", "ProjectId", "Approver/Title", "ConsultantName/Title", "ID").expand("Approver", "ConsultantName").get();
    let tempData = [];
    for (let i = 0; i < itemAssignment.length; i++) {
      let CN = '';
      if (itemAssignment[i].ConsultantName !== undefined) {
        CN = itemAssignment[i].ConsultantName.Title;
      }

      tempData.push({
        "AssignmentId": itemAssignment[i].AssignmentId,
        "ConsultantID": itemAssignment[i].ConsultantID,
        "ProjectId": itemAssignment[i].ProjectId,
        "Approver": itemAssignment[i].Approver.Title,
        "ConsultantName": CN,
        "ID": itemAssignment[i].ID
      });
    }
    this.setState({ AllAssignmentData: tempData });


    const itemPayment: any = await pnp.sp.web.lists.getByTitle("PaymentTerms").items.select("ConsultantName/Title", "PaymentTerms", "FinanceRepresentative/Title", "AssignmentID", "ID").expand("FinanceRepresentative", "ConsultantName").get();
    let tempDataPayment = [];
    for (let i = 0; i < itemPayment.length; i++) {
      tempDataPayment.push({
        "ConsultantName": itemPayment[i].ConsultantName.Title,
        "PaymentTerms": itemPayment[i].PaymentTerms,
        "FinanceRepresentative": itemPayment[i].FinanceRepresentative.Title,
        "AssignmentID": itemPayment[i].AssignmentID,
        "ID": itemPayment[i].ID,
      });
    }
    this.setState({ AllPaymentData: tempDataPayment });


  }

  public render(): React.ReactElement<IDashboardProps> {


    return (
      <div>
        <div id="DeleteModel" className={styles.modal}>
          <div className={styles['modal-content']} style={{ width: '50%' }}>
            <div className={styles["modal-header"]}>
              <span className={styles["close"]} onClick={() => this.spanClose()}>&times;</span>
              <h2>Delete Confirmation</h2>
            </div>
            <div className={styles["modal-body"]}>
              <div className="ms-Grid-row" style={{border:'1px solid black',padding: '10px'}}>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12" style={{height:'50px', marginTop: '10px'}}>
                  <label>Are you sure to delete?</label>
                </div>
                <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
                  <div className="ms-Grid-row">
                    <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                      <DefaultButton
                        text="Cancel"
                        //iconProps={addIcon}
                        onClick={() => this.DeleteRecordSpanClose()}
                        style={{width: '100%'}}
                      //label="Submit"
                      //allowDisabledFocus
                      //disabled={disabled}
                      //checked={checked} 
                      />
                    </div>
                    <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                      <DefaultButton
                        text="Delete"
                        //iconProps={addIcon}
                        onClick={() => this.onDeleteRecord()}
                        style={{width: '100%'}}
                      //label="Submit"
                      //allowDisabledFocus
                      //disabled={disabled}
                      //checked={checked} 
                      />
                    </div>
                  </div>

                </div>
              </div>
            </div>
            <div className={styles["modal-footer"]}>
              <h3></h3>
            </div>
          </div>
        </div>


        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md12 ms-lg12">
            <Pivot>
              <PivotItem headerText="Assignment Id">
                <div className={"ag-theme-alpine"} style={{ height: "450px" }}>
                  <AgGridReact
                    columnDefs={this.state.columnDefAssignment}
                    defaultColDef={this.state.defaultColDef}
                    rowData={this.state.AllAssignmentData}
                    //pagination={true} paginationPageSize={5}
                    animateRows={true}
                    onGridReady={this.onGridReady}
                    rowSelection={'single'}
                  //onSelectionChanged={e => this.AGChange(e)}
                  >
                  </AgGridReact>
                </div>
              </PivotItem>
              <PivotItem headerText="Payment Terms">
                <div className={"ag-theme-alpine"} style={{ height: "450px" }}>
                  <AgGridReact
                    columnDefs={this.state.columnDefPayment}
                    defaultColDef={this.state.defaultColDef}
                    rowData={this.state.AllPaymentData}
                    //pagination={true} paginationPageSize={5}
                    animateRows={true}
                    onGridReady={this.onGridReady}
                    rowSelection={'single'}
                  //onSelectionChanged={e => this.AGChange(e)}
                  >
                  </AgGridReact>
                </div>
              </PivotItem>
            </Pivot>
          </div>
        </div>
      </div>
    );
  }
}
