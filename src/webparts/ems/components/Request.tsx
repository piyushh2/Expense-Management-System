import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { DataGrid, GridColDef, GridPaginationModel } from "@mui/x-data-grid";
import { Button, Box, Typography } from "@mui/material";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import New from "./New";

interface IRequestProps {
  siteUrl: string;
  context: any;
}
interface EmployeeDetails {
  EmployeeID?: string;
  EmployeeName?: string;
  Department?: string;
  Country?: string;
  Manager?: string;
  RequestNo?: number;
  Currency?: string;
}

const Request: React.FC<IRequestProps> = (props) => {
  const userEmail = props.context.pageContext.user.email;
  const [paginationModel, setPaginationModel] = useState<GridPaginationModel>({ page: 0, pageSize: 10, });
  const [EmployeeDetails, setEmployeeDetails] = useState<EmployeeDetails>({});
  const [selectedRow, setSelectedRow] = useState<any>(null);
  const [selectedMenu, setSelectedMenu] = useState("");
  const [ExpenseData, setExpenseData] = useState([]);
  const [viewTable, setViewTable] = useState(false);
  const [neww, setNeww] = useState(false);
  const [view, setView] = useState(false);
  const [edit, setEdit] = useState(false);

  const getEmployeeDetails = useCallback(async () => {
    try {
      const listUrl = `${props.siteUrl}/_api/web/lists/getbytitle('Employee')/items`;
      const listResponse: SPHttpClientResponse = await props.context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );
      const res = await listResponse.json();
      const filteredEmployee = res.value.filter((item: any) => item.Email === userEmail);
      setEmployeeDetails(filteredEmployee[0]);
    }
    catch (error) {
      console.log(`Error: ${error}`);
    }
  }, [props.siteUrl, props.context, userEmail]);
  const getExpenseData = useCallback(async () => {
    try {
      const listUrl = `${props.siteUrl}/_api/web/lists/getbytitle('Expenses')/items`;
      const listResponse: SPHttpClientResponse = await props.context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );
      const res = await listResponse.json();
      if (res?.value?.length) setExpenseData(res.value);
    }
    catch (error) {
      console.log(`Error: ${error}`);
    }
  }, [props.siteUrl, props.context, userEmail]);
  const refreshData = useCallback(async () => {
    await Promise.all([getEmployeeDetails(), getExpenseData()]);
  }, [getEmployeeDetails, getExpenseData]);

  useEffect(() => {
    void refreshData();
  }, [selectedMenu, refreshData]);

  const handleNew = () => {
    setSelectedMenu("New");
    setNeww(true);
    setEdit(false);
    setView(false);
    setSelectedRow(null);
    setViewTable(true);
  };
  const handleView = (menu: string, row: any) => {
    setSelectedMenu(menu);
    setNeww(false);
    setView(true);
    setEdit(false);
    setSelectedRow(row.id);
    setViewTable(true);
  };
  const handleEdit = (menu: string, row: any) => {
    setSelectedMenu(menu);
    setNeww(false);
    setView(false);
    setEdit(true);
    setSelectedRow(row.id);
    setViewTable(true);
  };
  const getStatusColor = (status: string): "success" | "warning" | "error" => {
    const statusMap: Record<string, "success" | "warning" | "error"> = {
      Approved: "success",
      Rejected: "error",
      "Pending at Manager": "warning",
      "Pending at Finance": "warning",
    };
    return statusMap[status] || "primary";
  };
  useEffect(() => {
    void getEmployeeDetails();
    void getExpenseData();
  }, [props.siteUrl, props.context, selectedMenu]);

  const userExpenses = ExpenseData.filter((item: any) => item.Email === userEmail);
  const groupedByRequest = userExpenses.reduce((acc: any, item: any) => {
    const requestNo = item.RequestNo || `no-id-${item.Email}-${item.SubmissionDate}`;
    if (!acc[requestNo]) acc[requestNo] = [];
    acc[requestNo].push(item);
    return acc;
  }, {});

  const requestsRows = Object.entries(groupedByRequest).map(([requestNo, items]: [string, any[]]) => {
    const totalAmount = items.reduce((sum, item) => sum + (item.TotalAmount || 0), 0);
    const firstItem = items[0];
    return {
      showId: `Exp-${requestNo}`,
      id: requestNo,
      name: firstItem.EmployeeName || "Unknown",
      manager: firstItem.Manager,
      amount: totalAmount,
      status: firstItem.Status,
      SubmissionDate: new Date(firstItem.SubmissionDate).toLocaleDateString('en-GB')
    };
  });
  const requestsColumns: GridColDef[] = [
    { field: "showId", headerName: "Request No", flex: 0.7, headerAlign: 'center', align: 'center' },
    { field: "name", headerName: "Name", minWidth: 130, flex: 0.7, headerAlign: 'center', align: 'center' },
    { field: "manager", headerName: "Manager", minWidth: 130, flex: 0.7, headerAlign: 'center', align: 'center' },
    { field: "amount", headerName: "Amount (₹)", minWidth: 130, flex: 0.7, headerAlign: 'center', align: 'center' },
    {
      field: "status", headerName: "Status", minWidth: 150, flex: 0.8, headerAlign: 'center', align: 'center',
      renderCell: (params) => (
        <Button
          variant="outlined"
          color={getStatusColor(params.value)}
          size="small"
          sx={{ fontWeight: '700', borderRadius: 2, textTransform: 'none' }}
        >
          {params.value}
        </Button>
      ),
    },
    { field: "SubmissionDate", headerName: "Submission Date", flex: 0.9, headerAlign: 'center', align: 'center' },
    {
      field: "actions", headerName: "Actions", minWidth: 80, flex: 0.6, headerAlign: "center", align: "center", sortable: false,
      renderCell: (params) => {
        const status = params.row.status;
        const color = getStatusColor(status);
        const editableStatuses = ["Draft", "Revision Requested"];
        const isViewOnly = !editableStatuses.includes(status);
        const label = isViewOnly ? "VIEW" : "EDIT";
        const onClick = () => isViewOnly ? handleView("New", params.row) : handleEdit("New", params.row);
        return (
          <Button variant="contained" color={color} size="small" onClick={onClick} sx={{ borderRadius: 2, textTransform: "none" }}>
            {label}
          </Button>
        );
      },
    }
  ];

  return (
    <Box sx={{ display: "flex", flexDirection: "column", height: "auto", px: 2, mb: 2 }}>
      <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
        <Typography></Typography>
        {!neww && !edit &&
          <Button variant="contained" color="success" sx={{ borderRadius: 2, textTransform: 'none', mr: 5 }} onClick={handleNew}>Raise New Request</Button>
        }
      </Box>
      <Box sx={{ width: "100%" }}>
        {selectedMenu === "New" &&
          <New setSelectedMenu={setSelectedMenu} siteUrl={props.siteUrl} context={props.context} EmployeeDetails={EmployeeDetails} neww={neww} view={view} edit={edit} setNeww={setNeww} setEdit={setEdit} selectedRow={selectedRow} isManager={false} viewTable={viewTable} setViewTable={setViewTable} />
        }
      </Box>

      {selectedMenu !== "New" && (
        <Box sx={{ flexGrow: 1, overflow: "visible", display: "flex", flexDirection: "column", alignItems: "center", mt: 2 }}>
          <Box sx={{ width: "fit-content", minWidth: "95%" }}>
            <DataGrid
              rows={requestsRows}
              columns={requestsColumns}
              paginationModel={paginationModel}
              onPaginationModelChange={setPaginationModel}
              pagination
              pageSizeOptions={[10, 25, 50]}
              disableRowSelectionOnClick
              initialState={{
                pagination: {
                  paginationModel: {
                    pageSize: 10,
                  },
                },
                sorting: {
                  sortModel: [{ field: 'showId', sort: 'desc' }],
                },
              }}
              sx={{
                height: 700,
                "& .MuiDataGrid-columnHeaders": {
                  whiteSpace: "normal",
                  fontWeight: "bold !important",
                  textAlign: "center !important",
                  bgcolor: 'primary.main',
                },
                "& .MuiDataGrid-columnHeaderTitle": {
                  fontSize: "1.1rem !important",
                  fontWeight: "bold",
                },
                "& .MuiDataGrid-cell": {
                  fontWeight: 550,
                },
                border: "3px solid #ddd",
                boxShadow: "0px 4px 10px rgba(0, 0, 0, 0.3)",
                borderRadius: "8px",
              }}
            />
          </Box>
        </Box>
      )}
    </Box>
  );
};

export default Request;