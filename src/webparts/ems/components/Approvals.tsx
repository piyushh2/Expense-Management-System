import * as React from "react";
import { useState, useEffect } from "react";
import { DataGrid, GridColDef, GridPaginationModel } from "@mui/x-data-grid";
import { Button, Box, Typography, IconButton } from "@mui/material";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import ArrowBackOutlinedIcon from '@mui/icons-material/ArrowBackOutlined';
import New from './New'
import Approved from "./Approved";
interface IApprovalsProps {
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
const Approvals: React.FC<IApprovalsProps> = (props) => {
  const userEmail = props.context.pageContext.user.email;
  const [paginationModel, setPaginationModel] = useState<GridPaginationModel>({ page: 0, pageSize: 10, });
  const [EmployeeDetails, setEmployeeDetails] = useState<EmployeeDetails>({});
  const [selectedRow, setSelectedRow] = useState(null);
  const [selectedMenu, setSelectedMenu] = useState("");
  const [ExpenseData, setExpenseData] = useState([]);
  const [isManager, setIsManager] = useState(false);
  const [view, setView] = useState(false);
  const [show, setShow] = useState(true);
  const [viewTable, setViewTable] = useState(false);
  const [showApproved, setShowApproved] = useState(false);
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [approveReject, setApproveReject] = useState(false);

  const getEmployeeDetails = async () => {
    try {
      const listUrl = `${props.siteUrl}/_api/web/lists/getbytitle('Employee')/items`;
      const listResponse: SPHttpClientResponse = await props.context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );
      const res = await listResponse.json();
      const hasMatchingManager = res.value.some((item: any) => item.ManagerEmail === userEmail);
      const filteredEmployee = res.value.filter((item: any) => item.Email === userEmail);
      setEmployeeDetails(filteredEmployee[0]);
      setApproveReject(hasMatchingManager);
    }
    catch (error) {
      console.log(`Error: ${error}`);
    }
  }
  const getExpenseData = async () => {
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
  }
  useEffect(() => {
    void getEmployeeDetails();
    void getExpenseData();
  }, [props.siteUrl, props.context, refreshTrigger, selectedMenu, showApproved]);
  const userExpenses = ExpenseData.filter((item: any) => item.ManagerEmail === userEmail);
  const groupedByRequest = userExpenses.reduce((acc: any, item: any) => {
    const requestNo = item.RequestNo || `no-id-${item.Email}-${item.SubmissionDate}`;
    if (!acc[requestNo]) acc[requestNo] = [];
    acc[requestNo].push(item);
    return acc;
  }, {});
  const approvalsRows = Object.entries(groupedByRequest).map(([requestNo, items]: [string, any[]]) => {
    const totalAmount = items.reduce((sum, item) => sum + (item.TotalAmount || 0), 0);
    const firstItem = items[0];
    return {
      showId: `Exp-${requestNo}`,
      id: requestNo,
      name: firstItem.EmployeeName || "Unknown",
      amount: totalAmount,
      status: firstItem.Status,
      SubmissionDate: new Date(firstItem.SubmissionDate).toLocaleDateString("en-GB")
    };
  }).filter(row => {
    const status = row.status?.toLowerCase();
    return status === "pending at manager" || status === "pending at finance";
  });

  const handleShow = (menu: string, row: any) => {
    setSelectedMenu(menu);
    setView(true);
    setSelectedRow(row.id);
    setIsManager(true);
    setViewTable(true);
    setRefreshTrigger(prev => prev + 1);
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
  const handleOpen = async () => {
    setShowApproved(true);
    setIsManager(true);
    setShow(false);
    setRefreshTrigger(prev => prev + 1);
  }
  const handleClose = async () => {
    setShowApproved(false);
    setIsManager(false);
    setShow(true);
    setRefreshTrigger(prev => prev + 1);
  }
  const approvalsColumns: GridColDef[] = [
    { field: "showId", headerName: "Request No", flex: 0.7, headerAlign: 'center', align: 'center' },
    { field: "name", headerName: "Name", minWidth: 130, flex: 0.7, headerAlign: 'center', align: 'center' },
    { field: "amount", headerName: "Amount (₹)", minWidth: 130, flex: 0.7, headerAlign: 'center', align: 'center' },
    {
      field: "status", headerName: "Status", minWidth: 150, flex: 0.7, headerAlign: 'center', align: 'center',
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
      field: "actions", headerName: "Actions", minWidth: 80, flex: 0.7, headerAlign: "center", align: "center", sortable: false,
      renderCell: (params) => {
        const status = params.row.status;
        const color = getStatusColor(status);
        return (
          <Button variant="contained" size="small" color={color}
            sx={{ borderRadius: 2, textTransform: "none", marginRight: 1 }}
            onClick={() => handleShow("New", params.row)}>
            VIEW
          </Button>
        );
      },
    }
  ];

  return (
    <Box sx={{ display: "flex", flexDirection: "column", height: "auto", px: 2, mb: 2 }}>
      <Box sx={{ display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%" }}>
        <Typography></Typography>
        {show && approveReject &&
          <Button variant="contained" color="success" sx={{ borderRadius: 2, textTransform: 'none', mr: 5 }} onClick={handleOpen}>Approved/Rejected Requests</Button>
        }
        {!show &&
          <IconButton onClick={handleClose} sx={{ color: 'blue', mr: 6 }}><ArrowBackOutlinedIcon /></IconButton>
        }
      </Box>
      {selectedMenu === "New" && <New selectedRow={selectedRow} view={view} setSelectedMenu={setSelectedMenu} siteUrl={props.siteUrl} context={props.context} EmployeeDetails={EmployeeDetails} isManager={isManager} neww={false} edit={false} viewTable={viewTable} setViewTable={undefined} setNeww={undefined} setEdit={undefined} />}
      {showApproved && <Approved context={props.context} siteUrl={props.siteUrl} />}
      {selectedMenu !== "New" && !showApproved &&
        <Box sx={{ display: "flex", flexGrow: 1, flexDirection: "column", alignItems: "center", mt: 2 }}>
          <Box sx={{ width: "fit-content", minWidth: "95%" }}>
            <DataGrid
              rows={approvalsRows}
              columns={approvalsColumns}
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
              }}
              sx={{
                height: 700,
                "& .MuiDataGrid-columnHeaders": {
                  whiteSpace: "normal",
                  fontWeight: "bold !important",
                  textAlign: "center !important",
                  bgcolor: 'primary.light',
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
      }
    </Box>
  );
};

export default Approvals;