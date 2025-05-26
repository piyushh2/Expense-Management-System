import * as React from "react";
import { useState, useEffect, useCallback } from "react";
import { TextField, Button, MenuItem, Select, Paper, Typography, Box, FormControl, InputLabel, SelectChangeEvent, Input, IconButton } from "@mui/material";
import { DataGrid, GridColDef, GridPaginationModel } from "@mui/x-data-grid";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import 'react-toastify/dist/ReactToastify.css';
import AddCircleOutlinedIcon from "@mui/icons-material/AddCircleOutlined";
import CancelIcon from '@mui/icons-material/Cancel';
import RemoveOutlinedIcon from '@mui/icons-material/RemoveOutlined';
import FileUploadOutlinedIcon from '@mui/icons-material/FileUploadOutlined';
import VisibilityIcon from '@mui/icons-material/Visibility';
import DownloadIcon from '@mui/icons-material/Download';
import ExpenseTable from './Table'
import ApprovalHistory from './ApprovalHistory';
import { v4 as uuidv4 } from 'uuid';

const formatDateForDisplay = (isoDate: string): string => {
  if (!isoDate) return "";
  const date = new Date(isoDate);
  return isNaN(date.getTime()) ? "" : date.toLocaleDateString('en-GB');
};
const formatDateForSharePoint = (dateString: string): string => {
  if (!dateString) return '';
  try {
    const date = new Date(dateString);
    return isNaN(date.getTime()) ? dateString : date.toISOString();
  } catch {
    return dateString;
  }
};
const parseDateFromSharePoint = (spDate: string): string => {
  if (!spDate) return "";
  if (spDate.includes('T')) return spDate;
  if (spDate.includes('/')) {
    const parts = spDate.split('/');
    if (parts.length === 3) {
      const date = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
      return isNaN(date.getTime()) ? "" : date.toISOString();
    }
  }
  return new Date().toISOString();
};
const updateDataToSharePoint = async (listName: string, item: any, siteUrl: string, itemId: number, context: any): Promise<void> => {
  try {
    const url = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
    const payload = {
      __metadata: { type: `SP.Data.${listName}ListItem` },
      ...item
    };
    const spHttpClientOptions = {
      headers: {
        'Accept': 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
        'X-HTTP-Method': 'MERGE',
        'If-Match': '*'
      },
      body: JSON.stringify(payload)
    };
    const response: SPHttpClientResponse = await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      spHttpClientOptions
    );
    if (!response.ok) {
      const error = await response.text();
      throw new Error(`Failed to update item in ${listName}: ${error}`);
    }
  } catch (error) {
    console.error(`Error updating item in ${listName}:`, error);
    throw error;
  }
};

interface ExpenseRow {
  CMSID: string;
  Company: string;
  id: any;
  RequestID: any;
  expenseDate: string;
  merchant: string;
  expenseType: string;
  currency: string;
  expenseAmount: string;
  multiplier: string;
  totalAmount: string;
  reason: string;
  file: File | null;
  existingFileUrl: string | null;
  ID: any;
  ExpenseDate: any;
  Merchant: any;
  ExpenseType: any;
  Currency: any;
  ExpenseAmount: any;
  Multiplier: any;
  TotalAmount: any;
  Reason: any;
}
interface Expense {
  EmployeeID: string;
  EmployeeName: string;
  Department: string;
  Country: string;
  SubmissionDate: string;
  Manager: string;
  RequestNo: string;
  Currency: string;
  Company: string;
  CMSID: string;
  Purpose: string;
  Status: string;
}
interface NewProps {
  setSelectedMenu: (menu: string) => void;
  siteUrl: string;
  context: any;
  EmployeeDetails: any;
  neww: boolean;
  view: boolean;
  edit: boolean;
  selectedRow: any;
  isManager: boolean;
  viewTable: boolean;
  setViewTable: any;
  setNeww: any;
  setEdit: any;
}
interface ApprovalHistoryEntry {
  RequestNo: any;
  approvalDate: string;
  approver: string;
  remarks: string;
}

const ExpenseList = "Expenses";
const CurrencyList = "CurrencyMaster";
const ExpenseTypeList = "ExpenseTypeMaster";
const CMSIDList = "CMSRequest";
const RequestList = "Request";
const libraryName = "ExpenseAttachments";
const ApprovalHistoryList = "ApprovalHistory"

const New: React.FC<NewProps> = ({ setSelectedMenu, siteUrl, context, EmployeeDetails, neww, view, edit, setNeww, setEdit, selectedRow, isManager, viewTable, setViewTable }) => {
  const FixedEmployeeDetails = ({ selectedExpense }: { selectedExpense: Expense }) => {
    return (
      <>
        <FormControl sx={{ width: '23%' }}>
          <InputLabel shrink sx={{ color: 'black' }}>Employee ID</InputLabel>
          <Input value={selectedExpense.EmployeeID} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
        </FormControl>
        <FormControl sx={{ width: '23%' }}>
          <InputLabel shrink sx={{ color: 'black' }}>Employee Name</InputLabel>
          <Input value={selectedExpense.EmployeeName} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
        </FormControl>
        <FormControl sx={{ width: '23%' }}>
          <InputLabel shrink sx={{ color: 'black' }}>Department</InputLabel>
          <Input value={selectedExpense.Department} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
        </FormControl>
        <FormControl sx={{ width: '23%' }}>
          <InputLabel shrink sx={{ color: 'black' }}>Country</InputLabel>
          <Input value={selectedExpense.Country} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
        </FormControl>
        <FormControl sx={{ width: '23%' }}>
          <InputLabel shrink sx={{ color: 'black' }}>Submission Date</InputLabel>
          <Input value={selectedExpense.SubmissionDate ? formatDateForDisplay(selectedExpense.SubmissionDate) : ""} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
        </FormControl>
        <FormControl sx={{ width: '23%' }}>
          <InputLabel shrink sx={{ color: 'black' }}>Manager</InputLabel>
          <Input value={selectedExpense.Manager} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
        </FormControl>
        <FormControl sx={{ width: '23%' }}>
          <InputLabel shrink sx={{ color: 'black' }}>Request No</InputLabel>
          <Input value={selectedExpense?.RequestNo ? `Exp-${selectedExpense.RequestNo}` : "Loading..."} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
        </FormControl>
        <FormControl sx={{ width: '23%' }}>
          <InputLabel shrink sx={{ color: 'black' }}>Local Currency</InputLabel>
          <Input value="INR" disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
        </FormControl>
      </>
    );
  };
  const GrandTotal = ({ grandTotal }: { grandTotal: number }) => {
    return (
      <Box display="flex" justifyContent="flex-end" alignItems="center" width="100%" bgcolor="#ebebeb" py={1.5} boxShadow={3}>
        <Typography variant="h6" fontWeight="bold" color="text.primary" sx={{ mr: 3, fontSize: '1.15rem', letterSpacing: '0.6px', }} >
          Grand Total:&nbsp;
          <Box component="span" color="text.primary">{grandTotal}</Box>
        </Typography>
      </Box>
    )
  };

  const [paginationModel, setPaginationModel] = useState<GridPaginationModel>({ page: 0, pageSize: 10, });
  const [approvalHistory, setApprovalHistory] = useState<ApprovalHistoryEntry[]>([]);
  const [ExpenseData, setExpenseData] = useState<ExpenseRow[]>([]);
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [selectedExpense, setSelectedExpense] = useState<any>([]);
  const [expenseType, setExpenseType] = useState<any[]>([]);
  const [currency, setCurrency] = useState<any[]>([]);
  const [grandTotal, setGrandTotal] = useState(0);
  const [CMSID, setCMSID] = useState<any[]>([]);
  const [company, setCompany] = useState("");
  const [purpose, setPurpose] = useState("");
  const [remarks, setRemarks] = useState("");
  const [cmsid, setCmsid] = useState("");
  const [ReqNo, setReqNo] = useState(1);
  const filteredApprovalHistory = approvalHistory.filter((item) => item.RequestNo === String(selectedExpense.RequestNo));

  const [inputData, setInputData] = useState<{
    id: any;
    expenseDate: string;
    merchant: string;
    expenseType: string;
    currency: string;
    expenseAmount: string;
    multiplier: string;
    totalAmount: any;
    reason: string;
    file: any;
    requestId: string;
    existingFileUrl: any,
  }[]>([]);
  const [rowsToDelete, setRowsToDelete] = useState<{ id: number; requestId: string }[]>([]);
  useEffect(() => {
    if (neww) {
      setCompany("");
      setCmsid("");
      setInputData([]);
      setPurpose("");
    }
  }, [neww]);
  const fetchAllData = useCallback(async () => {
    try {
      setExpenseData([]);
      setSelectedExpense([]);
      setInputData([]);
      setCurrency([]);
      setExpenseType([]);
      setCMSID([]);
      setReqNo(1);
      const [reqNoResponse, expenseResponse, currencyResponse, expenseTypeResponse, cmsidResponse, approvalHistoryResponse] = await Promise.all([
        context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('Request')/items?$orderby=Id desc&$top=1`,
          SPHttpClient.configurations.v1
        ),
        context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items`,
          SPHttpClient.configurations.v1
        ),
        fetch(`${siteUrl}/_api/web/lists/getbytitle('${CurrencyList}')/items`, {
          method: "GET",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }),
        fetch(
          `${siteUrl}/_api/web/lists/getbytitle('${ExpenseTypeList}')/items?$select=Title,AttachRequired`,
          {
            method: "GET",
            headers: {
              Accept: "application/json;odata=verbose",
            },
          }
        ),
        fetch(`${siteUrl}/_api/web/lists/getbytitle('${CMSIDList}')/items`, {
          method: "GET",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }),
        fetch(`${siteUrl}/_api/web/lists/getbytitle('${ApprovalHistoryList}')/items`, {
          method: "GET",
          headers: {
            Accept: "application/json;odata=verbose",
          },
        }),
      ]);
      const reqNoData = await reqNoResponse.json();
      setReqNo(reqNoData.value?.[0]?.RequestNo ? reqNoData.value[0].RequestNo + 1 : 1);
      const expenseData = await expenseResponse.json();
      if (expenseData?.value?.length) setExpenseData(expenseData.value);
      const currencyData = await currencyResponse.json();
      setCurrency(currencyData.d.results);
      const expenseTypeData = await expenseTypeResponse.json();
      setExpenseType(expenseTypeData.d.results);
      const cmsidData = await cmsidResponse.json();
      setCMSID(cmsidData.d.results);
      const approvalHistoryData = await approvalHistoryResponse.json();
      setApprovalHistory(approvalHistoryData.d.results.map((item: any) => ({
        RequestNo: item.RequestNo || '',
        approvalDate: item.ApprovalDate || '',
        approver: item.Approver || '',
        remarks: item.Remarks || '',
      })));
    } catch (error) {
      console.error("Error fetching data:", error);
    }
  }, [siteUrl, context]);
  const getAttachments = async (requestId: string) => {
    try {
      const response = await context.spHttpClient.get(
        `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryName}')/Files?$filter=startswith(Name, '${requestId}_')`,
        SPHttpClient.configurations.v1
      );
      const data = await response.json();
      return data.value || [];
    } catch (error) {
      console.error(`Error fetching attachments for RequestID ${requestId}:`, error);
      return [];
    }
  };
  const addRow = () => {
    const uuid = uuidv4();
    const newRow = {
      id: uuid,
      expenseDate: "",
      merchant: "",
      expenseType: "",
      currency: "",
      expenseAmount: "",
      multiplier: "",
      totalAmount: "",
      reason: "",
      file: null,
      requestId: uuid,
      existingFileUrl: null,
    };
    setInputData((prevRows) => [...prevRows, newRow]);
  };
  const removeRow = (id: any): void => {
    try {
      const rowToDelete = inputData.find((row) => row.id === id);
      if (!rowToDelete) {
        console.warn(`Row with id ${id} not found in inputData`);
        return;
      }
      const updatedRows = inputData.filter((row) => row.id !== id);
      setInputData(updatedRows);
      if (typeof id === 'number') setRowsToDelete((prev) => [...prev, { id, requestId: rowToDelete.requestId }]);
    } catch (error) {
      console.error(`Error marking row ${id} for deletion:`, error);
      alert('An error occurred while marking the row for deletion. Please check the console for details.');
    }
  };
  const handleCellEdit = (id: number, field: keyof ExpenseRow, value: string | number) => {
    const updatedData = inputData.map((row) => {
      if (row.id === id) {
        const updatedRow = { ...row, [field]: value };
        if (field === 'expenseAmount' || field === 'multiplier') {
          const amount = parseFloat(updatedRow.expenseAmount) || 0;
          const rawMultiplier = parseFloat(updatedRow.multiplier);
          const multiplier = isNaN(rawMultiplier) ? 1 : rawMultiplier;
          updatedRow.totalAmount = amount * multiplier;
        }
        return updatedRow;
      }
      return row;
    });
    setInputData(updatedData);
  };
  const replaceFileForRequest = async (row: any) => {
    if (!row.file || !row.requestId) {
      throw new Error('File or RequestID is missing for upload');
    }
    const requestId = row.requestId;
    try {
      const existingFiles = await getAttachments(requestId);
      for (const file of existingFiles) {
        const deleteUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${file.ServerRelativeUrl}')`;
        const deleteResponse = await context.spHttpClient.post(deleteUrl, SPHttpClient.configurations.v1, {
          headers: {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*',
          },
        });
        if (!deleteResponse.ok) {
          console.warn(`Failed to delete existing file: ${file.Name}`);
        }
      }
      const fileName = `${requestId}_${row.file.name}`;
      const fileArrayBuffer = await row.file.arrayBuffer();
      const uploadUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryName}')/Files/add(url='${fileName}', overwrite=true)`;
      const uploadResponse = await context.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
        body: fileArrayBuffer,
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-Type': row.file.type,
          'odata-version': '',
        },
      });
      if (!uploadResponse.ok) {
        throw new Error(`File upload failed for ${fileName}`);
      }
      const uploadResult = await uploadResponse.json();
      const fileItemUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${uploadResult.d.ServerRelativeUrl}')/ListItemAllFields`;
      const listItemResponse = await context.spHttpClient.get(fileItemUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'odata-version': '',
        },
      });
      if (listItemResponse.ok) {
        const listItem = await listItemResponse.json();
        const updateItemUrl = `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${listItem.d.ID})`;
        const updateResponse = await context.spHttpClient.post(updateItemUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose',
            'odata-version': '',
            'X-HTTP-Method': 'MERGE',
            'If-Match': listItem.d.__metadata.etag,
          },
          body: JSON.stringify({
            __metadata: { type: listItem.d.__metadata.type },
            RequestID: requestId,
          }),
        });
        if (!updateResponse.ok) console.warn(`Failed to update metadata for ${fileName}`);
      }
    } catch (error) {
      console.error(`Error uploading file for RequestID ${requestId}:`, error);
      throw error;
    }
  };
  const handleUpdateData = async () => {
    try {
      if (inputData.length === 0) {
        alert('No expense data to update.');
        return;
      }
      for (const row of inputData) {
        if (!row.expenseDate || !row.expenseType || !row.expenseAmount || Number(row.expenseAmount) < 0 || !row.reason || !row.totalAmount) {
          alert("Please complete all required fields before submitting.");
          return;
        }
      }
      for (const { id, requestId } of rowsToDelete) {
        const attachments = await getAttachments(requestId);
        for (const file of attachments) {
          const deleteFileUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${file.ServerRelativeUrl}')`;
          const deleteFileResponse = await context.spHttpClient.post(deleteFileUrl, SPHttpClient.configurations.v1, {
            headers: {
              'X-HTTP-Method': 'DELETE',
              'IF-MATCH': '*',
            },
          });
          if (!deleteFileResponse.ok) {
            console.warn(`Failed to delete file: ${file.Name}`);
          }
        }
        const deleteItemUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items(${id})`;
        const deleteItemResponse = await context.spHttpClient.post(deleteItemUrl, SPHttpClient.configurations.v1, {
          headers: {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*',
          },
        });
        if (!deleteItemResponse.ok) {
          const error = await deleteItemResponse.text();
          throw new Error(`Failed to delete expense item ${id}: ${error}`);
        }
      }
      setRowsToDelete([]);
      const updatePayload = inputData
        .filter((row) => row.id && row.expenseDate)
        .map((row) => {
          const formattedDate = formatDateForSharePoint(row.expenseDate);
          if (!formattedDate) {
            console.warn(`Invalid date for row ID ${row.id}: ${row.expenseDate}`);
            return null;
          }
          return {
            ID: typeof row.id === 'number' ? row.id : undefined,
            ExpenseDate: formattedDate,
            Merchant: row.merchant,
            ExpenseType: row.expenseType,
            Currency: row.currency,
            ExpenseAmount: row.expenseAmount,
            Multiplier: row.multiplier,
            TotalAmount: row.totalAmount,
            Reason: row.reason,
            Purpose: selectedExpense.Purpose,
            Status: "Pending at Manager",
            Company: company,
            CMSID: cmsid,
            RequestID: row.requestId,
          };
        }).filter((item): item is NonNullable<typeof item> => item !== null);
      if (updatePayload.length === 0) {
        alert('No valid data to update');
        return;
      }
      updatePayload.forEach(async (item, index) => {
        const row = inputData[index];
        if (!item.ID || typeof row.id === 'string') {
          const itemData = {
            __metadata: { type: "SP.Data.ExpensesListItem" },
            RequestID: row.requestId,
            Title: EmployeeDetails?.EmployeeName || "Unknown",
            EmployeeID: EmployeeDetails?.EmployeeID || "",
            EmployeeName: EmployeeDetails?.EmployeeName || "",
            Department: EmployeeDetails?.Department || "",
            Country: EmployeeDetails?.Country || "",
            Manager: EmployeeDetails?.Manager || "",
            ManagerEmail: EmployeeDetails?.ManagerEmail || "",
            RequestNo: selectedExpense.RequestNo || "",
            Company: company || "",
            CMSID: cmsid || "",
            ExpenseDate: item.ExpenseDate,
            Merchant: item.Merchant,
            ExpenseType: item.ExpenseType,
            Currency: item.Currency,
            ExpenseAmount: item.ExpenseAmount,
            Multiplier: item.Multiplier,
            TotalAmount: item.TotalAmount,
            Reason: item.Reason,
            SubmissionDate: new Date().toISOString(),
            Status: item.Status,
            Email: context.pageContext.user.email || "",
            Purpose: item.Purpose,
          };
          const url = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items`;
          const response = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: JSON.stringify(itemData)
          });
          if (!response.ok) {
            const error = await response.text();
            throw new Error(`Failed to create new expense item: ${error}`);
          }
        } else {
          await updateDataToSharePoint(ExpenseList, item, siteUrl, item.ID, context);
        }
      })
      for (const row of inputData) {
        if (row.file) {
          try {
            await replaceFileForRequest(row);
          }
          catch (err) {
            console.warn(`File update failed for row ${row.id}:`, err);
          }
        }
      }
      alert(`Request updated successfully. Request ID: Exp-${ReqNo - 1}`);
      setRefreshTrigger(prev => prev + 1);
      setSelectedMenu('Request');
      setNeww(false);
      setEdit(false);
    } catch (error) {
      console.error('Update error:', error);
    }
  };
  const handleDelete = async () => {
    const confirmDelete = window.confirm("Are you sure you want to delete this request?");
    if (!confirmDelete) return;
    try {
      if (!selectedExpense.RequestNo) {
        alert('No request selected for deletion.');
        return;
      }
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items?$filter=RequestNo eq '${selectedExpense.RequestNo}'`;
      const response = await context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
      const expenseItems = await response.json();
      if (!expenseItems.value || expenseItems.value.length === 0) {
        alert('No expense items found for this request.');
        return;
      }
      for (const item of expenseItems.value) {
        const uuid = item.RequestID;
        const files = await getAttachments(uuid);
        for (const file of files) {
          const deleteFileUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${file.ServerRelativeUrl}')`;
          const deleteFileResponse = await context.spHttpClient.post(deleteFileUrl, SPHttpClient.configurations.v1, {
            headers: {
              'X-HTTP-Method': 'DELETE',
              'IF-MATCH': '*',
            },
          });
          if (!deleteFileResponse.ok) {
            console.warn(`Failed to delete file: ${file.Name}`);
          }
        }
        const deleteItemUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items(${item.ID})`;
        const deleteItemResponse = await context.spHttpClient.post(deleteItemUrl, SPHttpClient.configurations.v1, {
          headers: {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*',
          },
        });
        if (!deleteItemResponse.ok) {
          const error = await deleteItemResponse.text();
          throw new Error(`Failed to delete expense item ${item.ID}: ${error}`);
        }
      }
      const requestUrl = `${siteUrl}/_api/web/lists/getbytitle('${RequestList}')/items?$filter=RequestNo eq '${selectedExpense.RequestNo}'`;
      const requestResponse = await context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
      const requestItems = await requestResponse.json();
      if (requestItems.value && requestItems.value.length > 0) {
        for (const request of requestItems.value) {
          const deleteRequestUrl = `${siteUrl}/_api/web/lists/getbytitle('${RequestList}')/items(${request.ID})`;
          const deleteRequestResponse = await context.spHttpClient.post(deleteRequestUrl, SPHttpClient.configurations.v1, {
            headers: {
              'X-HTTP-Method': 'DELETE',
              'IF-MATCH': '*',
            },
          });
          if (!deleteRequestResponse.ok) {
            const error = await deleteRequestResponse.text();
            throw new Error(`Failed to delete request item ${request.ID}: ${error}`);
          }
        }
      }
      alert("Request deleted successfully.");
      setRefreshTrigger(prev => prev + 1);
      setInputData([]);
      setSelectedExpense([]);
      setViewTable(false);
      setSelectedMenu('Request');
      setNeww(false);
      setEdit(false);
    } catch (error) {
      console.error('Delete error:', error);
    }
  };
  const handleExit = async () => {
    setRefreshTrigger(prev => prev + 1);
    setSelectedMenu("Request");
    setNeww(false);
    setEdit(false);
  }
  const handleDraft = async () => {
    try {
      if (company === "" || cmsid === "" || inputData.length === 0) {
        alert("Please provide all required data.");
        return;
      }
      for (const row of inputData) {
        if (!row.expenseDate) {
          alert("Expense Date is required for all rows.");
          return;
        }
        if (!row.expenseAmount || Number(row.expenseAmount) < 0 || !row.totalAmount) {
          alert("Valid Expense Amount is required for all rows.");
          return;
        }
      }
      let requestNo = ReqNo;
      let isExistingDraft = false;
      if (selectedExpense?.RequestNo && selectedExpense?.Status === "Draft") {
        const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items?$filter=RequestNo eq '${selectedExpense.RequestNo}' and Status eq 'Draft'`;
        const response = await context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
        const data = await response.json();
        if (data.value && data.value.length > 0) {
          isExistingDraft = true;
          requestNo = selectedExpense.RequestNo;
        }
      }
      if (!isExistingDraft) {
        const requestData = {
          __metadata: { type: "SP.Data.RequestListItem" },
          Title: EmployeeDetails?.EmployeeName || "Unknown",
          EmployeeID: EmployeeDetails?.EmployeeID,
          RequestNo: requestNo,
        };
        const requestUrl = `${siteUrl}/_api/web/lists/getbytitle('${RequestList}')/items`;
        const requestOptions = {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: JSON.stringify(requestData)
        };
        const requestResponse: SPHttpClientResponse = await context.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, requestOptions);
        if (!requestResponse.ok) {
          const error = await requestResponse.text();
          console.error("Failed to submit request item:", error);
          alert("Error submitting request. Please try again.");
          return;
        }
      }
      for (const { id, requestId } of rowsToDelete) {
        const attachments = await getAttachments(requestId);
        for (const file of attachments) {
          const deleteFileUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${file.ServerRelativeUrl}')`;
          const deleteFileResponse = await context.spHttpClient.post(deleteFileUrl, SPHttpClient.configurations.v1, {
            headers: {
              'X-HTTP-Method': 'DELETE',
              'IF-MATCH': '*',
            },
          });
          if (!deleteFileResponse.ok) console.warn(`Failed to delete file: ${file.Name}`);
        }
        const deleteItemUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items(${id})`;
        const deleteItemResponse = await context.spHttpClient.post(deleteItemUrl, SPHttpClient.configurations.v1, {
          headers: {
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*',
          },
        });
        if (!deleteItemResponse.ok) {
          const error = await deleteItemResponse.text();
          throw new Error(`Failed to delete expense item ${id}: ${error}`);
        }
      }
      setRowsToDelete([]);
      for (const row of inputData) {
        const uuid = row.id && typeof row.id === 'number' ? row.requestId : uuidv4();
        const itemData = {
          __metadata: { type: "SP.Data.ExpensesListItem" },
          RequestID: uuid,
          Title: EmployeeDetails?.EmployeeName || "Unknown",
          EmployeeID: EmployeeDetails?.EmployeeID || "",
          EmployeeName: EmployeeDetails?.EmployeeName || "",
          Department: EmployeeDetails?.Department || "",
          Country: EmployeeDetails?.Country || "",
          Manager: EmployeeDetails?.Manager || "",
          ManagerEmail: EmployeeDetails?.ManagerEmail || "",
          RequestNo: requestNo || "",
          Company: company || "",
          CMSID: cmsid || "",
          ExpenseDate: formatDateForSharePoint(row.expenseDate) || "",
          Merchant: row.merchant || "",
          ExpenseType: row.expenseType || "",
          Currency: row.currency || "",
          ExpenseAmount: row.expenseAmount || null,
          Multiplier: row.multiplier || null,
          TotalAmount: row.totalAmount || null,
          Reason: row.reason || "",
          SubmissionDate: new Date().toISOString() || "",
          Status: "Draft",
          Email: context.pageContext.user.email || "",
          Purpose: purpose || "",
        };
        if (row.id && typeof row.id === 'number' && isExistingDraft) await updateDataToSharePoint(ExpenseList, itemData, siteUrl, row.id, context);
        else {
          const url = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items`;
          const spHttpClientOptions = {
            headers: {
              'Accept': 'application/json;odata=verbose',
              'Content-type': 'application/json;odata=verbose',
              'odata-version': ''
            },
            body: JSON.stringify(itemData)
          };
          const response: SPHttpClientResponse = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions);
          if (!response.ok) {
            const error = await response.text();
            console.error("Failed to submit expense item:", error);
            alert("Some data may not have been submitted. Please try again.");
            return;
          }
        }
        if (row.file) {
          try {
            await replaceFileForRequest({ ...row, requestId: uuid });
          } catch (uploadError) {
            console.error("File upload error:", uploadError);
          }
        }
      }
      alert(`Draft saved successfully. Draft ID: Exp-${ReqNo}`);
      setRefreshTrigger(prev => prev + 1);
      setSelectedMenu("Request");
      setNeww(false);
      setEdit(false);
    } catch (error) {
      console.error("Submission error:", error);
    }
  };
  const handleSubmit = async () => {
    try {
      if (!company || !cmsid || inputData.length === 0) {
        alert("Please provide all required data.");
        return;
      }
      for (const row of inputData) {
        if (!row.expenseDate || !row.expenseType || !row.expenseAmount || Number(row.expenseAmount) < 0 || !row.reason || !row.totalAmount) {
          alert("Please complete all required fields before submitting.");
          return;
        }
      }
      const requestData = {
        __metadata: { type: "SP.Data.RequestListItem" },
        Title: EmployeeDetails?.EmployeeName || "Unknown",
        EmployeeID: EmployeeDetails?.EmployeeID,
        RequestNo: ReqNo,
      };
      const requestUrl = `${siteUrl}/_api/web/lists/getbytitle('${RequestList}')/items`;
      const requestResponse = await context.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=verbose',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: JSON.stringify(requestData)
      });
      if (!requestResponse.ok) {
        const error = await requestResponse.text();
        console.error("Failed to submit request item:", error);
        throw new Error("Error submitting request");
      }
      for (const row of inputData) {
        const uuid = uuidv4();
        const fileName = row.file ? `${uuid}_${row.file.name}` : null;
        const itemData = {
          __metadata: { type: "SP.Data.ExpensesListItem" },
          RequestID: uuid,
          Title: EmployeeDetails?.EmployeeName || "Unknown",
          EmployeeID: EmployeeDetails?.EmployeeID || "",
          EmployeeName: EmployeeDetails?.EmployeeName || "",
          Department: EmployeeDetails?.Department || "",
          Country: EmployeeDetails?.Country || "",
          Manager: EmployeeDetails?.Manager || "",
          ManagerEmail: EmployeeDetails?.ManagerEmail || "",
          RequestNo: ReqNo || "",
          Company: company || "",
          CMSID: cmsid || "",
          ExpenseDate: formatDateForSharePoint(row.expenseDate) || "",
          Merchant: row.merchant || "",
          ExpenseType: row.expenseType || "",
          Currency: row.currency || "",
          ExpenseAmount: parseFloat(row.expenseAmount) || 0,
          Multiplier: parseFloat(row.multiplier) || 1,
          TotalAmount: parseFloat(row.totalAmount) || 0,
          Reason: row.reason || "",
          SubmissionDate: new Date().toISOString(),
          Status: "Pending at Manager",
          Email: context.pageContext.user.email || "",
          Purpose: purpose || "",
        };
        const expenseUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items`;
        const expenseResponse = await context.spHttpClient.post(expenseUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata=verbose',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': ''
          },
          body: JSON.stringify(itemData)
        });
        if (!expenseResponse.ok) {
          const error = await expenseResponse.text();
          console.error("Failed to submit expense item:", error);
          throw new Error("Failed to submit expense item");
        }
        if (row.file) {
          try {
            const fileArrayBuffer = await row.file.arrayBuffer();
            const fileUrl = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${libraryName}')/Files/add(url='${fileName}', overwrite=true)`;
            const uploadResponse = await context.spHttpClient.post(fileUrl, SPHttpClient.configurations.v1, {
              body: fileArrayBuffer,
              headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-Type': row.file.type,
                'odata-version': ''
              }
            });
            if (!uploadResponse.ok) throw new Error(`File upload failed with status ${uploadResponse.status}`);
            const uploadResult = await uploadResponse.json();
            const fileItemUrl = `${siteUrl}/_api/web/GetFileByServerRelativeUrl('${uploadResult.d.ServerRelativeUrl}')/ListItemAllFields`;
            const listItemResponse = await context.spHttpClient.get(fileItemUrl, SPHttpClient.configurations.v1, {
              headers: {
                'Accept': 'application/json;odata=verbose',
                'odata-version': ''
              }
            });
            if (listItemResponse.ok) {
              const listItem = await listItemResponse.json();
              const updateItemUrl = `${siteUrl}/_api/web/lists/getbytitle('${libraryName}')/items(${listItem.d.ID})`;
              const updateResponse = await context.spHttpClient.post(updateItemUrl, SPHttpClient.configurations.v1, {
                headers: {
                  'Accept': 'application/json;odata=verbose',
                  'Content-Type': 'application/json;odata=verbose',
                  'odata-version': '',
                  'X-HTTP-Method': 'MERGE',
                  'If-Match': listItem.d.__metadata.etag
                },
                body: JSON.stringify({
                  __metadata: { type: listItem.d.__metadata.type },
                  RequestID: uuid
                })
              });
              if (!updateResponse.ok) console.warn("Failed to update file metadata, but file was uploaded");
            }
          } catch (uploadError) {
            console.error("File upload error:", uploadError)
          }
        }
      }
      alert(`Request submitted successfully. Request ID: Exp-${ReqNo}`);
      setRefreshTrigger(prev => prev + 1);
      setSelectedMenu("Request");
      setNeww(false);
      setEdit(false);
    } catch (error) {
      console.error("Submission error:", error);
    }
  };
  const handleReject = async () => {
    try {
      if (!remarks) {
        alert("Remarks is required");
        return;
      }
      const updatedStatus = "Rejected";
      const requestNumber = selectedExpense.RequestNo;
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items?$filter=RequestNo eq '${requestNumber}'`;
      const response = await context.spHttpClient.get(
        listUrl,
        SPHttpClient.configurations.v1
      );
      if (!response.ok) {
        throw new Error("Failed to fetch expense items");
      }
      const data = await response.json();
      for (const item of data.value) {
        await updateDataToSharePoint(ExpenseList,
          {
            ID: item.ID,
            Status: updatedStatus,
          },
          siteUrl,
          item.ID,
          context
        );
      }
      const itemData = {
        __metadata: { type: "SP.Data.ApprovalHistoryListItem" },
        RequestNo: String(selectedExpense.RequestNo),
        ApprovalDate: new Date().toISOString(),
        Approver: context.pageContext.user.displayName,
        Remarks: remarks || "",
      };
      const expenseUrl: string = `${siteUrl}/_api/web/lists/getbytitle('${ApprovalHistoryList}')/items`;
      const expenseResponse: Response = await context.spHttpClient.post(expenseUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-type": "application/json;odata=verbose",
            "odata-version": "",
          },
          body: JSON.stringify(itemData),
        }
      );
      if (!expenseResponse.ok) {
        const error = await expenseResponse.text();
        console.error("Failed to submit approval history item:", error);
        throw new Error("Failed to submit approval history item");
      }
      setSelectedExpense((prev: any) => ({ ...prev, Status: updatedStatus }));
      alert("Request rejected successfully.");
      setSelectedMenu("Request");
      setNeww(false);
      setEdit(false);
    } catch (error) {
      console.error("Rejection Error:", error);
    }
  };
  const handleRevision = async () => {
    try {
      if (!remarks) {
        alert("Remarks is required");
        return;
      }
      const employeeUrl = `${siteUrl}/_api/web/lists/getbytitle('Employee')/items`;
      const employeeResponse = await context.spHttpClient.get(employeeUrl, SPHttpClient.configurations.v1);
      const employeeData = await employeeResponse.json();
      const currentManagerInfo = employeeData.value.find((item: any) => item.EmployeeName === selectedExpense.EmployeeName);
      const newManager = currentManagerInfo?.Manager || "";
      const newEmail = currentManagerInfo?.ManagerEmail || "";
      const updatedStatus = "Revision Requested";
      const requestNumber = selectedExpense.RequestNo;
      const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items?$filter=RequestNo eq '${requestNumber}'`;
      const response = await context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw new Error("Failed to fetch expense items");
      }
      const data = await response.json();
      for (const item of data.value) {
        await updateDataToSharePoint(ExpenseList, {
          ID: item.ID,
          Status: updatedStatus,
          Manager: newManager,
          ManagerEmail: newEmail,
        }, siteUrl, item.ID, context);
      }
      const itemData = {
        __metadata: { type: "SP.Data.ApprovalHistoryListItem" },
        RequestNo: String(selectedExpense.RequestNo),
        ApprovalDate: new Date().toISOString(),
        Approver: context.pageContext.user.displayName,
        Remarks: remarks || "",
      };
      const expenseUrl: string = `${siteUrl}/_api/web/lists/getbytitle('${ApprovalHistoryList}')/items`;
      const expenseResponse: Response = await context.spHttpClient.post(expenseUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-type": "application/json;odata=verbose",
            "odata-version": "",
          },
          body: JSON.stringify(itemData),
        }
      );
      if (!expenseResponse.ok) {
        const error = await expenseResponse.text();
        console.error("Failed to submit approval history item:", error);
        throw new Error("Failed to submit approval history item");
      }
      setSelectedExpense((prev: any) => ({ ...prev, Status: updatedStatus }));
      alert("Revision requested successfully.");
      setRefreshTrigger(prev => prev + 1);
      setSelectedMenu("Request");
      setNeww(false);
      setEdit(false);
    } catch (error) {
      console.log("Rejection Error:", error);
    }
  }
  const handleApprove = async () => {
    try {
      if (!remarks) {
        alert("Remarks is required");
        return;
      }
      const employeeUrl = `${siteUrl}/_api/web/lists/getbytitle('Employee')/items`;
      const employeeResponse = await context.spHttpClient.get(employeeUrl, SPHttpClient.configurations.v1);
      const employeeData = await employeeResponse.json();
      const ans = (employeeData.value.find((item: any) => item.Email === selectedExpense.Email)).HigherAuthority || "False";
      const currentManagerInfo = employeeData.value.find((item: any) => item.EmployeeName === selectedExpense.Manager);
      const newManager = currentManagerInfo?.Manager || "";
      const newEmail = currentManagerInfo?.ManagerEmail || "";
      const status = selectedExpense.Status?.trim().toLowerCase();
      const requestNumber = selectedExpense.RequestNo;
      if (ans === "True") {
        if (status === "pending at manager") {
          const updatedStatus = "Approved";
          await updateDataToSharePoint(ExpenseList, { ID: selectedExpense.ID, Status: updatedStatus, }, siteUrl, selectedExpense.ID, context);
          const itemData = {
            __metadata: { type: "SP.Data.ApprovalHistoryListItem" },
            RequestNo: String(selectedExpense.RequestNo),
            ApprovalDate: new Date().toISOString(),
            Approver: context.pageContext.user.displayName,
            Remarks: remarks || "",
          };
          const expenseUrl: string = `${siteUrl}/_api/web/lists/getbytitle('${ApprovalHistoryList}')/items`;
          const expenseResponse: Response = await context.spHttpClient.post(expenseUrl, SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
              },
              body: JSON.stringify(itemData),
            }
          );
          if (!expenseResponse.ok) {
            const error = await expenseResponse.text();
            console.error("Failed to submit approval history item:", error);
            throw new Error("Failed to submit approval history item");
          }
          setSelectedExpense((prev: any) => ({ ...prev, Status: updatedStatus }));
          alert("Request approved successfully.");
          setRefreshTrigger(prev => prev + 1);
          setSelectedMenu("Request");
          setNeww(false);
          setEdit(false);
        }
      } else {
        if (status === "pending at manager") {
          const updatedStatus = "Pending at Finance";
          const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items?$filter=RequestNo eq '${requestNumber}'`;
          const response = await context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
          const data = await response.json();
          for (const item of data.value) await updateDataToSharePoint(ExpenseList, { ID: item.ID, Status: updatedStatus, Manager: newManager, ManagerEmail: newEmail, }, siteUrl, item.ID, context);
          const itemData = {
            __metadata: { type: "SP.Data.ApprovalHistoryListItem" },
            RequestNo: String(selectedExpense.RequestNo),
            ApprovalDate: new Date().toISOString(),
            Approver: context.pageContext.user.displayName,
            Remarks: remarks || "",
          };
          const expenseUrl: string = `${siteUrl}/_api/web/lists/getbytitle('${ApprovalHistoryList}')/items`;
          const expenseResponse: Response = await context.spHttpClient.post(expenseUrl, SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
              },
              body: JSON.stringify(itemData),
            }
          );
          if (!expenseResponse.ok) {
            const error = await expenseResponse.text();
            console.error("Failed to submit approval history item:", error);
            throw new Error("Failed to submit approval history item");
          }
          setSelectedExpense((prev: any) => ({ ...prev, Status: updatedStatus }));
          alert("Request sent to Finance for approval.");
          setRefreshTrigger(prev => prev + 1);
          setSelectedMenu("Request");
          setNeww(false);
          setEdit(false);
        }
        if (status === "pending at finance") {
          const updatedStatus = "Approved";
          const listUrl = `${siteUrl}/_api/web/lists/getbytitle('${ExpenseList}')/items?$filter=RequestNo eq '${requestNumber}'`;
          const response = await context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
          const data = await response.json();
          for (const item of data.value) await updateDataToSharePoint(ExpenseList, { ID: item.ID, Status: updatedStatus, }, siteUrl, item.ID, context);
          const itemData = {
            __metadata: { type: "SP.Data.ApprovalHistoryListItem" },
            RequestNo: String(selectedExpense.RequestNo),
            ApprovalDate: new Date().toISOString(),
            Approver: context.pageContext.user.displayName,
            Remarks: remarks || "",
          };
          const expenseUrl: string = `${siteUrl}/_api/web/lists/getbytitle('${ApprovalHistoryList}')/items`;
          const expenseResponse: Response = await context.spHttpClient.post(expenseUrl, SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=verbose",
                "Content-type": "application/json;odata=verbose",
                "odata-version": "",
              },
              body: JSON.stringify(itemData),
            }
          );
          if (!expenseResponse.ok) {
            const error = await expenseResponse.text();
            console.error("Failed to submit approval history item:", error);
            throw new Error("Failed to submit approval history item");
          }
          setSelectedExpense((prev: any) => ({ ...prev, Status: updatedStatus }));
          alert("Request approved successfully.");
          setRefreshTrigger(prev => prev + 1);
          setSelectedMenu("Request");
          setNeww(false);
          setEdit(false);
        }
      }
    } catch (error) {
      console.log("Approval Error:", error);
    }
  };
  const columns: GridColDef[] = [
    {
      field: "expenseDate", flex: 0.8,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Expense Date<span style={{ color: "red", fontSize: '1.3rem' }}> *</span></span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        const inputDateValue = currentRow?.expenseDate ? new Date(currentRow.expenseDate).toISOString().split('T')[0] : "";
        const displayValue = currentRow?.expenseDate ? new Date(currentRow.expenseDate).toLocaleDateString('en-GB') : "";
        return (
          <TextField
            type={neww || edit ? 'date' : 'text'}
            fullWidth
            value={neww || edit ? inputDateValue : displayValue}
            onChange={(e) => {
              const newValue = e.target.value;
              const isoDate = newValue ? new Date(newValue).toISOString() : "";
              setInputData((prev) => prev.map((row) => row.id === rowId ? { ...row, expenseDate: isoDate } : row));
              handleCellEdit(rowId as number, "expenseDate", isoDate);
            }}
          />
        );
      }
    },
    {
      field: "merchant", flex: 0.6,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Vendor</span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        return (
          <TextField fullWidth value={currentRow?.merchant || ""}
            onChange={(e) => {
              const newValue = e.target.value;
              setInputData((prev) => prev.map((row) => row.id === rowId ? { ...row, merchant: newValue } : row));
              handleCellEdit(rowId as number, "merchant", newValue);
            }}
            inputProps={{
              style: { whiteSpace: 'pre' },
            }}
            onKeyDown={(e) => {
              if (e.key === ' ') e.stopPropagation();
            }}
          />
        );
      }
    },
    {
      field: "expenseType", flex: 0.7,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Expense Type<span style={{ color: "red", fontSize: '1.3rem' }}> *</span></span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        return (
          <Select fullWidth value={currentRow?.expenseType || ""}
            onChange={(e: SelectChangeEvent<string>) => {
              const newValue = e.target.value;
              setInputData((prev) => prev.map((row) => row.id === rowId ? { ...row, expenseType: newValue } : row));
              handleCellEdit(rowId as number, "expenseType", newValue);
            }}
          >
            {expenseType.map((item, index) => (
              <MenuItem key={index} value={item.Title}>
                {item.Title}
              </MenuItem>
            ))}
          </Select>
        );
      },
    },
    {
      field: "currency", flex: 0.6,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Currency</span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        return (
          <Select fullWidth value={currentRow?.currency || ""}
            onChange={(e: SelectChangeEvent<string>) => {
              const newValue = e.target.value;
              setInputData((prev) => prev.map((row) => row.id === rowId ? { ...row, currency: newValue } : row));
              handleCellEdit(rowId as number, "currency", newValue);
            }}
          >
            {currency.map((item, index) => (
              <MenuItem key={index} value={item.Currency}>
                {item.Currency}
              </MenuItem>
            ))}
          </Select>
        );
      },
    },
    {
      field: "expenseAmount", flex: 0.5,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Amount<span style={{ color: "red", fontSize: '1.3rem' }}> *</span></span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        return (
          <TextField type="number" fullWidth value={currentRow?.expenseAmount || ""}
            onChange={(e) => {
              const newValue = e.target.value;
              setInputData((prev) => prev.map((row) => row.id === rowId ? { ...row, expenseAmount: newValue } : row));
              handleCellEdit(rowId as number, "expenseAmount", newValue);
            }}
          />
        );
      },
    },
    {
      field: "multiplier", flex: 0.5,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Multiplier</span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        return (
          <TextField type="number" fullWidth value={currentRow?.multiplier || ""}
            onChange={(e) => {
              const newValue = e.target.value;
              setInputData((prev) => prev.map((row) => row.id === rowId ? { ...row, multiplier: newValue } : row));
              handleCellEdit(rowId as number, "multiplier", newValue);
            }}
          />
        );
      },
    },
    {
      field: "totalAmount", flex: 0.6,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Total Amount</span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        return (
          <TextField fullWidth disabled value={currentRow?.totalAmount || ""} />
        );
      },
    },
    {
      field: "reason", flex: 0.6,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Reason<span style={{ color: "red", fontSize: '1.3rem' }}> *</span></span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        return (
          <TextField fullWidth multiline rows={1} value={currentRow?.reason || ""}
            onChange={(e) => {
              const newValue = e.target.value;
              setInputData((prev) => prev.map((row) => row.id === rowId ? { ...row, reason: newValue } : row));
              handleCellEdit(rowId as number, "reason", newValue);
            }}
            inputProps={{
              style: { whiteSpace: 'pre' },
            }}
            onKeyDown={(e) => {
              if (e.key === ' ') e.stopPropagation();
            }}
          />
        );
      },
    },
    {
      field: "attachments",
      flex: 0.6,
      cellClassName: "interactive-cell",
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}>Attachments</span>
      ),
      renderCell: (params) => {
        const rowId = params.id;
        const currentRow = inputData.find((row) => row.id === rowId);
        const [attachments, setAttachments] = useState<any[]>([]);
        const requestId = currentRow?.requestId;
        useEffect(() => {
          const fetchAttachments = async () => {
            if (requestId) {
              const files = await getAttachments(requestId);
              setAttachments(files);
            }
          };
          void fetchAttachments();
        }, [requestId]);
        const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
          const file = event.target.files?.[0];
          if (file) setInputData((prev) => prev.map((row) => row.id === rowId ? { ...row, file, existingFileUrl: null } : row));
        };
        const handleDownload = async () => {
          if (attachments.length > 0) {
            const fileUrl = `${siteUrl}/${attachments[0].ServerRelativeUrl.slice(19)}`;
            const fileName = attachments[0].Name || `attachment-${params.row.UUID || params.row.RequestID}`;
            try {
              const response = await fetch(fileUrl);
              const blob = await response.blob();
              const link = document.createElement('a');
              link.href = window.URL.createObjectURL(blob);
              link.download = fileName;
              document.body.appendChild(link);
              link.click();
              document.body.removeChild(link);
              window.URL.revokeObjectURL(link.href);
            } catch (error) {
              console.error('Error downloading file:', error);
              alert('Failed to download the file. Please try again.');
            }
          }
        };
        const handlePreview = () => {
          if (currentRow?.file) {
            const previewUrl = URL.createObjectURL(currentRow.file);
            const newWindow = window.open(previewUrl, '_blank');
            if (newWindow) {
              newWindow.addEventListener('unload', () => {
                URL.revokeObjectURL(previewUrl);
              });
            }
          } else if (attachments.length > 0) window.open(attachments[0].ServerRelativeUrl, '_blank', 'noopener,noreferrer');
        };
        return (
          <Box display="flex" alignItems="center">
            {(neww || edit) && (
              <IconButton color="primary" component="label" sx={{ ml: 1 }}>
                <FileUploadOutlinedIcon />
                <input type="file" hidden onChange={handleFileChange} disabled={view} />
              </IconButton>
            )}
            {(neww || edit || view) && (attachments.length > 0 || currentRow?.file) && (
              <IconButton color="primary" onClick={handlePreview} sx={{ ml: 1 }}>
                <VisibilityIcon />
              </IconButton>
            )}
            {view && attachments.length > 0 && (
              <IconButton onClick={handleDownload}><DownloadIcon /></IconButton>
            )}
          </Box>
        );
      },
    },
    {
      field: "actions", flex: 0.4,
      renderHeader: () => (
        <span style={{ fontWeight: 'bold' }}></span>
      ),
      renderCell: (params) => (
        <Box display="flex" alignItems="center">
          {(neww || edit) &&
            <IconButton color="error" onClick={() => removeRow(params.id as number)}>
              <RemoveOutlinedIcon />
            </IconButton>
          }
        </Box>
      ),
    },
  ];

  useEffect(() => {
    void fetchAllData();
  }, [fetchAllData, refreshTrigger]);
  useEffect(() => {
    if (ExpenseData.length > 0 && selectedRow) {
      const filtered = ExpenseData.filter((item: any) => item.RequestNo == selectedRow);
      setSelectedExpense(filtered[0]);
      setCompany(filtered[0]?.Company || "");
      setCmsid(filtered[0]?.CMSID || "");
      if (!neww) {
        const fetchAttachments = async () => {
          const inputDataArray = await Promise.all(
            filtered.map(async (item: any) => {
              const attachments = await getAttachments(item.RequestID);
              const existingFile = attachments[0] || null;
              return {
                id: item.ID,
                requestId: item.RequestID,
                expenseDate: parseDateFromSharePoint(item.ExpenseDate),
                merchant: item.Merchant,
                expenseType: item.ExpenseType,
                currency: item.Currency,
                expenseAmount: item.ExpenseAmount,
                multiplier: item.Multiplier,
                totalAmount: item.TotalAmount,
                reason: item.Reason,
                file: null,
                existingFileUrl: existingFile ? existingFile.ServerRelativeUrl : null,
              };
            })
          );
          setInputData(inputDataArray);
        };
        void fetchAttachments();
      }
    }
  }, [ExpenseData, selectedRow]);
  useEffect(() => {
    const total = inputData.reduce((sum, row) => { return sum + (parseFloat(row.totalAmount) || 0); }, 0);
    setGrandTotal(total);
  }, [inputData]);

  return (
    <Paper sx={{ mb: 4, mt: 3, border: "1px solid #ccc", boxShadow: "0px 4px 10px rgba(0, 0, 0, 0.3)", borderRadius: "8px", height: 'auto', position: 'relative' }}>
      <Box display="flex" alignItems="center" justifyContent="space-between" sx={{ bgcolor: 'primary.main', color: 'white', borderTopLeftRadius: 8, borderTopRightRadius: 8 }}>
        <Typography variant="h6" gutterBottom sx={{ bgcolor: 'primary.main', color: 'white', p: 2, borderTopLeftRadius: 8, borderTopRightRadius: 8 }}>
          Expense Application Form
        </Typography>
        <IconButton sx={{ color: 'white', mr: 2 }}>
          <CancelIcon fontSize="medium" onClick={handleExit} />
        </IconButton>
      </Box>
      {neww &&
        <>
          <Box display="flex" flexWrap="wrap" gap={2} sx={{ padding: 3 }}>
            <FormControl sx={{ width: '23%' }}>
              <InputLabel shrink sx={{ color: 'black', fontSize: '1.2rem' }}>Employee ID</InputLabel>
              <Input value={EmployeeDetails.EmployeeID} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
            <FormControl sx={{ width: '23%' }}>
              <InputLabel shrink sx={{ color: 'black', fontSize: '1.2rem' }}>Employee Name</InputLabel>
              <Input value={EmployeeDetails.EmployeeName} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
            <FormControl sx={{ width: '23%' }}>
              <InputLabel shrink sx={{ color: 'black', fontSize: '1.2rem' }}>Department</InputLabel>
              <Input value={EmployeeDetails.Department} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
            <FormControl sx={{ width: '23%' }}>
              <InputLabel shrink sx={{ color: 'black', fontSize: '1.2rem' }}>Country</InputLabel>
              <Input value={EmployeeDetails.Country} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
            <FormControl sx={{ width: '23%' }}>
              <InputLabel shrink sx={{ color: 'black', fontSize: '1.2rem' }}>Submission Date</InputLabel>
              <Input value={new Date().toLocaleDateString('en-GB')} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
            <FormControl sx={{ width: '23%' }}>
              <InputLabel shrink sx={{ color: 'black', fontSize: '1.2rem' }}>Manager</InputLabel>
              <Input value={EmployeeDetails.Manager} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
            <FormControl sx={{ width: '23%' }}>
              <InputLabel shrink sx={{ color: 'black', fontSize: '1.2rem' }}>Local Currency</InputLabel>
              <Input value="INR" disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
            <TextField select label="Company" value={company} onChange={(e) => setCompany(e.target.value)} sx={{ width: '48%' }}>
              <MenuItem value="Credent Infotech Solutions LLP">Credent Infotech Solutions LLP</MenuItem>
              <MenuItem value="IPAI Technology Solutions LLP">IPAI Technology Solutions LLP</MenuItem>
            </TextField>
            <TextField select label="CMS ID" value={cmsid} onChange={(e) => setCmsid(e.target.value)} sx={{ width: '47%' }}>
              {CMSID.map((item, index) => (
                <MenuItem key={index} value={item.RequestID}>{item.RequestID}</MenuItem>
              ))}
            </TextField>
          </Box>
          <Box display="flex" justifyContent="space-between" sx={{ bgcolor: 'primary.main', color: 'white', p: 2, borderTopLeftRadius: 8, borderTopRightRadius: 8 }}>
            <Typography variant="h6" gutterBottom>Expense Details</Typography>
            <IconButton onClick={addRow} sx={{ color: 'white' }}><AddCircleOutlinedIcon /></IconButton>
          </Box>
          <Box sx={{ m: 2 }}>
            <DataGrid
              rows={inputData}
              columns={columns}
              paginationModel={paginationModel}
              onPaginationModelChange={setPaginationModel}
              pagination
              pageSizeOptions={[10, 25, 50]}
              disableRowSelectionOnClick
              initialState={{
                pagination: { paginationModel: { pageSize: 10, }, },
              }}
              sx={{
                "& .MuiDataGrid-columnHeaders": {
                  whiteSpace: "normal !important",
                  bgcolor: 'primary.main',
                  color: 'white',
                  lineHeight: '1.2 !important',
                  wordBreak: 'break-word !important',
                },
                border: "2px solid #ccc",
                borderRadius: "8px",
              }}
              rowHeight={60}
              slots={{
                footer: () => <GrandTotal grandTotal={grandTotal} />
              }}
            />
            {viewTable &&
              <Box display='flex' justifyContent="space-between">
                <TextField label="Purpose (Optional)" multiline rows={2} value={purpose} onChange={(e) => setPurpose(e.target.value)} sx={{ mt: 1, width: '40%' }} />
                <ExpenseTable ExpenseType={expenseType} Currency={currency} rows={inputData} />
              </Box>
            }
            <Box display="flex" justifyContent="flex-start" mt={2} gap={1}>
              <Button variant="contained" color="success" size="large" sx={{ borderRadius: 2, textTransform: "none", }} onClick={handleSubmit}>
                Send For Approval
              </Button>
              <Button variant="contained" color="warning" size="large" sx={{ borderRadius: 2, textTransform: "none", }} onClick={handleDraft}>
                Save as Draft
              </Button>
              <Button variant="contained" color="error" size="large" sx={{ borderRadius: 2, textTransform: "none", }} onClick={handleExit}>
                Exit
              </Button>
            </Box>
          </Box>
        </>
      }
      {view &&
        <>
          <Box display="flex" flexWrap="wrap" gap={2} sx={{ padding: 3 }}>
            <FixedEmployeeDetails selectedExpense={selectedExpense} />
            <FormControl sx={{ width: '48%' }}>
              <InputLabel shrink sx={{ color: 'black' }}>Company</InputLabel>
              <Input value={selectedExpense.Company} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
            <FormControl sx={{ width: '47%' }}>
              <InputLabel shrink sx={{ color: 'black' }}>CMS ID</InputLabel>
              <Input value={selectedExpense.CMSID} disabled sx={{ '& .MuiInput-input.Mui-disabled': { WebkitTextFillColor: 'black' } }} />
            </FormControl>
          </Box>
          <Box display="flex" justifyContent="space-between" sx={{ bgcolor: 'primary.main', color: 'white', p: 2, borderTopLeftRadius: 8, borderTopRightRadius: 8 }}>
            <Typography variant="h6" gutterBottom>Expense Details</Typography>
          </Box>
          <Box sx={{ m: 2 }}>
            <DataGrid
              rows={inputData.map((item) => ({
                id: item.id,
                UUID: item.requestId,
                merchant: item.merchant,
                amount: item.expenseAmount,
                expenseDate: item.expenseDate ? new Date(item.expenseDate).toLocaleDateString('en-GB') : "",
                expenseType: item.expenseType,
                currency: item.currency,
                expenseAmount: item.expenseAmount,
                multiplier: item.multiplier,
                totalAmount: item.totalAmount,
                reason: item.reason,
              }))}
              columns={columns}
              getRowId={(row) => row.id}
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
              rowHeight={60}
              sx={{
                "& .MuiDataGrid-columnHeaders": {
                  whiteSpace: "normal !important",
                  bgcolor: 'primary.main',
                  color: 'white',
                  lineHeight: '1.2 !important',
                  wordBreak: 'break-word !important',
                },
                "& .MuiDataGrid-cell": {
                  pointerEvents: 'none',
                  backgroundColor: '#f5f5f5',
                  "& *": {
                    color: 'grey !important',
                  },
                },
                "& .interactive-cell": {
                  pointerEvents: 'auto !important',
                  "& *": {
                    color: 'blue !important',
                  },
                },
                border: "2px solid #ccc",
                borderRadius: "8px",
              }}
              slots={{
                footer: () => <GrandTotal grandTotal={grandTotal} />
              }}
            />
            <Box display="flex">
              <Box display="flex" flexDirection="column" width="50%">
                {selectedExpense.Purpose &&
                  <TextField multiline rows={2} sx={{ mt: 1, width: '60%' }} value={selectedExpense.Purpose} disabled />
                }
                {selectedExpense.ManagerEmail === context.pageContext.user.email && (selectedExpense.Status === "Pending at Manager" || selectedExpense.Status === "Pending at Finance") &&
                  <TextField label={
                    <span>
                      Remarks <span style={{ color: 'red', fontSize: '1.3rem' }}>*</span>
                    </span>
                  }
                    multiline rows={2} value={remarks} onChange={(e) => setRemarks(e.target.value)} sx={{ mt: 1, width: '60%' }} disabled={!isManager} />
                }
              </Box>
              {filteredApprovalHistory.length !== 0 &&
                <ApprovalHistory selectedExpense={selectedExpense} rows={approvalHistory} />
              }
            </Box>
            <Box display="flex" justifyContent="end" mt={2}>
              {!isManager || selectedExpense.Status === "Approved" || selectedExpense.Status === "Rejected" || selectedExpense.ManagerEmail !== context.pageContext.user.email ?
                (
                  <Button variant="contained" color="error" size="large"
                    sx={{ borderRadius: 2, textTransform: "none" }}
                    onClick={() => setSelectedMenu("Request")}>
                    Close
                  </Button>
                )
                :
                (
                  <Box display="flex" justifyContent="flex-start" mt={2} sx={{ width: '100%' }} gap={1}>
                    <Button variant="contained" color="success" size="large"
                      sx={{ borderRadius: 2, textTransform: "none" }}
                      onClick={handleApprove}>
                      Approve
                    </Button>
                    <Button variant="contained" color="warning" size="large"
                      sx={{ borderRadius: 2, textTransform: "none" }}
                      onClick={handleRevision}>
                      Request Revision
                    </Button>
                    <Button variant="contained" color="error" size="large"
                      sx={{ borderRadius: 2, textTransform: "none" }}
                      onClick={handleReject}>
                      Reject
                    </Button>
                  </Box>

                )}
            </Box>
          </Box>
        </>
      }
      {edit &&
        <>
          <Box display="flex" flexWrap="wrap" gap={2} sx={{ padding: 3 }}>
            <FixedEmployeeDetails selectedExpense={selectedExpense} />
            <TextField select label={selectedExpense.Company || "Loading..."} value={company} onChange={(e) => setCompany(e.target.value)} sx={{ width: '48%' }}>
              <MenuItem value="Credent Infotech Solutions LLP">Credent Infotech Solutions LLP</MenuItem>
              <MenuItem value="IPAI Technology Solutions LLP">IPAI Technology Solutions LLP</MenuItem>
            </TextField>
            <TextField select label={selectedExpense.CMSID || "Loading..."} value={cmsid} onChange={(e) => setCmsid(e.target.value)} sx={{ width: '47%' }}>
              {CMSID.map((item, index) => (
                <MenuItem key={index} value={item.RequestID}>{item.RequestID}</MenuItem>
              ))}
            </TextField>
          </Box>
          <Box display="flex" justifyContent="space-between" sx={{ bgcolor: 'primary.main', color: 'white', p: 2, borderTopLeftRadius: 8, borderTopRightRadius: 8 }}>
            <Typography variant="h6" gutterBottom>Expense Details</Typography>
            <IconButton onClick={addRow} sx={{ color: 'white' }}><AddCircleOutlinedIcon /></IconButton>
          </Box>
          <Box sx={{ m: 2 }}>
            <DataGrid
              rows={inputData.map((item) => ({
                id: item.id,
                UUID: item.requestId,
                merchant: item.merchant,
                amount: item.expenseAmount,
                expenseDate: item.expenseDate ? new Date(item.expenseDate).toLocaleDateString('en-GB') : "",
                expenseType: item.expenseType,
                currency: item.currency,
                expenseAmount: item.expenseAmount,
                multiplier: item.multiplier,
                totalAmount: item.totalAmount,
                reason: item.reason,
              }))}
              columns={columns}
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
                "& .MuiDataGrid-columnHeaders": {
                  whiteSpace: "normal !important",
                  bgcolor: 'primary.main',
                  color: 'white',
                  lineHeight: '1.2 !important',
                  wordBreak: 'break-word !important',
                },
                border: "2px solid #ccc",
                borderRadius: "8px",
              }}
              rowHeight={60}
              slots={{
                footer: () => <GrandTotal grandTotal={grandTotal} />
              }}
            />
            {viewTable &&
              <Box display='flex' justifyContent="space-between">
                <TextField multiline rows={2} sx={{ mt: 1, width: '40%' }} value={selectedExpense.Purpose} onChange={(e) => setSelectedExpense({ ...selectedExpense, Purpose: e.target.value })} />
                <ExpenseTable ExpenseType={expenseType} Currency={currency} rows={inputData} />
              </Box>
            }
            <Box display="flex" justifyContent="end" mt={2}>
              <Box display="flex" justifyContent="flex-start" mt={2} sx={{ width: '100%' }} gap={1}>
                <Button variant="contained" color="success" size="large" sx={{ borderRadius: 2, textTransform: "none", }}
                  onClick={handleUpdateData}>
                  Send For Approval
                </Button>
                {selectedExpense.Status === "Draft" &&
                  <>
                    <Button variant="contained" color="warning" size="large" sx={{ borderRadius: 2, textTransform: "none", }} onClick={handleDraft}>
                      Save as Draft
                    </Button>
                    <Button variant="contained" color="error" size="large" sx={{ borderRadius: 2, textTransform: "none", }}
                      onClick={handleDelete}>
                      Delete
                    </Button>
                  </>
                }
                <Button variant="contained" color="primary" size="large" sx={{ borderRadius: 2, textTransform: "none", }}
                  onClick={handleExit}>
                  Cancel
                </Button>
              </Box>
            </Box>
          </Box>
        </>
      }
    </Paper>
  );
};
export default New;