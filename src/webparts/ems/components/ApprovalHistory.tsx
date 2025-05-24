import * as React from 'react';
import { Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Paper, Box, } from '@mui/material';

interface ApprovalHistoryEntry {
  RequestNo: any;
  approvalDate: string;
  approver: string;
  remarks: string;
}
interface ApprovalHistoryProps {
  selectedExpense: any;
  rows: ApprovalHistoryEntry[];
}

const ApprovalHistory: React.FC<ApprovalHistoryProps> = ({ selectedExpense,rows }) => {
  const formatDate = (date: string): string => {
    if (!date) return '';
    const parsedDate = new Date(date);
    return isNaN(parsedDate.getTime()) ? 'Invalid Date' : parsedDate.toLocaleDateString('en-GB');
  };

  return (
    <Box sx={{ m: 2, width: '40%' }}>
      <TableContainer component={Paper} sx={{ maxWidth: 800, height: 'fit-content', borderRadius: 2, boxShadow: 3 }}>
        <Table>
          <TableHead>
            <TableRow sx={{ bgcolor: 'primary.main' }}>
              <TableCell sx={{ color: 'white', fontWeight: 'bold' }}>Date</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 'bold' }}>Approver</TableCell>
              <TableCell sx={{ color: 'white', fontWeight: 'bold' }}>Remarks</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
             {rows.filter((row) => row.RequestNo === String(selectedExpense.RequestNo)).map((row) => (
              <TableRow key={row.RequestNo}>
                <TableCell>{formatDate(row.approvalDate)}</TableCell>
                <TableCell>{row.approver}</TableCell>
                <TableCell>{row.remarks}</TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </TableContainer>
    </Box>
  );
};

export default ApprovalHistory;