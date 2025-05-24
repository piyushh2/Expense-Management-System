import * as React from 'react';
import { Table, TableBody, TableCell, TableContainer, TableHead, TableRow, Paper, Box } from '@mui/material';

interface ExpenseTableProps {
  ExpenseType: any[];
  Currency: any[];
  rows: {
    expenseType: string;
    currency: string;
    totalAmount: string;
  }[];
}

function ExpenseTable({ ExpenseType = [], Currency = [], rows = [] }: ExpenseTableProps) {

  function getTotalByType(type: string): number {
    return rows
      .filter(row => row.expenseType === type)
      .reduce((sum, row) => sum + parseFloat(row.totalAmount || "0"), 0);
  }
  function getTotalByCurrency(curr: string): number {
    return rows
      .filter(row => row.currency === curr)
      .reduce((sum, row) => sum + parseFloat(row.totalAmount || "0"), 0);
  }

  return (
    <Box display="flex" width="100%" justifyContent="space-evenly" flexWrap="wrap" marginTop={1}>
      <TableContainer component={Paper} sx={{
        width: 300, height: 320, borderRadius: 2, boxShadow: 3, overflowY: 'auto',
        '&::-webkit-scrollbar': {
          width: '0px',
          height: '0px',
        },
        '&:hover::-webkit-scrollbar': {
          width: '4px',
        },
        '&::-webkit-scrollbar-track': {
          background: '#f1f1f1',
        },
        '&::-webkit-scrollbar-thumb': {
          background: '#888',
          transition: 'background-color 0.3s ease',
        },
        '&::-webkit-scrollbar-thumb:hover': {
          background: '#555',
        },
      }} >
        <Table>
          <TableHead>
            <TableRow sx={{ bgcolor: 'primary.main' }}>
              <TableCell sx={{ color: 'white', fontWeight: 'bold' }}>Expense Type</TableCell>
              <TableCell align="right" sx={{ color: 'white', fontWeight: 'bold' }}>Amount</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {ExpenseType.map((item) => (
              <TableRow key={item.Title}>
                <TableCell>{item.Title}</TableCell>
                <TableCell align="right">{getTotalByType(item.Title).toFixed(2)}</TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </TableContainer>

      <TableContainer component={Paper} sx={{ width: 300, height: 'fit-content', borderRadius: 2, boxShadow: 3 }}>
        <Table>
          <TableHead>
            <TableRow sx={{ bgcolor: 'primary.main' }}>
              <TableCell sx={{ color: 'white', fontWeight: 'bold' }}>Currency</TableCell>
              <TableCell align="right" sx={{ color: 'white', fontWeight: 'bold' }}>Amount</TableCell>
            </TableRow>
          </TableHead>
          <TableBody>
            {Currency.map((curr, index) => (
              <TableRow key={index}>
                <TableCell>{curr.Currency}</TableCell>
                <TableCell align="right">
                  {getTotalByCurrency(curr.Currency).toFixed(2)}
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </TableContainer>

    </Box>
  );
}
export default ExpenseTable;