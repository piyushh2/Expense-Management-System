import * as React from "react";
import { useState } from "react";
import { Box, Tabs, Tab } from "@mui/material";
import Request from "./Request";
import Approvals from "./Approvals";

interface IDashboardProps {
  siteUrl: string;
  context: any;
}
const Dashboard: React.FC<IDashboardProps> = (props) => {
  const [view, setView] = useState("requests");
  const handleChange = (_event: React.SyntheticEvent, newValue: string) => {
    setView(newValue);
  };

  return (
    <Box sx={{ mt: 2, px: 3, display: "flex", flexDirection: "column", height: "auto" }}>
      <Tabs value={view} onChange={handleChange} variant="fullWidth" textColor="primary" indicatorColor="primary"
        sx={{
          '& .MuiTabs-indicator': {
            height: '4px',
          },
        }}>
        <Tab label="My Requests" value="requests"
          sx={{
            fontWeight: '700',
            fontSize: '1.1rem',
            borderBottom: '4px solid lightgrey',
            '&.Mui-selected': {
              borderBottom: 'none',
            },
          }} />
        <Tab label="Requests Waiting for Your Approval" value="approvals"
          sx={{
            fontWeight: '700',
            fontSize: '1.1rem',
            borderBottom: '4px solid lightgrey',
            '&.Mui-selected': {
              borderBottom: 'none',
            },
          }} />
      </Tabs>
      <Box sx={{ mt: 2 }}>
        {view === "requests" ?
          (
            <Request siteUrl={props.siteUrl} context={props.context} />
          )
          :
          (
            <Approvals siteUrl={props.siteUrl} context={props.context} />
          )
        }
      </Box>
    </Box>
  );
};

export default Dashboard;