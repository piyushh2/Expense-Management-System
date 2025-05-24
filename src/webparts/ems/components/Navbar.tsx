import * as React from "react";
import { AppBar, Toolbar, Typography, IconButton, Box } from "@mui/material";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faHouseChimney, faFile, faBook, faLock, faFileLines, faChartLine } from "@fortawesome/free-solid-svg-icons";

const Navbar: React.FC = () => {
  const icons = [
    { icon: faHouseChimney, label: "Home" },
    { icon: faFile, label: "Documents" },
    { icon: faBook, label: "HRDMS" },
    { icon: faLock, label: "Policies" },
    { icon: faFileLines, label: "E-Forms" },
    { icon: faChartLine, label: "Org Chart" },
  ];

  return (
    <AppBar position="static" color="primary" sx={{ padding: "5px 20px" }}>
      <Toolbar sx={{ display: "flex", alignItems: "center" }}>
        <Box sx={{ display: "flex", alignItems: "center", gap: 2 }}>
          <img src={require('../assets/logo.png')} alt="Logo" style={{ height: 50 }} />
          <Typography variant="h5" sx={{ fontWeight: "bold", color: "white" }}>CREDENT INFOTECH</Typography>
        </Box>
        <Box sx={{ flexGrow: 1 }} />
        <Box sx={{ display: "flex", gap: 3 }}>
          {icons.map(({ icon, label }, index) => (
            <Box key={index} sx={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
              <IconButton color="inherit" sx={{ padding: "8px" }}><FontAwesomeIcon icon={icon} size="sm" /></IconButton>
              <Typography variant="caption" sx={{ color: "white" }}>{label}</Typography>
            </Box>
          ))}
        </Box>
      </Toolbar>
    </AppBar>
  );
};

export default Navbar;