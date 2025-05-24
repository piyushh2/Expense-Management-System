import * as React from "react";
import { AppBar, Toolbar, Typography } from "@mui/material";

const Footer: React.FC = () => {
  return (
    <AppBar component="footer" position="static" color="primary" sx={{top:"auto",bottom:0,}}>
      <Toolbar sx={{ justifyContent: "center", py: 1 }}>
        <Typography variant="subtitle1" color="inherit">
          © {new Date().getFullYear()} Credent Infotech. All rights reserved.
        </Typography>
      </Toolbar>
    </AppBar>
  );
};

export default Footer;