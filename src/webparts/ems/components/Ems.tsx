/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises*/

import * as React from 'react';
import { IEmsProps } from './IEmsProps';
import Navbar from './Navbar';
import Footer from './Footer';
import Dashboard from './Dashboard';

const Ems: React.FC<IEmsProps> = (props) => {
  return (
    <div style={{ display: "flex",width:'full', flexDirection: "column", minHeight: "100dvh" }}>
      <Navbar />
      <Dashboard siteUrl={props.siteUrl} context={props.context} />
      <Footer />
    </div>
  );
};

export default Ems;