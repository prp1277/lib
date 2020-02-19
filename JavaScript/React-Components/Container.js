import React from "react";
import Main from "./Main";
import Sidebar from "./Sidebar";

const Container = () => {
  return (
    <div className="container-fluid">
      <div className="row">
        <Sidebar />
        <Main />
      </div>
    </div>
  );
};
export default Container;
