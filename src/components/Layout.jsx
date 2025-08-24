import { NavLink, Outlet } from "react-router-dom";
import "./layout.css"; // optional styles below

export default function Layout() {
  return (
    <div className="app">
      <aside className="sidebar">
        <h2>Digital Toolbox</h2>
        <nav>
          <NavLink to="/dashboard" className="link">Dashboard</NavLink>
          <NavLink to="/file-merge" className="link">File Merge</NavLink>
          <NavLink to="/doc-pdf" className="link">Doc to PDF Converter</NavLink>
          <NavLink to="/json-csv" className="link">Json to CSV Converter</NavLink>
          </nav>
      </aside>

      <main className="content">
        <Outlet />
      </main>
    </div>
  );
}
