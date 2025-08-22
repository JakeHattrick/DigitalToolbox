import { NavLink, Outlet } from "react-router-dom";
import "./layout.css"; // optional styles below

export default function Layout() {
  return (
    <div className="app">
      <aside className="sidebar">
        <h2>Digital Toolbox</h2>
        <nav>
          <NavLink to="/dashboard" className="link">Dashboard</NavLink>
          </nav>
      </aside>

      <main className="content">
        <Outlet />
      </main>
    </div>
  );
}
