import './App.css';
import { Routes, Route, Navigate } from "react-router-dom";
import Layout from "./components/Layout.jsx"; // Updated to use .jsx extension
import Dashboard from "./pages/dashboard.js";
import FileMerge from './pages/fileMerge.js';
import DocToPdf from './pages/docToPdf.js';
import JsonToCsv from './pages/jsonToCsv.js';

function App() {
  return (
    <Routes>
      {/* Root uses the shared layout */}
      <Route element={<Layout />}>
        {/* Default route goes to /dashboard */}
        <Route path="/" element={<Navigate to="/dashboard" replace />} />
        <Route path="/dashboard" element={<Dashboard />} />
        <Route path="/file-merge" element={<FileMerge />} />
        <Route path="/doc-pdf" element={<DocToPdf />} />
        <Route path="/json-csv" element={<JsonToCsv />} />
        {/* Optional: 404 */}
        <Route path="*" element={<Navigate to="/dashboard" replace />} />
      </Route>
    </Routes>
  );
}

export default App;