import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import './index.css';
import App from './App.jsx';
import ParentPortal from './ParentPortal.jsx';

const params = new URLSearchParams(window.location.search);
const isParentView = params.get('view') === 'parent';

createRoot(document.getElementById('root')).render(
  <StrictMode>
    {isParentView ? <ParentPortal /> : <App />}
  </StrictMode>
);