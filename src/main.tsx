import React from 'react'
import * as ReactDOM from 'react-dom/client'
import App from './App.tsx'

declare global {
  interface Window {
    ZOHO?: {
      embeddedApp: {
        on: (event: string, callback: (data: any) => void) => void;
        init: () => void;
      }
    }
  }
}

const renderApp = (data?: any) => {
  const root = ReactDOM.createRoot(document.getElementById('root')!);
  root.render(
    <React.StrictMode>
      <App data={data} />
    </React.StrictMode>
  );
}

if (window.ZOHO && window.ZOHO.embeddedApp) {
  window.ZOHO.embeddedApp.on("PageLoad", function(data) {
    renderApp(data);
  });

  window.ZOHO.embeddedApp.init();
  console.log('ZOHO embedded app initialized');
} else {
  console.log('ZOHO embedded app not found, rendering app without data');
  renderApp();
}
