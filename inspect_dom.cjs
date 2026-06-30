const http = require('http');

function getJSON(url) {
  return new Promise((resolve, reject) => {
    http.get(url, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => resolve(JSON.parse(data)));
    }).on('error', reject);
  });
}

async function main() {
  try {
    const pages = await getJSON('http://127.0.0.1:9222/json');
    const tabs = pages.filter(p => p.type === 'page' && p.url.includes('localhost:5173'));
    
    if (tabs.length === 0) {
      console.log("No localhost:5173 tabs found!");
      return;
    }
    
    const wsUrl = tabs[0].webSocketDebuggerUrl;
    console.log(`Connecting to tab: ${tabs[0].url}`);
    
    const ws = new WebSocket(wsUrl);
    ws.onopen = () => {
      ws.send(JSON.stringify({
        id: 1,
        method: 'Runtime.evaluate',
        params: {
          expression: `(() => {
            const salesDashboard = document.querySelector('.sales-dashboard');
            const panel = document.querySelector('.panel');
            const tableContainer = document.querySelector('.table-container');
            const paginationBar = document.querySelector('.pagination-bar');
            const appContent = document.querySelector('.app-content');
            const appLayout = document.querySelector('.app-layout');
            
            const pgRect = paginationBar ? paginationBar.getBoundingClientRect() : null;
            
            return JSON.stringify({
              salesDashboard: salesDashboard ? {
                heightStyle: salesDashboard.style.height,
                clientHeight: salesDashboard.clientHeight,
                offsetHeight: salesDashboard.offsetHeight,
                computedHeight: window.getComputedStyle(salesDashboard).height,
                computedDisplay: window.getComputedStyle(salesDashboard).display,
              } : null,
              panel: panel ? {
                clientHeight: panel.clientHeight,
                offsetHeight: panel.offsetHeight,
                computedHeight: window.getComputedStyle(panel).height,
                computedDisplay: window.getComputedStyle(panel).display,
                computedFlex: window.getComputedStyle(panel).flex,
              } : null,
              tableContainer: tableContainer ? {
                clientHeight: tableContainer.clientHeight,
                offsetHeight: tableContainer.offsetHeight,
                computedHeight: window.getComputedStyle(tableContainer).height,
                computedDisplay: window.getComputedStyle(tableContainer).display,
              } : null,
              paginationBar: pgRect ? {
                top: pgRect.top,
                bottom: pgRect.bottom,
                height: pgRect.height,
              } : null,
              appContent: appContent ? {
                clientHeight: appContent.clientHeight,
                offsetHeight: appContent.offsetHeight,
                computedHeight: window.getComputedStyle(appContent).height,
              } : null,
              appLayout: appLayout ? {
                clientHeight: appLayout.clientHeight,
                offsetHeight: appLayout.offsetHeight,
                computedHeight: window.getComputedStyle(appLayout).height,
              } : null,
            });
          })()`,
          returnByValue: true
        }
      }));
    };
    ws.onmessage = (event) => {
      const msg = JSON.parse(event.data);
      if (msg.result && msg.result.result) {
        const value = msg.result.result.value;
        console.log("DOM Properties:", JSON.parse(value));
      } else {
        console.log("No result value in message:", msg);
      }
      ws.close();
    };
    ws.onerror = (err) => {
      console.error("WS error:", err);
    };
  } catch (err) {
    console.error("Error:", err);
  }
}

main();
