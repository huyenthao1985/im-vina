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
    const tabs = pages.filter(p => p.type === 'page' && p.url.includes('localhost:5173') && !p.url.includes('TargetActualDashboard.tsx'));
    
    if (tabs.length === 0) {
      console.log("No localhost:5173 tabs found!");
      return;
    }
    
    const wsUrl = tabs[0].webSocketDebuggerUrl;
    console.log(`Connecting to tab: ${tabs[0].url}`);
    
    const ws = new WebSocket(wsUrl);
    ws.onopen = () => {
      console.log("Connected. Enabling Runtime and Log domains...");
      ws.send(JSON.stringify({ id: 1, method: 'Runtime.enable' }));
      ws.send(JSON.stringify({ id: 2, method: 'Log.enable' }));
      
      // Let's also evaluate the current state of the upload in the app
      ws.send(JSON.stringify({
        id: 3,
        method: 'Runtime.evaluate',
        params: {
          expression: `(() => {
            const uploadError = document.querySelector('.hero div[style*="color: var(--rose)"]')?.innerText;
            const emptyPanel = document.querySelector('.panel-empty')?.innerText;
            const isPerCapitaVisible = !!document.getElementById('pc-chart-timeline');
            
            return JSON.stringify({
              uploadError: uploadError || null,
              emptyPanel: emptyPanel || null,
              isPerCapitaVisible
            });
          })()`,
          returnByValue: true
        }
      }));
    };
    
    let receivedConsole = false;
    ws.onmessage = (event) => {
      const msg = JSON.parse(event.data);
      if (msg.method === 'Runtime.consoleAPICalled') {
        console.log('[Browser Console]', msg.params.args.map(arg => arg.value || arg.description).join(' '));
        receivedConsole = true;
      } else if (msg.method === 'Log.entryAdded') {
        console.log('[Browser Log]', msg.params.entry.text);
        receivedConsole = true;
      } else if (msg.id === 3) {
        console.log("Page DOM State:", JSON.parse(msg.result.result.value));
      }
    };
    
    // Let's run for 2 seconds to collect logs
    setTimeout(() => {
      ws.close();
      console.log("Finished log inspection.");
    }, 2000);
    
    ws.onerror = (err) => {
      console.error("WS error:", err);
    };
  } catch (err) {
    console.error("Error:", err);
  }
}

main();
