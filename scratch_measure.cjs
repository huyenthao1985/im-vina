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
    const tabs = pages.filter(p => p.type === 'page');
    for (const tab of tabs) {
      if (tab.url.includes('localhost:5173') && !tab.url.includes('TargetActualDashboard.tsx')) {
        console.log(`Inspecting Tab: "${tab.title}" (${tab.url})`);
        await inspectTab(tab.webSocketDebuggerUrl);
      }
    }
  } catch (err) {
    console.error("Error in main:", err);
  }
}

function inspectTab(wsUrl) {
  return new Promise((resolve) => {
    const ws = new WebSocket(wsUrl);
    let id = 1;

    ws.onopen = () => {
      const expr = `(() => {
        const panels = Array.from(document.querySelectorAll('.panel.chart-panel')).map((p, idx) => {
          const header = p.querySelector('.card-header-styled');
          const holder = p.querySelector('.chart-holder');
          const svg = p.querySelector('svg');
          const plotly = p.querySelector('.js-plotly-plot');
          
          return {
            index: idx + 1,
            title: header ? header.innerText : 'No Header',
            panelHeight: p.getBoundingClientRect().height,
            headerHeight: header ? header.getBoundingClientRect().height : 0,
            holderHeight: holder ? holder.getBoundingClientRect().height : 0,
            svgHeight: svg ? svg.getBoundingClientRect().height : 0,
            plotlyHeight: plotly ? plotly.getBoundingClientRect().height : 0,
            holderChildrenCount: holder ? holder.children.length : 0
          };
        });
        
        return {
          href: window.location.href,
          panels: panels
        };
      })()`;

      ws.send(JSON.stringify({
        id: id++,
        method: 'Runtime.evaluate',
        params: {
          expression: expr,
          returnByValue: true
        }
      }));
    };

    ws.onmessage = (event) => {
      const msg = JSON.parse(event.data);
      if (msg.result && msg.result.result) {
        console.log("Tab panels heights details:", JSON.stringify(msg.result.result.value, null, 2));
        ws.close();
        resolve();
      } else {
        ws.close();
        resolve();
      }
    };

    ws.onerror = (err) => {
      console.error("WS error:", err);
      resolve();
    };
  });
}

main();
