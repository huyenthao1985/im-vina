const http = require('http');
const fs = require('fs');
const path = require('path');

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
    
    console.log(`Found ${tabs.length} tabs for screenshotting:`);
    for (let i = 0; i < tabs.length; i++) {
      const tab = tabs[i];
      // Let's identify if it is old or new dashboard based on its panels
      const isNew = await checkIsNew(tab.webSocketDebuggerUrl);
      const label = isNew ? 'new' : 'old';
      console.log(`Tab ${i + 1}: URL="${tab.url}", Title="${tab.title}" is identified as: ${label}`);
      await capture(tab.webSocketDebuggerUrl, `dashboard_${label}_final.png`);
    }
  } catch (err) {
    console.error("Error in main:", err);
  }
}

function checkIsNew(wsUrl) {
  return new Promise((resolve) => {
    const ws = new WebSocket(wsUrl);
    ws.onopen = () => {
      ws.send(JSON.stringify({
        id: 1,
        method: 'Runtime.evaluate',
        params: {
          expression: `!!document.querySelector('.second-dashboard')`,
          returnByValue: true
        }
      }));
    };
    ws.onmessage = (event) => {
      const msg = JSON.parse(event.data);
      if (msg.result && msg.result.result) {
        resolve(msg.result.result.value);
      } else {
        resolve(false);
      }
      ws.close();
    };
    ws.onerror = () => {
      resolve(false);
    };
  });
}

function capture(wsUrl, filename) {
  return new Promise((resolve) => {
    const ws = new WebSocket(wsUrl);
    let id = 1;

    ws.onopen = () => {
      console.log(`Capturing screenshot for ${filename}...`);
      ws.send(JSON.stringify({
        id: id++,
        method: 'Page.captureScreenshot',
        params: {
          format: 'png',
          captureBeyondViewport: false
        }
      }));
    };

    ws.onmessage = (event) => {
      const msg = JSON.parse(event.data);
      if (msg.result && msg.result.data) {
        const imgBuffer = Buffer.from(msg.result.data, 'base64');
        const destPath = path.join('C:\\Users\\dell\\.gemini\\antigravity-ide\\brain\\406c864b-6932-4298-a7d8-ea4e2c69f0a8', filename);
        fs.writeFileSync(destPath, imgBuffer);
        console.log(`Screenshot saved to: ${destPath}`);
        ws.close();
        resolve();
      } else {
        ws.close();
        resolve();
      }
    };

    ws.onerror = (err) => {
      console.error("WS error during capture:", err);
      resolve();
    };
  });
}

main();
