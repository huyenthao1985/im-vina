const http = require('http');
const fs = require('fs');

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
    const tab = pages.find(p => p.type === 'page' && p.url.includes('localhost:5173'));
    if (!tab) {
      console.error("No active localhost:5173 tab found!");
      return;
    }
    
    console.log(`Connecting to tab: ${tab.url}`);
    const ws = new WebSocket(tab.webSocketDebuggerUrl);
    
    ws.onopen = () => {
      console.log("WebSocket connected. Enabling Runtime and Page...");
      ws.send(JSON.stringify({ id: 1, method: 'Runtime.enable' }));
      ws.send(JSON.stringify({ id: 2, method: 'Page.enable' }));
      
      // We will evaluate script to change select value and trigger change event
      setTimeout(() => {
        console.log("Evaluating script to switch CHI TIẾT to 'ALL'...");
        ws.send(JSON.stringify({
          id: 3,
          method: 'Runtime.evaluate',
          params: {
            expression: `(() => {
              // Find the custom select or native select for Period (CHI TIẾT)
              // Let's find custom select trigger that contains "2026" or click it
              const selects = document.querySelectorAll('select');
              console.log('Select count:', selects.length);
              
              // We have CustomSelect. Find the trigger button or check what components are there
              // Let's evaluate a script that will click the option "Tất cả"
              // In CustomSelect, we can find elements with class "csel-trigger"
              const triggers = document.querySelectorAll('.csel-trigger');
              console.log('Custom Select triggers:', triggers.length);
              
              // Let's find the Period select: the options list contains "Tất cả", "2025", "2026" etc.
              // We can click the second custom select trigger (the first ones are yearFrom, yearTo, origin, selModel, viewMode, selectedPeriod)
              // Wait, let's find the trigger that has text containing "2026" or "2025" or click them and inspect the csel-menu options
              // Let's inspect the DOM structure for custom selects.
              // Alternatively, we can just trigger the select change programmatically if we can find the react state setter,
              // or click the csel-trigger that controls selectedPeriod.
              // Let's find all csel-triggers. The trigger for selectedPeriod is the one that has width 140px or is next to viewMode toggle.
              // Let's click the csel-trigger and then find csel-option with text "Tất cả".
              // Let's try:
              const periodTrigger = Array.from(document.querySelectorAll('.csel-trigger')).find(el => {
                const txt = el.textContent;
                return txt.includes('2026') || txt.includes('2025') || txt.includes('Tất cả') || txt.includes('All') || txt.includes('전체');
              });
              if (periodTrigger) {
                console.log('Found period trigger:', periodTrigger.textContent);
                periodTrigger.click(); // Open dropdown
                setTimeout(() => {
                  const options = Array.from(document.querySelectorAll('.csel-option'));
                  const allOpt = options.find(opt => opt.textContent.includes('Tất cả') || opt.textContent.includes('All') || opt.textContent.includes('전체'));
                  if (allOpt) {
                    console.log('Found All option, clicking:', allOpt.textContent);
                    allOpt.click();
                  } else {
                    console.log('Could not find All option. Options found:', options.map(o=>o.textContent.trim()).join(', '));
                  }
                }, 500);
              } else {
                console.log('Could not find period trigger!');
              }
            })()`,
            returnByValue: true
          }
        }));
      }, 1000);
    };
    
    ws.onmessage = (event) => {
      const msg = JSON.parse(event.data);
      if (msg.method === 'Runtime.exceptionThrown') {
        console.log("\n!!! EXCEPTION THROWN IN PAGE !!!");
        console.log(JSON.stringify(msg.params.exceptionDetails, null, 2));
      }
      if (msg.method === 'Runtime.consoleAPICalled') {
        const args = msg.params.args.map(a => a.value || a.description || JSON.stringify(a));
        console.log(`[Console ${msg.params.type}]`, ...args);
      }
    };
    
    ws.onerror = (err) => {
      console.error("WebSocket error:", err);
    };
    
    // Keep running for 5 seconds to capture errors
    setTimeout(() => {
      console.log("Finished listening.");
      ws.close();
    }, 6000);
    
  } catch (err) {
    console.error("Error:", err);
  }
}

main();
