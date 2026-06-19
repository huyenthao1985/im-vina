const fs = require('fs');
const path = 'src/components/SalesDashboard.tsx';
let lines = fs.readFileSync(path, 'utf8').split('\n');

// Find the start of the monthly section
const startIdx = lines.findIndex(l => l.includes('===== Monthly Section (Row 3 Image1 inspired) - Plotly charts ====='));
// Find the line with monthsFullList
const endIdx = lines.findIndex(l => l.includes("const monthsFullList = ['JAN','FEB'"));

if (startIdx !== -1 && endIdx !== -1) {
    // Delete from startIdx to endIdx inclusive
    lines.splice(startIdx, endIdx - startIdx + 1);
    fs.writeFileSync(path, lines.join('\n'));
    console.log('Successfully removed unused variables and plotting logic.');
} else {
    console.log('Could not find the section to remove.');
}
