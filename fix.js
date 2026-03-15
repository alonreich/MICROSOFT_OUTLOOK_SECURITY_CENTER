const fs = require('fs');
let code = fs.readFileSync('main.js', 'utf8');
code = code.replace(/catch\s*\{\}/g, 'catch (err) { if(err && err.message) { console.error(err); logToFile("Handled Exception: " + err.message, "ERROR"); } }');
fs.writeFileSync('main.js', code);
console.log('Done!');
