const fs = require('fs');
const path = require('path');

function removePS1Comments(content) {
  const regex = /((?:"(?:\\.|[^"])*?"|'(?:\\.|[^'])*?'))|(<#[\s\S]*?#>|#[^\n]*)/g;
  let count = 0;
  let result = content.replace(regex, (match, p1, p2) => {
    count++;
    if (p1 === undefined) {
      console.log(`Match ${count} (COMMENT): match=${JSON.stringify(match)}, p2=${JSON.stringify(p2)}`);
    }
    if (p1 !== undefined) return p1;
    return '';
  });
  console.log(`Total matches: ${count}`);
  return result;
}

const fullPath = 'C:\\MICROSOFT_OUTLOOK_SECURITY_CENTER\\outlook-scanner.ps1';
let content = fs.readFileSync(fullPath, 'utf8');
console.log('Original length:', content.length);
let newContent = removePS1Comments(content);
console.log('New length:', newContent.length);

if (content !== newContent) {
    fs.writeFileSync(fullPath, newContent, 'utf8');
    console.log('Comments removed.');
} else {
    console.log('No change.');
}
