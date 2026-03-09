const fs = require('fs');
const path = require('path');

function removeJSComments(content) {
  const regex = /((?:"(?:\\.|[^"])*?"|'(?:\\.|[^'])*?'|`(?:\\.|[^`])*?`|(?:\/(?:\\.|[^\/])+\/[gimuy]*)))|(\/\/[^\n]*|\/\*[\s\S]*?\*\/)/g;
  return content.replace(regex, (match, p1, p2) => {
    if (p1 !== undefined) return p1;
    return '';
  });
}

function removePS1Comments(content) {
  const regex = /((?:"(?:\\.|[^"])*?"|'(?:\\.|[^'])*?'))|(<#[\s\S]*?#>|#[^\n]*)/g;
  return content.replace(regex, (match, p1, p2) => {
    if (p1 !== undefined) return p1;
    return '';
  });
}

function removeHTMLComments(content) {
  let newContent = content.replace(/<!--[\s\S]*?-->/g, '');
  newContent = newContent.replace(/(<script[\s\S]*?>)([\s\S]*?)(<\/script>)/gi, (match, start, script, end) => {
    return start + removeJSComments(script) + end;
  });
  newContent = newContent.replace(/(<style[\s\S]*?>)([\s\S]*?)(<\/style>)/gi, (match, start, style, end) => {
    return start + style.replace(/\/\*[\s\S]*?\*\//g, '') + end;
  });
  return newContent;
}

function removeVBSComments(content) {
  const regex = /((?:"(?:\\.|[^"])*?"))|('[^\n]*|\bRem\b[^\n]*)/gi;
  return content.replace(regex, (match, p1, p2) => {
    if (p1 !== undefined) return p1;
    return '';
  });
}

function removeBATComments(content) {
  return content.replace(/(^|\n)\s*(REM\s[^\n]*|::[^\n]*)/gi, '$1');
}

const files = [
  { path: 'main.js', type: 'js' },
  { path: 'outlook-scanner.ps1', type: 'ps1' },
  { path: 'preload.js', type: 'js' },
  { path: 'index.html', type: 'html' },
  { path: 'MICROSOFT_OUTLOOK_SECURITY_CENTER.bat', type: 'bat' },
  { path: 'fix_all.js', type: 'js' },
  { path: 'build.bat', type: 'bat' },
  { path: 'silent_launcher.vbs', type: 'vbs' }
];

files.forEach(fileInfo => {
  const fullPath = path.join('C:\\MICROSOFT_OUTLOOK_SECURITY_CENTER', fileInfo.path);
  if (!fs.existsSync(fullPath)) return;
  
  let content = fs.readFileSync(fullPath, 'utf8');
  let originalContent = content;
  
  if (fileInfo.path === 'outlook-scanner.ps1') {
    console.log(`Processing outlook-scanner.ps1, length: ${content.length}`);
    const hasComment = content.includes('#');
    console.log(`Has '#' character: ${hasComment}`);
  }
  
  switch (fileInfo.type) {
    case 'js': content = removeJSComments(content); break;
    case 'ps1': content = removePS1Comments(content); break;
    case 'html': content = removeHTMLComments(content); break;
    case 'vbs': content = removeVBSComments(content); break;
    case 'bat': content = removeBATComments(content); break;
  }
  
  if (content !== originalContent) {
    fs.writeFileSync(fullPath, content, 'utf8');
    console.log(`Comments removed from: ${fileInfo.path}`);
  } else {
    console.log(`No comments to remove in: ${fileInfo.path}`);
  }
});
