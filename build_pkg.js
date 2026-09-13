const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const root = __dirname;
const staging = path.join(root, 'dist_tmp');
const distApp = path.join(root, 'dist', 'DeskGuard for Microsoft Outlook-win32-x64');
const stagedApp = path.join(staging, 'DeskGuard for Microsoft Outlook-win32-x64');

console.log('[DeskGuard Packager] Generating distribution binaries...');

// 1. Clean any stale staging dir
if (fs.existsSync(staging)) {
    try { fs.rmSync(staging, { recursive: true, force: true }); } catch (e) {}
}

// 2. Run electron-packager into staging
const packCmd = `npx electron-packager . "DeskGuard for Microsoft Outlook" --platform=win32 --arch=x64 --out=dist_tmp --icon=icon.ico --ignore="^/dist" --ignore="^/dist_tmp"`;
execSync(packCmd, { cwd: root, stdio: 'inherit' });

// 3. Terminate running instance if any, and mirror into dist using Windows robocopy
console.log('[DeskGuard Packager] Closing active instances and mirroring package into dist\\...');
try {
    execSync('taskkill /F /IM "DeskGuard for Microsoft Outlook.exe" /T', { stdio: 'ignore' });
} catch (e) {}

if (!fs.existsSync(path.dirname(distApp))) {
    fs.mkdirSync(path.dirname(distApp), { recursive: true });
}

try {
    // In robocopy, exit codes 0-7 indicate successful copy operations
    execSync(`robocopy "${stagedApp}" "${distApp}" /MIR /NP /NFL /NDL /NJH /NJS /R:2 /W:1`, { stdio: 'inherit' });
} catch (err) {
    if (err.status && err.status > 7) {
        throw new Error(`Robocopy failed with exit code ${err.status}`);
    }
}

// 4. Cleanup staging
try {
    if (fs.existsSync(staging)) {
        fs.rmSync(staging, { recursive: true, force: true });
    }
} catch (e) {}

console.log('[DeskGuard Packager] Binary packaging complete: dist\\');
