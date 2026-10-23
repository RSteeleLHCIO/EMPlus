#!/usr/bin/env node

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('Building lambda deployment package...\n');

try {
  // Remove existing zip
  if (fs.existsSync('lambda-deploy.zip')) {
    console.log('Removing old lambda-deploy.zip...');
    fs.unlinkSync('lambda-deploy.zip');
  }

  // Stamp the server with its version/commit/build time (served by GET /api/version)
  const stampFile = path.join('server', 'buildInfo.json');
  const stamp = execSync(
    'node --input-type=module -e "import(\'./server/buildInfo.js\').then(m => process.stdout.write(JSON.stringify(m.computeBuildInfo())))"'
  ).toString();
  fs.writeFileSync(stampFile, stamp);
  console.log(`Build info: ${stamp}`);

  // Create the zip file using PowerShell
  console.log('Creating zip archive...');
  try {
    execSync('powershell -NoProfile -Command "Compress-Archive -Path package.json,server,node_modules -DestinationPath lambda-deploy.zip -Force"', {
      stdio: 'inherit'
    });
  } finally {
    fs.unlinkSync(stampFile);
  }

  // Get file size
  if (fs.existsSync('lambda-deploy.zip')) {
    const stats = fs.statSync('lambda-deploy.zip');
    const sizeMB = (stats.size / 1048576).toFixed(1);
    console.log(`\n✓ lambda-deploy.zip ready. Size: ${sizeMB} MB`);
    process.exit(0);
  } else {
    throw new Error('lambda-deploy.zip was not created');
  }
} catch (error) {
  console.error('\n✗ Error building lambda package:', error.message);
  process.exit(1);
}
