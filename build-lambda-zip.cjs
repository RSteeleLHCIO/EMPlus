#!/usr/bin/env node

const { execSync } = require('child_process');
const fs = require('fs');

console.log('Building lambda deployment package...\n');

try {
  // Remove existing zip
  if (fs.existsSync('lambda-deploy.zip')) {
    console.log('Removing old lambda-deploy.zip...');
    fs.unlinkSync('lambda-deploy.zip');
  }

  // Create the zip file using PowerShell
  console.log('Creating zip archive...');
  execSync('powershell -NoProfile -Command "Compress-Archive -Path package.json,server,node_modules -DestinationPath lambda-deploy.zip -Force"', {
    stdio: 'inherit'
  });

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
