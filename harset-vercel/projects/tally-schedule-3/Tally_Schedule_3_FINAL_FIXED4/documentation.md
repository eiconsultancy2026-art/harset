# Tally_Schedule_3 - Application Documentation

This document explains how to set up, run, and manage the Tally_Schedule_3 application.

## Prerequisites
- **Node.js**: Make sure you have Node.js installed on your Windows machine.
- **TallyPrime / Tally ERP 9**: Ensure that Tally is open and configured to allow API requests if this application fetches data locally.

## Setup
1. Open a terminal (PowerShell or Command Prompt).
2. Navigate to the project directory:
   ```powershell
   cd C:\Kevin\ei\Tally_Schedule_3
   ```
3. Install the required Node.js dependencies:
   ```powershell
   npm install
   ```

## Running the Application Locally
To start the express server for development or local usage, run:
```powershell
npm start
```
Alternatively, you can run the server directly using Node:
```powershell
node server.js
```
The server will start and be ready to process requests.

## Running the Mapping Tests
To test the mappings (Note 1 to Note 18) and view the output directly in your console, you can run the provided test script:
```powershell
node test_mapping.js
```
This will print the extracted details for items such as Share Capital, Secured/Unsecured Loans, Trade Payables, etc.

## Running in Production (Background Service with PM2)
If you want to run the server continuously in the background and have it restart automatically if the system reboots, you should use **PM2**, a process manager for Node.js.

### 1. Install PM2
Install PM2 globally on your machine:
```powershell
npm install -g pm2
```

### 2. Start the Application
Start the server and name the process "TallyMapper":
```powershell
pm2 start server.js --name "TallyMapper"
```

### 3. Setup Auto-start on Boot
To ensure the application starts automatically when Windows boots up:
```powershell
npm install -g pm2-windows-startup
pm2-startup install
pm2 save
```

### PM2 Useful Commands
- **View running processes**:
  ```powershell
  pm2 list
  ```
- **View live logs**:
  ```powershell
  pm2 logs TallyMapper
  ```
- **Stop the application**:
  ```powershell
  pm2 stop TallyMapper
  ```
- **Restart the application**:
  ```powershell
  pm2 restart TallyMapper
  ```
- **Remove the application from PM2**:
  ```powershell
  pm2 delete TallyMapper
  ```
