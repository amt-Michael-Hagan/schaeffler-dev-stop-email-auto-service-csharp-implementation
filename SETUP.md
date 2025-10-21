# C# Email Automation Service - .NET Framework 4.0 Setup Guide

## Overview

This is a complete C# rewrite of the Node.js email automation service, designed to work with .NET Framework 4.0 for maximum compatibility with legacy systems.

## Features

✅ **Fully Automated** - No user interaction required  
✅ **Application Permissions** - Uses Client Credentials flow  
✅ **Legacy Compatible** - Works with .NET Framework 4.0  
✅ **Manual Graph API** - Direct HTTP calls to Microsoft Graph  
✅ **Attachment Download** - Automatically saves to disk  
✅ **Duplicate Prevention** - Smart tracking prevents re-downloads  
✅ **Comprehensive Logging** - Daily log files with detailed information  

## Prerequisites

- **.NET Framework 4.0** or higher
- **Visual Studio 2010** or higher (or MSBuild)
- **NuGet Package Manager** (comes with VS 2012+, available separately for VS 2010)
- **Microsoft 365** account with admin access

## Project Structure

```
Legacy/
├── EmailAutomationLegacy.sln          # Visual Studio Solution
├── EmailAutomationLegacy.csproj       # Project file
├── App.config                         # Configuration & logging setup
├── packages.config                    # NuGet dependencies
├── Program.cs                         # Main entry point
├── AppSettings.cs                     # Configuration helper
├── Services/
│   ├── TokenManager.cs               # OAuth token management
│   ├── GraphApiClient.cs             # HTTP client for Graph API
│   └── EmailProcessor.cs             # Email processing logic
├── Models/
│   └── GraphModels.cs                # Data models for Graph API
└── Properties/
    └── AssemblyInfo.cs               # Assembly metadata
```

## Setup Steps

### 1. Build the Project

#### Option A: Using Visual Studio
1. Open `EmailAutomationLegacy.sln` in Visual Studio
2. Right-click on the solution → "Enable NuGet Package Restore"
3. Build → Build Solution (Ctrl+Shift+B)

#### Option B: Using MSBuild (Command Line)
```bash
# Navigate to the Legacy directory
cd CSharp-Implementation\Legacy

# Restore NuGet packages (VS 2012+ or standalone NuGet.exe)
nuget restore

# Build the project
msbuild EmailAutomationLegacy.sln /p:Configuration=Release
```

### 2. Configure Azure AD (Same as Node.js version)

Follow the same Azure AD setup from the original `AUTOMATION_SETUP.md`:

1. **Create App Registration** in Azure Portal
2. **Add Application Permissions**: `Mail.Read`, `Mail.ReadWrite`, `User.Read.All`
3. **Grant Admin Consent**
4. **Create Client Secret**

### 3. Update Configuration

Edit `App.config`:

```xml
<appSettings>
  <!-- From Azure Portal -->
  <add key="ClientId" value="your-client-id-here" />
  <add key="TenantId" value="your-tenant-id-here" />
  <add key="ClientSecret" value="your-client-secret-here" />
  
  <!-- Target email account -->
  <add key="TargetEmail" value="user@yourdomain.com" />
  
  <!-- Optional: Customize processing settings -->
  <add key="HoursToFetch" value="24" />
  <add key="DownloadsDirectory" value="./downloads" />
</appSettings>
```

### 4. Run the Application

```bash
# From the bin\Release or bin\Debug directory
EmailAutomationLegacy.exe
```

## Dependencies (.NET Framework 4.0 Compatible)

- **Newtonsoft.Json 6.0.8** - JSON serialization (last version supporting .NET 4.0)
- **log4net 2.0.8** - Logging framework
- **System.Net.Http** - HTTP client (via NuGet for .NET 4.0)

## Key Differences from Node.js Version

### Authentication
- **Node.js**: Uses `@azure/identity` library
- **C#**: Manual OAuth2 Client Credentials implementation with HttpWebRequest

### HTTP Calls
- **Node.js**: Uses `@microsoft/microsoft-graph-client`
- **C#**: Direct HTTP calls to Graph API endpoints using HttpWebRequest

### Configuration
- **Node.js**: Uses `.env` file with `dotenv`
- **C#**: Uses `App.config` with `ConfigurationManager`

### Logging
- **Node.js**: Console + file logging with custom implementation
- **C#**: log4net with console and rolling file appenders

## Example Output

```
🤖 C# Email Automation Service (.NET Framework 4.0)
============================================================
🔍 Testing connection...
[08:00:01] [INFO] Testing Microsoft Graph API connection...
[08:00:02] [INFO] Successfully obtained access token, expires at: 10/20/2025 9:00:02 AM
[08:00:03] [INFO] Successfully connected. User: John Doe (john@company.com)
[08:00:04] [INFO] Mailbox access confirmed. Found 1 messages in test query
✅ Connection successful!
📧 Starting email processing...
[08:00:05] [INFO] Starting email processing...
[08:00:05] [INFO] Target mailbox: john@company.com
[08:00:05] [INFO] Fetching emails from last 24 hours
[08:00:06] [INFO] Loaded tracking data: 15 attachments previously processed
[08:00:07] [INFO] Found 8 emails with attachments
[08:00:07] [INFO] Filtered out 2 sent emails (processing only incoming emails)

--- Email 1/6 ---
[08:00:08] [INFO] Subject: Invoice October 2025
[08:00:08] [INFO] From: vendor@supplier.com
[08:00:08] [INFO] Date: 2025-10-20 07:30:00
[08:00:09] [INFO]   Found 2 attachment(s)
[08:00:10] [INFO]     ✅ Downloaded: 2025-10-20_invoice.pdf (245.67 KB)
[08:00:10] [INFO]     ⏭️  Skipped (already downloaded): receipt.pdf

============================================================
📊 Processing Complete
============================================================
Emails processed: 6
Attachments found: 12
New downloads: 8
Already processed: 4
Downloads directory: C:\Path\To\Your\Project\downloads

✅ Service completed successfully!
```

## Troubleshooting

### "Could not load file or assembly 'Newtonsoft.Json'"
- Ensure NuGet packages are restored
- Check that `packages.config` versions match your .NET Framework version

### "The request was aborted: Could not create SSL/TLS secure channel"
- This can happen with .NET 4.0 and modern TLS requirements
- Add this to your `App.config` in `<appSettings>`:
  ```xml
  <add key="Switch.System.Net.DontEnableSchUseStrongCrypto" value="false" />
  ```

### "Authentication failed: Unauthorized"
- Verify your Client ID, Tenant ID, and Client Secret
- Ensure admin consent was granted for Application permissions
- Check that the target email account exists in your tenant

### Visual Studio 2010 NuGet Issues
- Install NuGet Package Manager extension for VS 2010
- Or manually download packages and reference DLLs

## Advanced Configuration

### Custom Retry Logic
```xml
<add key="RetryAttempts" value="5" />
<add key="RetryDelayMs" value="3000" />
```

### Logging Levels
In `App.config`, change the log level:
```xml
<root>
  <level value="DEBUG" />  <!-- INFO, WARN, ERROR, DEBUG -->
  ...
</root>
```

### Download Directory Structure
```xml
<add key="DownloadsDirectory" value="C:\EmailAttachments" />
```

## Production Deployment

1. **Build in Release mode** for better performance
2. **Copy the entire bin\Release folder** to target machine
3. **Install .NET Framework 4.0** on target machine if not present
4. **Configure App.config** with production values
5. **Set up scheduled task** to run the executable periodically

## Migration Notes from Node.js

The C# version maintains **100% functional parity** with the Node.js service:

- ✅ Same Azure AD configuration
- ✅ Same email filtering logic
- ✅ Same attachment tracking format
- ✅ Same duplicate prevention
- ✅ Same logging structure
- ✅ Same configuration options

You can run both versions side-by-side without conflicts, as they use the same tracking file format.