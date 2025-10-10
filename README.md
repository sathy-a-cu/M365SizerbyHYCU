# 🚀 HYCU M365 Sizing Tool

A comprehensive Microsoft 365 tenant analysis tool designed for backup planning, capacity management, and cost optimization. This tool provides detailed insights into your M365 environment to help you make informed decisions about backup strategies and resource allocation.

## ✨ What This Tool Does

This tool analyzes your Microsoft 365 tenant and generates comprehensive reports showing:
- **Storage usage** across Exchange, OneDrive, and SharePoint
- **User statistics** and licensing information
- **Cost estimates** for HYCU backup services
- **Top 5 largest** mailboxes, OneDrive accounts, and SharePoint sites
- **Growth projections** and capacity planning
- **Teams collaboration** metrics and costs

## 🚀 Quick Start Guide

### Step 1: Prerequisites

**Required Software:**
- Windows PowerShell 5.1 or later
- Internet connection for Microsoft Graph API access
- Microsoft 365 Global Administrator or equivalent permissions

**Required Permissions:**
Your account needs these Microsoft Graph permissions:
- `Reports.Read.All` - Access usage reports
- `User.Read.All` - Read user information
- `Group.Read.All` - Read group information
- `Team.ReadBasic.All` - Read Teams information
- `Sites.Read.All` - Read SharePoint information

### Step 2: Download and Setup

1. **Download the tool:**
   - Download all files to a folder on your computer (e.g., `C:\HYCU-M365Sizer\`)

2. **Install PowerShell modules:**
   Open PowerShell as Administrator and run:
   ```powershell
   Install-Module Microsoft.Graph.Reports, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Teams, Microsoft.Graph.Sites, ExchangeOnlineManagement -Force -AllowClobber
   ```

3. **Verify installation:**
   ```powershell
   Get-Module Microsoft.Graph.* -ListAvailable
   ```

### Step 3: Run the Analysis

1. **Open PowerShell** (as Administrator)

2. **Navigate to the tool folder:**
   ```powershell
   cd C:\HYCU-M365Sizer
   ```

3. **Run the main analysis:**
   ```powershell
   .\Get-HYCUM365SizingInfo.ps1
   ```

4. **Sign in when prompted:**
   - A browser window will open
   - Sign in with your Microsoft 365 Global Administrator account
   - Grant the requested permissions

5. **Wait for completion:**
   - The script will analyze your tenant (this may take 5-15 minutes)
   - You'll see progress messages in the PowerShell window
   - A report will be generated when complete

### Step 4: View Your Report

1. **Find your report:**
   - Look for a file named `HYCU-M365-Sizing-YYYY-MM-DD-HHMM.html`
   - This will be in the same folder as the script

2. **Open the report:**
   - Double-click the HTML file to open it in your web browser
   - The report contains all your tenant analysis data

## 📊 Understanding Your Report

### 📈 Tenant Overview
- Total users, active users, and guest users
- Basic tenant information and statistics

### 💾 Storage Analysis
- Exchange Online storage usage
- OneDrive for Business storage
- SharePoint Online storage
- Total storage across all services

### 📧 Mailbox Analysis
- Total mailboxes (active, shared, archive, resource)
- Archive mailbox percentage and thresholds
- Mailbox type breakdown

### 🌐 Sites & OneDrive Analysis
- OneDrive account counts
- SharePoint site counts
- Teams site counts
- Total site statistics

### 🏆 Top 5 by Size
- Largest mailboxes by storage size
- Largest OneDrive accounts
- Largest SharePoint sites
- Helps identify data-heavy users and sites

### 📋 HYCU Licensing Analysis
- Microsoft 365 license distribution
- HYCU entitlement calculations (50 GB per user)
- Additional license requirements
- Archive mailbox analysis

### 💰 Initial Cost Estimates
- Monthly and annual cost projections
- Per-user cost breakdown
- Storage and worker node costs
- Cost optimization insights

## 🔧 Advanced Usage

### Using App Authentication (For Automation)

If you want to run this without interactive login:

1. **Create an Azure AD App Registration:**
   - Go to Azure Portal > Azure Active Directory > App registrations
   - Click "New registration"
   - Name: "HYCU M365 Sizing Tool"
   - Supported account types: "Accounts in this organizational directory only"
   - Click "Register"

2. **Configure API Permissions:**
   - Go to "API permissions" in your app
   - Click "Add a permission" > "Microsoft Graph" > "Application permissions"
   - Add these permissions:
     - `Reports.Read.All`
     - `User.Read.All`
     - `Group.Read.All`
     - `Team.ReadBasic.All`
     - `Sites.Read.All`
   - Click "Grant admin consent"

3. **Create a Client Secret:**
   - Go to "Certificates & secrets"
   - Click "New client secret"
   - Add description: "HYCU M365 Sizing Tool"
   - Choose expiration (recommend 24 months)
   - Click "Add"
   - **Copy the secret value immediately** (you won't see it again)

4. **Run with app authentication:**
   ```powershell
   .\Get-HYCUM365SizingInfo.ps1 -UseAppAccess $true -TenantId "your-tenant-id" -ClientId "your-app-id" -ClientSecret "your-secret"
   ```

### Customizing the Analysis

**Skip certain analyses:**
```powershell
.\Get-HYCUM365SizingInfo.ps1 -SkipArchiveMailbox $true -SkipRecoverableItems $true
```

**Custom growth rate:**
```powershell
.\Get-HYCUM365SizingInfo.ps1 -AnnualGrowth 25
```

**Custom output location:**
```powershell
.\Get-HYCUM365SizingInfo.ps1 -OutputPath "C:\Reports"
```

**Filter to specific group:**
```powershell
.\Get-HYCUM365SizingInfo.ps1 -ADGroup "Sales Team"
```

## 🛠️ Troubleshooting

### Common Issues

**"Module not found" error:**
```powershell
# Install missing modules
Install-Module Microsoft.Graph.Reports -Force -AllowClobber
```

**"Access denied" error:**
- Ensure you're using a Global Administrator account
- Check that your account has the required permissions
- Try running PowerShell as Administrator

**"Authentication failed" error:**
- Clear browser cache and cookies
- Try a different browser
- Ensure your account has MFA configured properly

**"No data returned" error:**
- Check your Microsoft 365 license (some features require specific licenses)
- Ensure your tenant has been active for at least 30 days
- Verify API permissions are granted

### Getting Help

**Check the script output:**
- The script provides detailed progress messages
- Look for error messages in red text
- Check the final summary for any warnings

**Verify permissions:**
```powershell
# Check your current permissions
Get-MgContext
```

**Test connectivity:**
```powershell
# Test Microsoft Graph connection
Connect-MgGraph -Scopes "User.Read.All"
Get-MgUser -Top 1
```

## 📁 Output Files

After running the script, you'll find:

- **`HYCU-M365-Sizing-YYYY-MM-DD-HHMM.html`** - Main report file
- **Console output** - Progress messages and summary
- **Error logs** (if any) - Detailed error information

## 🔐 Security Notes

- **No data is stored externally** - All analysis happens locally
- **Uses Microsoft's official APIs** - Secure and compliant
- **Minimal permissions** - Only requests what's needed
- **Your data stays private** - No data sent to external services

## 📊 Sample Reports

Check the `sample-outputs` folder for example reports showing:
- Complete tenant analysis
- Cost breakdown examples
- Licensing analysis samples
- Top 5 analysis examples

## 🎯 Key Benefits

### For Backup Planning
- **Accurate sizing** for backup storage requirements
- **Growth projections** for capacity planning
- **Cost estimates** for backup services
- **Data distribution** analysis

### For Cost Optimization
- **Per-user cost breakdown**
- **Storage optimization** opportunities
- **License optimization** recommendations
- **Archive mailbox** analysis

### For Capacity Management
- **Current usage** vs. **entitlements**
- **Growth trends** and projections
- **Top consumers** identification
- **Capacity planning** insights

## 🔄 Version Information

**Current Version:** 1.3
- Comprehensive M365 analysis
- HYCU-specific licensing model
- Cost optimization features
- Modern web interface
- Top 5 analysis
- Teams cost analysis

## 📞 Support

**Self-Help Resources:**
- This README contains most common solutions
- Check the sample outputs for reference
- Review PowerShell error messages for specific issues

**Common Solutions:**
- **Permission issues:** Ensure Global Admin access
- **Module issues:** Reinstall PowerShell modules
- **Authentication issues:** Clear browser cache and retry
- **Data issues:** Verify tenant has sufficient data history

---

**Made with ❤️ by HYCU** - Empowering your Microsoft 365 backup strategy

*This tool helps you understand your Microsoft 365 environment to make informed decisions about backup planning, cost optimization, and capacity management.*