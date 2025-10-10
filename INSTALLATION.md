# 🚀 Quick Installation Guide

## Prerequisites Checklist

Before you begin, ensure you have:

- [ ] **Windows computer** with PowerShell 5.1 or later
- [ ] **Microsoft 365 Global Administrator** account access
- [ ] **Internet connection** for API access
- [ ] **Administrator privileges** on your computer

## Step-by-Step Installation

### 1. Download the Tool
- Download all files to a folder (e.g., `C:\HYCU-M365Sizer\`)
- Ensure all `.ps1` files and the `web-interface` folder are included

### 2. Install PowerShell Modules
Open **PowerShell as Administrator** and run:

```powershell
# Install required modules
Install-Module Microsoft.Graph.Reports -Force -AllowClobber
Install-Module Microsoft.Graph.Users -Force -AllowClobber  
Install-Module Microsoft.Graph.Groups -Force -AllowClobber
Install-Module Microsoft.Graph.Teams -Force -AllowClobber
Install-Module Microsoft.Graph.Sites -Force -AllowClobber
Install-Module ExchangeOnlineManagement -Force -AllowClobber
```

### 3. Verify Installation
```powershell
# Check if modules are installed
Get-Module Microsoft.Graph.* -ListAvailable
```

You should see all 5 Microsoft.Graph modules listed.

### 4. Test the Tool
```powershell
# Navigate to the tool folder
cd C:\HYCU-M365Sizer

# Run the tool
.\Get-HYCUM365SizingInfo.ps1
```

### 5. Sign In and Grant Permissions
- A browser window will open
- Sign in with your **Global Administrator** account
- Click **"Accept"** to grant the required permissions
- Return to PowerShell and wait for completion

## ✅ Success Indicators

You'll know the installation worked when:
- [ ] PowerShell modules install without errors
- [ ] Browser opens for authentication
- [ ] You can sign in successfully
- [ ] Analysis begins and shows progress messages
- [ ] An HTML report is generated

## 🚨 Common Issues & Solutions

### "Module not found" Error
**Solution:** Run PowerShell as Administrator and reinstall modules
```powershell
Install-Module Microsoft.Graph.Reports -Force -AllowClobber
```

### "Access Denied" Error  
**Solution:** Ensure you're using a Global Administrator account

### "Authentication Failed" Error
**Solution:** 
- Clear browser cache and cookies
- Try a different browser
- Ensure MFA is properly configured

### "No Data Returned" Error
**Solution:**
- Verify your tenant has been active for 30+ days
- Check that you have Microsoft 365 licenses assigned
- Ensure your account has the required permissions

## 📞 Need Help?

If you encounter issues:
1. **Check the main README.md** for detailed troubleshooting
2. **Review error messages** in the PowerShell window
3. **Verify your permissions** with your IT administrator
4. **Test with a different account** if possible

## 🎯 Next Steps

Once installation is complete:
1. **Run the analysis** following the main README guide
2. **Review your report** in the generated HTML file
3. **Use the web interface** for interactive analysis
4. **Share results** with your team for backup planning

---

**Ready to analyze your Microsoft 365 tenant?** Follow the main README.md for detailed usage instructions!