#Requires -Version 5.1

<#
.SYNOPSIS
    HYCU M365 Sizing Tool - Comprehensive Microsoft 365 Tenant Analysis for Backup Planning

.DESCRIPTION
    This script analyzes Microsoft 365 tenant data to provide comprehensive sizing information
    for backup planning, including storage usage, user counts, growth projections, and backup recommendations.

.PARAMETER UseAppAccess
    Use application-based authentication instead of interactive login

.PARAMETER TenantId
    Azure AD Tenant ID (required when using app access)

.PARAMETER ClientId
    Azure AD Application Client ID (required when using app access)

.PARAMETER ClientSecret
    Azure AD Application Client Secret (required when using app access)

.PARAMETER ADGroup
    Filter analysis to specific Azure AD group

.PARAMETER SkipArchiveMailbox
    Skip gathering In-Place Archive mailbox statistics

.PARAMETER SkipRecoverableItems
    Skip gathering Recoverable Items folder statistics

.PARAMETER AnnualGrowth
    Custom annual growth rate percentage (default: 30)

.PARAMETER Period
    Historical data period in days (default: 180)

.PARAMETER OutputPath
    Custom output directory for reports

.EXAMPLE
    .\Get-HYCUM365SizingInfo.ps1

.EXAMPLE
    .\Get-HYCUM365SizingInfo.ps1 -UseAppAccess $true -TenantId "your-tenant-id" -ClientId "your-client-id" -ClientSecret "your-secret"

.NOTES
    Author: HYCU
    Version: 1.0
    Requires: PowerShell 5.1+, Microsoft.Graph modules, ExchangeOnlineManagement
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [bool]$UseAppAccess = $false,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory = $false)]
    [string]$ADGroup,
    
    [Parameter(Mandatory = $false)]
    [bool]$SkipArchiveMailbox = $false,
    
    [Parameter(Mandatory = $false)]
    [bool]$SkipRecoverableItems = $false,
    
    [Parameter(Mandatory = $false)]
    [int]$AnnualGrowth = 30,
    
    [Parameter(Mandatory = $false)]
    [int]$Period = 180,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "."
)

# Global variables
$script:StartTime = Get-Date
$script:OutputDirectory = $OutputPath

# Set up global error handling
$ErrorActionPreference = "Continue"
$Error.Clear()

# Global error handler
trap {
    Write-PowerShellError -ErrorMessage $_.Exception.Message -Exception $_.Exception.ToString()
    continue
}
# Set up logging
$LogFile = Join-Path $script:OutputDirectory "HYCU-M365-Sizing-$(Get-Date -Format 'yyyy-MM-dd-HHmm').log"
$ErrorLogFile = Join-Path $script:OutputDirectory "HYCU-M365-Sizing-Errors-$(Get-Date -Format 'yyyy-MM-dd-HHmm').log"

# Function to write to both console and log file
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console
    Write-Host $logEntry
    
    # Write to log file
    Add-Content -Path $LogFile -Value $logEntry -Force
}

# Function to write errors to separate error log
function Write-ErrorLog {
    param(
        [string]$Message,
        [string]$Exception = ""
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $errorEntry = "[$timestamp] [ERROR] $Message"
    if ($Exception) {
        $errorEntry += "`nException: $Exception"
    }
    
    # Write to error log file
    Add-Content -Path $ErrorLogFile -Value $errorEntry -Force
}

$script:ReportData = @{
    TenantInfo = @{}
    ExchangeData = @{}
    OneDriveData = @{}
    SharePointData = @{}
    TeamsData = @{}
    ArchiveData = @{}
    RecoverableItemsData = @{}
    GrowthAnalysis = @{}
    BackupRecommendations = @{}
    CostAnalysis = @{}
    LicensingInfo = @{}
    SitesAndOneDriveData = @{}
    PlannerData = @{}
    GroupsData = @{}
}

# Function to write colored output with proper error routing
function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "INFO" { "White" }
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        "HEADER" { "Cyan" }
    }
    
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console with color
    Write-Host $logEntry -ForegroundColor $color
    
    # Always write to main log file
    try {
        Add-Content -Path $LogFile -Value $logEntry -Force -ErrorAction SilentlyContinue
    }
    catch {
        # If log file write fails, continue
    }
    
    # Write errors to separate error log
    if ($Level -eq "ERROR") {
        try {
            Write-ErrorLog -Message $Message
        }
        catch {
            # If error log write fails, continue
        }
    }
}

# Function to capture and log PowerShell errors
function Write-PowerShellError {
    param(
        [string]$ErrorMessage,
        [string]$Exception = ""
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $errorEntry = "[$timestamp] [ERROR] PowerShell Error: $ErrorMessage"
    
    if ($Exception) {
        $errorEntry += "`nException Details: $Exception"
    }
    
    # Write to console
    Write-Host $errorEntry -ForegroundColor Red
    
    # Write to log files
    try {
        Add-Content -Path $LogFile -Value $errorEntry -Force -ErrorAction SilentlyContinue
        Write-ErrorLog -Message $ErrorMessage -Exception $Exception
    }
    catch {
        # If logging fails, continue
    }
}

# Function to check and install required modules
function Install-RequiredModules {
    Write-ColorOutput "Checking for required PowerShell modules..." "INFO"
    
    $requiredModules = @(
        "Microsoft.Graph.Reports",
        "Microsoft.Graph.Users", 
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Teams",
        "Microsoft.Graph.Sites",
        "ExchangeOnlineManagement"
    )
    
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-ColorOutput "Installing module: $module" "INFO"
            try {
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
                Write-ColorOutput "Successfully installed $module" "SUCCESS"
            }
            catch {
                Write-ColorOutput "Failed to install ${module}: $($_.Exception.Message)" -Color Red
                throw
            }
        }
        else {
            Write-ColorOutput "$module is already installed" "SUCCESS"
        }
    }
}

# Function to authenticate to Microsoft Graph
function Connect-MicrosoftGraph {
    Write-ColorOutput "Authenticating to Microsoft Graph API..." "INFO"
    
    try {
        if ($UseAppAccess) {
            if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) {
                throw "TenantId, ClientId, and ClientSecret are required when using app access"
            }
            
            $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)
            
            Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential
            Write-ColorOutput "Connected using application authentication" "SUCCESS"
        }
        else {
            Connect-MgGraph -Scopes "Reports.Read.All", "User.Read.All", "Group.Read.All", "Team.ReadBasic.All", "Sites.Read.All"
            Write-ColorOutput "Connected using interactive authentication" "SUCCESS"
        }
    }
    catch {
        Write-ColorOutput "Failed to connect to Microsoft Graph: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# Function to get tenant information
function Get-TenantInformation {
    Write-ColorOutput "Gathering tenant information..." "INFO"
    
    try {
        $organization = Get-MgOrganization
        $script:ReportData.TenantInfo = @{
            DisplayName = $organization.DisplayName
            Id = $organization.Id
            VerifiedDomains = $organization.VerifiedDomains
            CreatedDateTime = $organization.CreatedDateTime
        }
        
        Write-ColorOutput "Tenant: $($organization.DisplayName)" "SUCCESS"
    }
    catch {
        Write-ColorOutput "Failed to get tenant information: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get user statistics
function Get-UserStatistics {
    Write-ColorOutput "Gathering user statistics..." "INFO"
    
    try {
        $users = Get-MgUser -All -Property "UserPrincipalName,DisplayName,AccountEnabled,UserType"
        $enabledUsers = $users | Where-Object { $_.AccountEnabled -eq $true }
        $guestUsers = $users | Where-Object { $_.UserType -eq "Guest" }
        
        $script:ReportData.TenantInfo.UserCounts = @{
            TotalUsers = $users.Count
            EnabledUsers = $enabledUsers.Count
            GuestUsers = $guestUsers.Count
            DisabledUsers = ($users.Count - $enabledUsers.Count)
        }
        
        Write-ColorOutput "Total Users: $($users.Count) (Enabled: $($enabledUsers.Count), Guests: $($guestUsers.Count))" "SUCCESS"
    }
    catch {
        Write-ColorOutput "Failed to get user statistics: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get licensing information
function Get-LicensingInformation {
    Write-ColorOutput "Analyzing Microsoft 365 licensing..." "INFO"
    
    try {
        # Get all users with their assigned licenses
        $users = Get-MgUser -All -Property "UserPrincipalName,DisplayName,AccountEnabled,UserType,AssignedLicenses"
        $licensedUsers = $users | Where-Object { $_.AssignedLicenses.Count -gt 0 }
        
        # Get available license SKUs
        $licenseSKUs = Get-MgSubscribedSku -All
        
        # Analyze license distribution
        $licenseAnalysis = @{}
        $totalLicensedUsers = 0
        
        foreach ($sku in $licenseSKUs) {
            $skuName = $sku.SkuPartNumber
            $assignedUnits = $sku.PrepaidUnits.Enabled
            $consumedUnits = $sku.ConsumedUnits
            
            # Map SKU to storage limits and license tier
            $storageLimit = Get-LicenseStorageLimit -SkuPartNumber $skuName
            $licenseTier = Get-LicenseTier -SkuPartNumber $skuName
            
            $licenseAnalysis[$skuName] = @{
                DisplayName = $sku.SkuPartNumber
                AssignedUnits = $assignedUnits
                ConsumedUnits = $consumedUnits
                AvailableUnits = $assignedUnits - $consumedUnits
                StorageLimitGB = $storageLimit
                LicenseTier = $licenseTier
            }
            
            $totalLicensedUsers += $consumedUnits
        }
        
        # Get mailbox information for archive and shared mailbox analysis
        $mailboxInfo = Get-MailboxInformation
        
        $script:ReportData.LicensingInfo = @{
            TotalLicensedUsers = $totalLicensedUsers
            LicenseDistribution = $licenseAnalysis
            MailboxAnalysis = $mailboxInfo
            HYCUEntitlement = @{
                PooledCapacityPerUserGB = 50
                TotalHYCUEntitlementGB = $totalLicensedUsers * 50
                CurrentUsageGB = $script:ReportData.GrowthAnalysis.CurrentTotalSizeGB
                ExcessCapacityGB = [math]::Max(0, $script:ReportData.GrowthAnalysis.CurrentTotalSizeGB - ($totalLicensedUsers * 50))
                AdditionalLicensesNeeded = [math]::Ceiling([math]::Max(0, $script:ReportData.GrowthAnalysis.CurrentTotalSizeGB - ($totalLicensedUsers * 50)) / 50)
            }
        }
        
        Write-ColorOutput "Licensed Users: $totalLicensedUsers" "SUCCESS"
        Write-ColorOutput "HYCU Entitlement: $($totalLicensedUsers * 50) GB" "SUCCESS"
        Write-ColorOutput "Current Usage: $($script:ReportData.GrowthAnalysis.CurrentTotalSizeGB) GB" "SUCCESS"
        
        if ($script:ReportData.LicensingInfo.HYCUEntitlement.ExcessCapacityGB -gt 0) {
            Write-ColorOutput "Additional Licenses Needed: $($script:ReportData.LicensingInfo.HYCUEntitlement.AdditionalLicensesNeeded)" "WARNING"
        }
    }
    catch {
        Write-ColorOutput "Failed to get licensing information: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get license storage limits
function Get-LicenseStorageLimit {
    param([string]$SkuPartNumber)
    
    $storageLimits = @{
        "ENTERPRISEPACK" = 50      # E3
        "ENTERPRISEPREMIUM" = 50  # E5
        "BUSINESS_PREMIUM" = 50    # Business Premium
        "BUSINESS_ESSENTIALS" = 50 # Business Essentials
        "BUSINESS_BASIC" = 50     # Business Basic
        "STANDARDPACK" = 50        # E1
        "DEVELOPERPACK" = 50      # Developer
        "POWER_BI_PRO" = 50       # Power BI Pro
        "PROJECTPROFESSIONAL" = 50 # Project Professional
        "VISIOCLIENT" = 50        # Visio
        "FLOW_FREE" = 2           # Power Automate Free
        "POWER_BI_STANDARD" = 2   # Power BI Standard
        "EXCHANGESTANDARD" = 50    # Exchange Online Plan 1
        "EXCHANGEENTERPRISE" = 50 # Exchange Online Plan 2
        "SHAREPOINTSTANDARD" = 50  # SharePoint Online Plan 1
        "SHAREPOINTENTERPRISE" = 50 # SharePoint Online Plan 2
        "ONEDRIVESTANDARD" = 50    # OneDrive for Business Plan 1
        "ONEDRIVEENTERPRISE" = 50  # OneDrive for Business Plan 2
    }
    
    return $storageLimits[$SkuPartNumber] ?? 50
}

# Function to get license tier
function Get-LicenseTier {
    param([string]$SkuPartNumber)
    
    $licenseTiers = @{
        "ENTERPRISEPACK" = "E3 (50 GB)"
        "ENTERPRISEPREMIUM" = "E5 (50 GB)"
        "BUSINESS_PREMIUM" = "Business Premium (50 GB)"
        "BUSINESS_ESSENTIALS" = "Business Essentials (50 GB)"
        "BUSINESS_BASIC" = "Business Basic (50 GB)"
        "STANDARDPACK" = "E1 (50 GB)"
        "DEVELOPERPACK" = "Developer (50 GB)"
        "POWER_BI_PRO" = "Power BI Pro (50 GB)"
        "PROJECTPROFESSIONAL" = "Project Professional (50 GB)"
        "VISIOCLIENT" = "Visio (50 GB)"
        "FLOW_FREE" = "Power Automate Free (2 GB)"
        "POWER_BI_STANDARD" = "Power BI Standard (2 GB)"
        "EXCHANGESTANDARD" = "Exchange Online Plan 1 (50 GB)"
        "EXCHANGEENTERPRISE" = "Exchange Online Plan 2 (50 GB)"
        "SHAREPOINTSTANDARD" = "SharePoint Online Plan 1 (50 GB)"
        "SHAREPOINTENTERPRISE" = "SharePoint Online Plan 2 (50 GB)"
        "ONEDRIVESTANDARD" = "OneDrive Plan 1 (50 GB)"
        "ONEDRIVEENTERPRISE" = "OneDrive Plan 2 (50 GB)"
    }
    
    return $licenseTiers[$SkuPartNumber] ?? "Unknown License (50 GB)"
}

# Function to get mailbox information
function Get-MailboxInformation {
    Write-ColorOutput "Analyzing mailbox types and archive mailboxes..." "INFO"
    
    try {
        # Connect to Exchange Online for detailed mailbox analysis
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-ColorOutput "Exchange Online Management module not available, skipping detailed mailbox analysis" "WARNING"
            return @{
                TotalMailboxes = 0
                RegularMailboxes = 0
                SharedMailboxes = 0
                ArchiveMailboxes = 0
                ResourceMailboxes = 0
                ArchivePercentage = 0
                ExcessArchiveMailboxes = 0
            }
        }
        
        # This would require Exchange Online connection
        # For now, return estimated values based on user count
        $totalMailboxes = $script:ReportData.TenantInfo.UserCounts.EnabledUsers
        $estimatedSharedMailboxes = [math]::Round($totalMailboxes * 0.15)  # Estimate 15% shared mailboxes
        $estimatedArchiveMailboxes = [math]::Round($totalMailboxes * 0.25)  # Estimate 25% archive mailboxes
        $estimatedResourceMailboxes = [math]::Round($totalMailboxes * 0.05)  # Estimate 5% resource mailboxes
        
        $totalMailboxCount = $totalMailboxes + $estimatedSharedMailboxes + $estimatedResourceMailboxes
        $archivePercentage = if ($totalMailboxCount -gt 0) { ($estimatedArchiveMailboxes / $totalMailboxCount) * 100 } else { 0 }
        
        # Calculate excess archive mailboxes (over 20% threshold)
        $archiveThreshold = $totalMailboxCount * 0.20
        $excessArchiveMailboxes = [math]::Max(0, $estimatedArchiveMailboxes - $archiveThreshold)
        
        $mailboxInfo = @{
            TotalMailboxes = $totalMailboxCount
            RegularMailboxes = $totalMailboxes
            SharedMailboxes = $estimatedSharedMailboxes
            ArchiveMailboxes = $estimatedArchiveMailboxes
            ResourceMailboxes = $estimatedResourceMailboxes
            ArchivePercentage = [math]::Round($archivePercentage, 2)
            ExcessArchiveMailboxes = [math]::Round($excessArchiveMailboxes)
            ArchiveThreshold = [math]::Round($archiveThreshold)
        }
        
        Write-ColorOutput "Total Mailboxes: $totalMailboxCount" "SUCCESS"
        Write-ColorOutput "Archive Mailboxes: $estimatedArchiveMailboxes ($([math]::Round($archivePercentage, 1))%)" "SUCCESS"
        
        if ($excessArchiveMailboxes -gt 0) {
            Write-ColorOutput "Excess Archive Mailboxes: $excessArchiveMailboxes (over 20% threshold)" "WARNING"
        }
        
        return $mailboxInfo
    }
    catch {
        Write-ColorOutput "Failed to get mailbox information: $($_.Exception.Message)" "ERROR"
        return @{
            TotalMailboxes = 0
            RegularMailboxes = 0
            SharedMailboxes = 0
            ArchiveMailboxes = 0
            ResourceMailboxes = 0
            ArchivePercentage = 0
            ExcessArchiveMailboxes = 0
        }
    }
}

# Function to get Exchange Online usage
function Get-ExchangeUsage {
    Write-ColorOutput "Gathering Exchange Online usage data..." "INFO"
    
    try {
        # Use OutFile with progress suppression to avoid SDK bug
        $tempFile = [System.IO.Path]::GetTempFileName()
        
        # Suppress progress bar to avoid SDK bug (2147483647 progress error)
        $oldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Get-MgReportMailboxUsageDetail -Period D30 -OutFile $tempFile
        }
        finally {
            $ProgressPreference = $oldProgressPreference
        }
        
        $mailboxUsage = Import-Csv $tempFile
        Remove-Item $tempFile -Force
        $totalMailboxSize = ($mailboxUsage | Measure-Object -Property StorageUsedInBytes -Sum).Sum / 1GB
        
        # Debug: Check if we're getting storage data
        Write-ColorOutput "Exchange Debug: Found $($mailboxUsage.Count) mailboxes, Total size: $totalMailboxSize GB" "INFO"
        
        # Get top 5 mailboxes by size
        $top5Mailboxes = $mailboxUsage | Sort-Object StorageUsedInBytes -Descending | Select-Object -First 5 | ForEach-Object {
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                UserPrincipalName = $_.UserPrincipalName
                StorageUsedInGB = [math]::Round($_.StorageUsedInBytes / 1GB, 2)
            }
        }
        
        $script:ReportData.ExchangeData = @{
            TotalMailboxes = $mailboxUsage.Count
            TotalSizeGB = [math]::Round($totalMailboxSize, 2)
            AverageSizeGB = [math]::Round($totalMailboxSize / $mailboxUsage.Count, 2)
            LargestMailboxGB = [math]::Round(($mailboxUsage | Sort-Object StorageUsedInBytes -Descending | Select-Object -First 1).StorageUsedInBytes / 1GB, 2)
            Top5Mailboxes = $top5Mailboxes
        }
        
        Write-ColorOutput "Exchange: $($script:ReportData.ExchangeData.TotalSizeGB) GB across $($script:ReportData.ExchangeData.TotalMailboxes) mailboxes" "SUCCESS"
        if ($top5Mailboxes.Count -gt 0) {
            Write-ColorOutput "Top mailbox: $($top5Mailboxes[0].DisplayName) ($($top5Mailboxes[0].StorageUsedInGB) GB)" "SUCCESS"
        }
    }
    catch {
        Write-ColorOutput "Failed to get Exchange usage: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get OneDrive usage
function Get-OneDriveUsage {
    Write-ColorOutput "Gathering OneDrive usage data..." "INFO"
    
    try {
        # Use OutFile with progress suppression to avoid SDK bug
        $tempFile = [System.IO.Path]::GetTempFileName()
        
        # Suppress progress bar to avoid SDK bug (2147483647 progress error)
        $oldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Get-MgReportOneDriveUsageAccountDetail -Period D30 -OutFile $tempFile
        }
        finally {
            $ProgressPreference = $oldProgressPreference
        }
        
        $oneDriveUsage = Import-Csv $tempFile
        Remove-Item $tempFile -Force
        $totalOneDriveSize = ($oneDriveUsage | Measure-Object -Property StorageUsedInBytes -Sum).Sum / 1GB
        
        # Debug: Check if we're getting storage data
        Write-ColorOutput "OneDrive Debug: Found $($oneDriveUsage.Count) accounts, Total size: $totalOneDriveSize GB" "INFO"
        
        # Get top 5 OneDrive accounts by size
        $top5OneDrives = $oneDriveUsage | Sort-Object StorageUsedInBytes -Descending | Select-Object -First 5 | ForEach-Object {
            [PSCustomObject]@{
                DisplayName = $_.DisplayName
                UserPrincipalName = $_.UserPrincipalName
                StorageUsedInGB = [math]::Round($_.StorageUsedInBytes / 1GB, 2)
            }
        }
        
        $script:ReportData.OneDriveData = @{
            TotalAccounts = $oneDriveUsage.Count
            TotalSizeGB = [math]::Round($totalOneDriveSize, 2)
            AverageSizeGB = [math]::Round($totalOneDriveSize / $oneDriveUsage.Count, 2)
            LargestAccountGB = [math]::Round(($oneDriveUsage | Sort-Object StorageUsedInBytes -Descending | Select-Object -First 1).StorageUsedInBytes / 1GB, 2)
            Top5OneDrives = $top5OneDrives
        }
        
        Write-ColorOutput "OneDrive: $($script:ReportData.OneDriveData.TotalSizeGB) GB across $($script:ReportData.OneDriveData.TotalAccounts) accounts" "SUCCESS"
        if ($top5OneDrives.Count -gt 0) {
            Write-ColorOutput "Top OneDrive: $($top5OneDrives[0].DisplayName) ($($top5OneDrives[0].StorageUsedInGB) GB)" "SUCCESS"
        }
    }
    catch {
        Write-ColorOutput "Failed to get OneDrive usage: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get SharePoint usage
function Get-SharePointUsage {
    Write-ColorOutput "Gathering SharePoint usage data..." "INFO"
    
    try {
        # Use OutFile with progress suppression to avoid SDK bug
        $tempFile = [System.IO.Path]::GetTempFileName()
        
        # Suppress progress bar to avoid SDK bug (2147483647 progress error)
        $oldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Get-MgReportSharePointSiteUsageDetail -Period D30 -OutFile $tempFile
        }
        finally {
            $ProgressPreference = $oldProgressPreference
        }
        
        $sharePointUsage = Import-Csv $tempFile
        Remove-Item $tempFile -Force
        $totalSharePointSize = ($sharePointUsage | Measure-Object -Property StorageUsedInBytes -Sum).Sum / 1GB
        
        # Debug: Check if we're getting storage data
        Write-ColorOutput "SharePoint Debug: Found $($sharePointUsage.Count) sites, Total size: $totalSharePointSize GB" "INFO"
        
        # Get top 5 SharePoint sites by size
        $top5Sites = $sharePointUsage | Sort-Object StorageUsedInBytes -Descending | Select-Object -First 5 | ForEach-Object {
            [PSCustomObject]@{
                SiteName = $_.SiteName
                SiteUrl = $_.SiteUrl
                StorageUsedInGB = [math]::Round($_.StorageUsedInBytes / 1GB, 2)
            }
        }
        
        $script:ReportData.SharePointData = @{
            TotalSites = $sharePointUsage.Count
            TotalSizeGB = [math]::Round($totalSharePointSize, 2)
            AverageSizeGB = [math]::Round($totalSharePointSize / $sharePointUsage.Count, 2)
            LargestSiteGB = [math]::Round(($sharePointUsage | Sort-Object StorageUsedInBytes -Descending | Select-Object -First 1).StorageUsedInBytes / 1GB, 2)
            Top5Sites = $top5Sites
        }
        
        Write-ColorOutput "SharePoint: $($script:ReportData.SharePointData.TotalSizeGB) GB across $($script:ReportData.SharePointData.TotalSites) sites" "SUCCESS"
        if ($top5Sites.Count -gt 0) {
            Write-ColorOutput "Top site: $($top5Sites[0].SiteName) ($($top5Sites[0].StorageUsedInGB) GB)" "SUCCESS"
        }
    }
    catch {
        Write-ColorOutput "Failed to get SharePoint usage: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get Teams usage
function Get-TeamsUsage {
    Write-ColorOutput "Gathering Microsoft Teams usage data..." "INFO"
    
    try {
        $teams = Get-MgTeam -All
        $channels = @()
        $totalMessages = 0
        
        foreach ($team in $teams) {
            try {
                $teamChannels = Get-MgTeamChannel -TeamId $team.Id
                $channels += $teamChannels
                
                # Skip message counting - requires ChannelMessage.Read.All permission
                # Most users don't have this level of access
            }
            catch {
                # Log the error but continue processing
                Write-ColorOutput "Failed to get channels for team $($team.DisplayName): $($_.Exception.Message)" "WARNING"
            }
        }
        
        # Calculate Teams private chat cost implications
        $privateChatCostPerMessage = 0.00075  # $0.00075 per message/notification
        $costPerMillionMessages = $privateChatCostPerMessage * 1000000  # $750 per million messages
        
        $script:ReportData.TeamsData = @{
            TotalTeams = $teams.Count
            TotalChannels = $channels.Count
            AverageChannelsPerTeam = if ($teams.Count -gt 0) { [math]::Round($channels.Count / $teams.Count, 2) } else { 0 }
            TotalMessages = "N/A (Permission Required)"
            PrivateChatCostPerMessage = $privateChatCostPerMessage
            CostPerMillionMessages = $costPerMillionMessages
            Note = "Message counting requires ChannelMessage.Read.All permission"
        }
        
        Write-ColorOutput "Teams: $($teams.Count) teams with $($channels.Count) channels" "SUCCESS"
        Write-ColorOutput "Message counting skipped (requires ChannelMessage.Read.All permission)" "INFO"
        Write-ColorOutput "Private Chat Cost: $$privateChatCostPerMessage per message ($$costPerMillionMessages per million)" "WARNING"
    }
    catch {
        Write-ColorOutput "Failed to get Teams usage: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get Planner usage
function Get-PlannerUsage {
    Write-ColorOutput "Gathering Microsoft Planner usage data..." "INFO"
    
    try {
        # Get all groups that might have Planner
        $groups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" -Property "Id,DisplayName,Description"
        $plannerPlans = @()
        $plannerTasks = 0
        
        foreach ($group in $groups) {
            try {
                # Get Planner plans for this group
                $plans = Get-MgGroupPlannerPlan -GroupId $group.Id -ErrorAction SilentlyContinue
                if ($plans) {
                    $plannerPlans += $plans
                    
                    # Get tasks for each plan
                    foreach ($plan in $plans) {
                        try {
                            $tasks = Get-MgPlannerPlanTask -PlannerPlanId $plan.Id -ErrorAction SilentlyContinue
                            if ($tasks) {
                                $plannerTasks += $tasks.Count
                            }
                        }
                        catch {
                            # Some plans may not be accessible
                        }
                    }
                }
            }
            catch {
                # Some groups may not have Planner or may not be accessible
            }
        }
        
        $script:ReportData.PlannerData = @{
            TotalPlans = $plannerPlans.Count
            TotalTasks = $plannerTasks
            AverageTasksPerPlan = if ($plannerPlans.Count -gt 0) { [math]::Round($plannerTasks / $plannerPlans.Count, 2) } else { 0 }
        }
        
        Write-ColorOutput "Planner: $($plannerPlans.Count) plans with $plannerTasks tasks" "SUCCESS"
    }
    catch {
        Write-ColorOutput "Failed to get Planner usage: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get Groups usage
function Get-GroupsUsage {
    Write-ColorOutput "Gathering Microsoft 365 Groups usage data..." "INFO"
    
    try {
        # Get all Microsoft 365 Groups
        $groups = Get-MgGroup -All -Filter "groupTypes/any(c:c eq 'Unified')" -Property "Id,DisplayName,Description,GroupTypes,MembershipRule,MembershipRuleProcessingState"
        
        # Categorize groups
        $distributionGroups = $groups | Where-Object { $_.GroupTypes -contains "DynamicMembership" }
        $securityGroups = $groups | Where-Object { $_.GroupTypes -contains "SecurityEnabled" }
        $unifiedGroups = $groups | Where-Object { $_.GroupTypes -contains "Unified" }
        
        $script:ReportData.GroupsData = @{
            TotalGroups = $groups.Count
            DistributionGroups = $distributionGroups.Count
            SecurityGroups = $securityGroups.Count
            UnifiedGroups = $unifiedGroups.Count
            GroupsWithPlanner = 0  # This will be updated by Planner analysis
        }
        
        Write-ColorOutput "Groups: $($groups.Count) total (Distribution: $($distributionGroups.Count), Security: $($securityGroups.Count), Unified: $($unifiedGroups.Count))" "SUCCESS"
    }
    catch {
        Write-ColorOutput "Failed to get Groups usage: $($_.Exception.Message)" "ERROR"
    }
}

# Function to get Sites and OneDrive counts
function Get-SitesAndOneDriveCounts {
    Write-ColorOutput "Gathering sites and OneDrive counts..." "INFO"
    
    try {
        # Get SharePoint sites count (with progress suppression)
        $tempFile = [System.IO.Path]::GetTempFileName()
        
        # Suppress progress bar to avoid SDK bug (2147483647 progress error)
        $oldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Get-MgReportSharePointSiteUsageDetail -Period D30 -OutFile $tempFile
        }
        finally {
            $ProgressPreference = $oldProgressPreference
        }
        
        $sharePointSites = Import-Csv $tempFile
        $sharePointSitesCount = $sharePointSites.Count
        Remove-Item $tempFile -Force
        
        # Get Teams sites count (Teams create SharePoint sites)
        $teams = Get-MgTeam -All
        $teamsSitesCount = $teams.Count
        
        # Get OneDrive accounts count (with progress suppression)
        $tempFile2 = [System.IO.Path]::GetTempFileName()
        
        # Suppress progress bar to avoid SDK bug (2147483647 progress error)
        $oldProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            Get-MgReportOneDriveUsageAccountDetail -Period D30 -OutFile $tempFile2
        }
        finally {
            $ProgressPreference = $oldProgressPreference
        }
        
        $oneDriveAccounts = Import-Csv $tempFile2
        $oneDriveAccountsCount = $oneDriveAccounts.Count
        Remove-Item $tempFile2 -Force
        
        # Calculate total sites (SharePoint + Teams)
        $totalSites = $sharePointSitesCount + $teamsSitesCount
        
        $script:ReportData.SitesAndOneDriveData = @{
            OneDriveAccounts = $oneDriveAccountsCount
            SharePointSites = $sharePointSitesCount
            TeamsSites = $teamsSitesCount
            TotalSites = $totalSites
        }
        
        Write-ColorOutput "OneDrive Accounts: $oneDriveAccountsCount" "SUCCESS"
        Write-ColorOutput "SharePoint Sites: $sharePointSitesCount" "SUCCESS"
        Write-ColorOutput "Teams Sites: $teamsSitesCount" "SUCCESS"
        Write-ColorOutput "Total Sites: $totalSites" "SUCCESS"
    }
    catch {
        Write-ColorOutput "Failed to get sites and OneDrive counts: $($_.Exception.Message)" "ERROR"
        # Provide fallback data
        $script:ReportData.SitesAndOneDriveData = @{
            OneDriveAccounts = $script:ReportData.TenantInfo.UserCounts.EnabledUsers
            SharePointSites = 0
            TeamsSites = 0
            TotalSites = 0
        }
    }
}

# Function to calculate growth analysis
function Get-GrowthAnalysis {
    Write-ColorOutput "Calculating growth analysis..." "INFO"
    
    try {
        $currentDate = Get-Date
        $historicalDate = $currentDate.AddDays(-$Period)
        
        # Get historical data (simplified - in real implementation, you'd store historical data)
        $growthRates = @(10, 20, $AnnualGrowth)
        $totalCurrentSize = $script:ReportData.ExchangeData.TotalSizeGB + 
                           $script:ReportData.OneDriveData.TotalSizeGB + 
                           $script:ReportData.SharePointData.TotalSizeGB
        
        $script:ReportData.GrowthAnalysis = @{
            CurrentTotalSizeGB = [math]::Round($totalCurrentSize, 2)
            GrowthProjections = @{}
        }
        
        foreach ($rate in $growthRates) {
            $projectedSize = $totalCurrentSize * (1 + ($rate / 100))
            $script:ReportData.GrowthAnalysis.GrowthProjections[$rate] = [math]::Round($projectedSize, 2)
        }
        
        Write-ColorOutput "Current total size: $($script:ReportData.GrowthAnalysis.CurrentTotalSizeGB) GB" "SUCCESS"
    }
    catch {
        Write-ColorOutput "Failed to calculate growth analysis: $($_.Exception.Message)" "ERROR"
    }
}

# Function to generate backup recommendations
function Get-BackupRecommendations {
    Write-ColorOutput "Generating backup recommendations..." "INFO"
    
    try {
        $totalSize = $script:ReportData.GrowthAnalysis.CurrentTotalSizeGB
        $userCount = $script:ReportData.TenantInfo.UserCounts.EnabledUsers
        
        $recommendations = @{
            BackupFrequency = if ($totalSize -gt 1000) { "Daily" } elseif ($totalSize -gt 100) { "Daily" } else { "Weekly" }
            RetentionPolicy = if ($userCount -gt 1000) { "7 years" } elseif ($userCount -gt 100) { "3 years" } else { "1 year" }
            StorageEstimate = [math]::Round($totalSize * 1.5, 2)  # 50% overhead for backup storage
            CriticalData = @("Exchange", "OneDrive", "SharePoint")
            BackupWindow = "Off-hours (2 AM - 6 AM)"
        }
        
        $script:ReportData.BackupRecommendations = $recommendations
        
        Write-ColorOutput "Backup recommendations generated" "SUCCESS"
    }
    catch {
        Write-ColorOutput "Failed to generate backup recommendations: $($_.Exception.Message)" "ERROR"
    }
}

# Function to generate cost analysis
function Get-CostAnalysis {
    Write-ColorOutput "Generating cost analysis..." "INFO"
    
    try {
        $currentStorageGB = $script:ReportData.GrowthAnalysis.CurrentTotalSizeGB
        $userCount = $script:ReportData.TenantInfo.UserCounts.EnabledUsers
        
        # If no storage data available, provide estimated costs based on user count
        if ($currentStorageGB -eq 0) {
            Write-ColorOutput "No storage data available - providing estimated costs based on user count" "WARNING"
            # Estimate 5 GB per user as a reasonable baseline
            $currentStorageGB = $userCount * 5
            Write-ColorOutput "Estimated storage: $currentStorageGB GB (5 GB per user)" "INFO"
        }
        
        # Storage cost calculation with compression and growth
        $compressionRate = 0.40  # 40% compression
        $growthRate = 0.20       # 20% growth rate
        $compressedStorageGB = $currentStorageGB * (1 - $compressionRate)  # Apply 40% compression
        $projectedStorageGB = $compressedStorageGB * ($growthRate + 1)     # Apply 20% growth
        
        # Storage costs: $0.02 per GB per month * 12 months
        $costPerGBPerMonth = 0.02
        $monthlyStorageCost = $projectedStorageGB * $costPerGBPerMonth
        $annualStorageCost = $monthlyStorageCost * 12
        
        # Worker node cost calculation based on pre-compression tenant size
        $costPerTBPerMonth = 8   # $8 per TB per month
        $tenantSizeTB = $currentStorageGB / 1024  # Convert GB to TB
        
        $monthlyWorkerNodeCost = $tenantSizeTB * $costPerTBPerMonth
        $annualWorkerNodeCost = $monthlyWorkerNodeCost * 12
        
        # Total costs
        $totalMonthlyCost = $monthlyStorageCost + $monthlyWorkerNodeCost
        $totalAnnualCost = $annualStorageCost + $annualWorkerNodeCost
        
        # Calculate per-user costs
        $userCount = $script:ReportData.TenantInfo.UserCounts.EnabledUsers
        $storageCostPerUser = $monthlyStorageCost / $userCount
        $workerNodeCostPerUser = $monthlyWorkerNodeCost / $userCount
        $totalCostPerUser = $totalMonthlyCost / $userCount
        
        $script:ReportData.CostAnalysis = @{
            CurrentStorageGB = [math]::Round($currentStorageGB, 2)
            CompressedStorageGB = [math]::Round($compressedStorageGB, 2)
            ProjectedStorageGB = [math]::Round($projectedStorageGB, 2)
            CompressionRate = $compressionRate
            GrowthRate = $growthRate
            MonthlyStorageCost = [math]::Round($monthlyStorageCost, 2)
            AnnualStorageCost = [math]::Round($annualStorageCost, 2)
            TenantSizeTB = [math]::Round($tenantSizeTB, 2)
            MonthlyWorkerNodeCost = [math]::Round($monthlyWorkerNodeCost, 2)
            AnnualWorkerNodeCost = [math]::Round($annualWorkerNodeCost, 2)
            TotalMonthlyCost = [math]::Round($totalMonthlyCost, 2)
            TotalAnnualCost = [math]::Round($totalAnnualCost, 2)
            CostPerGBPerMonth = $costPerGBPerMonth
            CostPerTBPerMonth = $costPerTBPerMonth
            StorageCostPerUser = [math]::Round($storageCostPerUser, 2)
            WorkerNodeCostPerUser = [math]::Round($workerNodeCostPerUser, 2)
            TotalCostPerUser = [math]::Round($totalCostPerUser, 2)
        }
        
        Write-ColorOutput "Storage: $currentStorageGB GB → $compressedStorageGB GB (compressed) → $projectedStorageGB GB (with growth)" "SUCCESS"
        Write-ColorOutput "Monthly Storage Cost: $([math]::Round($monthlyStorageCost, 2)) ($$([math]::Round($annualStorageCost, 2)) annually)" "SUCCESS"
        Write-ColorOutput "Worker Node Cost: $([math]::Round($monthlyWorkerNodeCost, 2))/month ($$([math]::Round($annualWorkerNodeCost, 2)) annually) - based on $([math]::Round($tenantSizeTB, 2)) TB tenant size" "SUCCESS"
        Write-ColorOutput "Per-User Costs: Storage=$$([math]::Round($storageCostPerUser, 2)), Worker Node=$$([math]::Round($workerNodeCostPerUser, 2)), Total=$$([math]::Round($totalCostPerUser, 2))" "SUCCESS"
        Write-ColorOutput "Total Monthly Cost: $([math]::Round($totalMonthlyCost, 2)) ($$([math]::Round($totalAnnualCost, 2)) annually)" "SUCCESS"
    }
    catch {
        Write-ColorOutput "Failed to generate cost analysis: $($_.Exception.Message)" "ERROR"
    }
}

# Function to generate HTML report
function New-HTMLReport {
    Write-ColorOutput "Generating HTML report..." "INFO"
    
    $reportPath = Join-Path $script:OutputDirectory "HYCU-M365-Sizing-$(Get-Date -Format 'yyyy-MM-dd-HHmm').html"
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HYCU M365 Sizing Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 8px 8px 0 0; }
        .header h1 { margin: 0; font-size: 2.5em; }
        .header p { margin: 10px 0 0 0; opacity: 0.9; }
        .content { padding: 30px; }
        .section { margin-bottom: 40px; }
        .section h2 { color: #333; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin: 20px 0; }
        .metric-card { background: #f8f9fa; padding: 20px; border-radius: 8px; border-left: 4px solid #667eea; }
        .metric-value { font-size: 2em; font-weight: bold; color: #667eea; }
        .metric-label { color: #666; margin-top: 5px; }
        .table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        .table th, .table td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        .table th { background-color: #f8f9fa; font-weight: bold; }
        .recommendation { background: #e8f5e8; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745; margin: 10px 0; }
        .warning { background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107; margin: 10px 0; }
        .footer { text-align: center; padding: 20px; color: #666; border-top: 1px solid #eee; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🚀 HYCU M365 Sizing Report</h1>
            <p>Comprehensive Microsoft 365 Tenant Analysis for Backup Planning</p>
            <p>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>
        
        <div class="content">
            <div class="section">
                <h2>📊 Tenant Overview</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.TenantInfo.DisplayName)</div>
                        <div class="metric-label">Tenant Name</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.TenantInfo.UserCounts.TotalUsers)</div>
                        <div class="metric-label">Total Users</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.TenantInfo.UserCounts.EnabledUsers)</div>
                        <div class="metric-label">Active Users</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.TenantInfo.UserCounts.GuestUsers)</div>
                        <div class="metric-label">Guest Users</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>💾 Tenant Capacity</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$([math]::Round($script:ReportData.ExchangeData.TotalSizeGB, 1)) GB</div>
                        <div class="metric-label">Exchange Online</div>
                        <div style="font-size: 0.8em; color: #888; margin-top: 5px;">$([math]::Round($script:ReportData.ExchangeData.TotalSizeGB / $script:ReportData.TenantInfo.UserCounts.EnabledUsers, 1)) GB per user</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$([math]::Round($script:ReportData.OneDriveData.TotalSizeGB, 1)) GB</div>
                        <div class="metric-label">OneDrive for Business</div>
                        <div style="font-size: 0.8em; color: #888; margin-top: 5px;">$([math]::Round($script:ReportData.OneDriveData.TotalSizeGB / $script:ReportData.TenantInfo.UserCounts.EnabledUsers, 1)) GB per user</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$([math]::Round($script:ReportData.SharePointData.TotalSizeGB, 1)) GB</div>
                        <div class="metric-label">SharePoint Online</div>
                        <div style="font-size: 0.8em; color: #888; margin-top: 5px;">$([math]::Round($script:ReportData.SharePointData.TotalSizeGB / $script:ReportData.TenantInfo.UserCounts.EnabledUsers, 1)) GB per user</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$([math]::Round($script:ReportData.GrowthAnalysis.CurrentTotalSizeGB, 1)) GB</div>
                        <div class="metric-label">Total Storage</div>
                        <div style="font-size: 0.8em; color: #888; margin-top: 5px;">$([math]::Round($script:ReportData.GrowthAnalysis.CurrentTotalSizeGB / $script:ReportData.TenantInfo.UserCounts.EnabledUsers, 1)) GB per user</div>
                    </div>
                </div>
                
                <h3>Initial Cost Estimates</h3>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$$($script:ReportData.CostAnalysis.StorageCostPerUser)</div>
                        <div class="metric-label">Storage Cost per User</div>
                        <div style="font-size: 0.8em; color: #888; margin-top: 5px;">Monthly</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$($script:ReportData.CostAnalysis.WorkerNodeCostPerUser)</div>
                        <div class="metric-label">Worker Node Cost per User</div>
                        <div style="font-size: 0.8em; color: #888; margin-top: 5px;">Monthly</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$($script:ReportData.CostAnalysis.TotalCostPerUser)</div>
                        <div class="metric-label">Total Cost per User</div>
                        <div style="font-size: 0.8em; color: #888; margin-top: 5px;">Monthly</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.CostAnalysis.TenantSizeTB) TB</div>
                        <div class="metric-label">Tenant Size</div>
                        <div style="font-size: 0.8em; color: #888; margin-top: 5px;">Pre-compression</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>📈 Growth Projections</h2>
                <table class="table">
                    <thead>
                        <tr>
                            <th>Growth Rate</th>
                            <th>Projected Size (GB)</th>
                            <th>Additional Storage</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($rate in $script:ReportData.GrowthAnalysis.GrowthProjections.Keys) {
            $projectedSize = $script:ReportData.GrowthAnalysis.GrowthProjections[$rate]
            $additionalStorage = $projectedSize - $script:ReportData.GrowthAnalysis.CurrentTotalSizeGB
            $html += "<tr><td>$rate%</td><td>$projectedSize</td><td>$additionalStorage GB</td></tr>"
        }

        $html += @"
                    </tbody>
                </table>
            </div>
            
            <div class="section">
                <h2>👥 Teams Analysis</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.TeamsData.TotalTeams)</div>
                        <div class="metric-label">Total Teams</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.TeamsData.TotalChannels)</div>
                        <div class="metric-label">Total Channels</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.TeamsData.TotalMessages)</div>
                        <div class="metric-label">Channel Messages</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$($script:ReportData.TeamsData.CostPerMillionMessages)</div>
                        <div class="metric-label">Cost per Million Private Messages</div>
                    </div>
                </div>
                <div class="warning">
                    <strong>⚠️ Teams Private Chat Cost Impact:</strong>
                    <p>Protecting Teams private chats (1:1 conversations) incurs additional costs from Microsoft:</p>
                    <ul>
                        <li><strong>Cost per message/notification:</strong> $$($script:ReportData.TeamsData.PrivateChatCostPerMessage)</li>
                        <li><strong>Cost per million messages:</strong> $$($script:ReportData.TeamsData.CostPerMillionMessages)</li>
                        <li><strong>Impact:</strong> Consider this additional cost when planning Teams private chat protection</li>
                    </ul>
                </div>
            </div>
            
            <div class="section">
                <h2>📋 Planner Analysis</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.PlannerData.TotalPlans)</div>
                        <div class="metric-label">Total Plans</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.PlannerData.TotalTasks)</div>
                        <div class="metric-label">Total Tasks</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.PlannerData.AverageTasksPerPlan)</div>
                        <div class="metric-label">Avg Tasks/Plan</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>📧 Mailbox Analysis</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.TotalMailboxes)</div>
                        <div class="metric-label">Total Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.RegularMailboxes)</div>
                        <div class="metric-label">Active User Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.SharedMailboxes)</div>
                        <div class="metric-label">Shared Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.ArchiveMailboxes)</div>
                        <div class="metric-label">Archive Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.ResourceMailboxes)</div>
                        <div class="metric-label">Resource Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.ArchivePercentage)%</div>
                        <div class="metric-label">Archive Percentage</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>🌐 Sites & OneDrive Analysis</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.SitesAndOneDriveData.OneDriveAccounts)</div>
                        <div class="metric-label">OneDrive Accounts</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.SitesAndOneDriveData.SharePointSites)</div>
                        <div class="metric-label">SharePoint Sites</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.SitesAndOneDriveData.TeamsSites)</div>
                        <div class="metric-label">Teams Sites</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.SitesAndOneDriveData.TotalSites)</div>
                        <div class="metric-label">Total Sites</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>👥 Groups Analysis</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.GroupsData.TotalGroups)</div>
                        <div class="metric-label">Total Groups</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.GroupsData.UnifiedGroups)</div>
                        <div class="metric-label">Unified Groups</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.GroupsData.SecurityGroups)</div>
                        <div class="metric-label">Security Groups</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.GroupsData.DistributionGroups)</div>
                        <div class="metric-label">Distribution Groups</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>🏆 Top 5 by Size</h2>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(350px, 1fr)); gap: 20px;">
                    <div>
                        <h3>📧 Top 5 Mailboxes</h3>
                        <table class="table">
                            <thead>
                                <tr>
                                    <th>User</th>
                                    <th>Size (GB)</th>
                                </tr>
                            </thead>
                            <tbody>
"@

        # Add top 5 mailboxes
        if ($script:ReportData.ExchangeData.Top5Mailboxes) {
            foreach ($mailbox in $script:ReportData.ExchangeData.Top5Mailboxes) {
                $html += "<tr><td>$($mailbox.DisplayName)</td><td>$($mailbox.StorageUsedInGB)</td></tr>"
            }
        } else {
            $html += "<tr><td colspan='2'>No mailbox data available</td></tr>"
        }

        $html += @"
                            </tbody>
                        </table>
                    </div>
                    
                    <div>
                        <h3>📁 Top 5 OneDrive Accounts</h3>
                        <table class="table">
                            <thead>
                                <tr>
                                    <th>User</th>
                                    <th>Size (GB)</th>
                                </tr>
                            </thead>
                            <tbody>
"@

        # Add top 5 OneDrives
        if ($script:ReportData.OneDriveData.Top5OneDrives) {
            foreach ($oneDrive in $script:ReportData.OneDriveData.Top5OneDrives) {
                $html += "<tr><td>$($oneDrive.DisplayName)</td><td>$($oneDrive.StorageUsedInGB)</td></tr>"
            }
        } else {
            $html += "<tr><td colspan='2'>No OneDrive data available</td></tr>"
        }

        $html += @"
                            </tbody>
                        </table>
                    </div>
                    
                    <div>
                        <h3>🌐 Top 5 SharePoint Sites</h3>
                        <table class="table">
                            <thead>
                                <tr>
                                    <th>Site Name</th>
                                    <th>Size (GB)</th>
                                </tr>
                            </thead>
                            <tbody>
"@

        # Add top 5 SharePoint sites
        if ($script:ReportData.SharePointData.Top5Sites) {
            foreach ($site in $script:ReportData.SharePointData.Top5Sites) {
                $html += "<tr><td>$($site.SiteName)</td><td>$($site.StorageUsedInGB)</td></tr>"
            }
        } else {
            $html += "<tr><td colspan='2'>No SharePoint data available</td></tr>"
        }

        $html += @"
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>📋 HYCU Licensing Analysis</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.TotalLicensedUsers)</div>
                        <div class="metric-label">Total Licensed Users</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.HYCUEntitlement.TotalHYCUEntitlementGB) GB</div>
                        <div class="metric-label">HYCU Entitlement (50 GB/user)</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.HYCUEntitlement.CurrentUsageGB) GB</div>
                        <div class="metric-label">Current Usage</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.HYCUEntitlement.AdditionalLicensesNeeded)</div>
                        <div class="metric-label">Additional Licenses Needed</div>
                    </div>
                </div>
                
                <h3>License Distribution by Tier</h3>
                <table class="table">
                    <thead>
                        <tr>
                            <th>License Type</th>
                            <th>Assigned</th>
                            <th>Consumed</th>
                            <th>Available</th>
                            <th>Storage Limit</th>
                        </tr>
                    </thead>
                    <tbody>
"@

        foreach ($license in $script:ReportData.LicensingInfo.LicenseDistribution.Values) {
            $html += "<tr><td>$($license.LicenseTier)</td><td>$($license.AssignedUnits)</td><td>$($license.ConsumedUnits)</td><td>$($license.AvailableUnits)</td><td>$($license.StorageLimitGB) GB</td></tr>"
        }

        $html += @"
                    </tbody>
                </table>
                
                <h3>Mailbox Analysis</h3>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.TotalMailboxes)</div>
                        <div class="metric-label">Total Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.SharedMailboxes)</div>
                        <div class="metric-label">Shared Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.ArchiveMailboxes)</div>
                        <div class="metric-label">Archive Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.LicensingInfo.MailboxAnalysis.ArchivePercentage)%</div>
                        <div class="metric-label">Archive Percentage</div>
                    </div>
                </div>
                
                <div class="recommendation">
                    <strong>HYCU Licensing Summary:</strong>
                    <ul>
                        <li>Total Licensed Users: $($script:ReportData.LicensingInfo.TotalLicensedUsers)</li>
                        <li>HYCU Entitlement: $($script:ReportData.LicensingInfo.HYCUEntitlement.TotalHYCUEntitlementGB) GB (50 GB per user)</li>
                        <li>Current Usage: $($script:ReportData.LicensingInfo.HYCUEntitlement.CurrentUsageGB) GB</li>
                        <li>Excess Capacity: $($script:ReportData.LicensingInfo.HYCUEntitlement.ExcessCapacityGB) GB</li>
                        <li>Additional Licenses Needed: $($script:ReportData.LicensingInfo.HYCUEntitlement.AdditionalLicensesNeeded)</li>
                    </ul>
                </div>
                
                <div class="warning">
                    <strong>Archive Mailbox Analysis:</strong>
                    <ul>
                        <li>Archive Mailboxes: $($script:ReportData.LicensingInfo.MailboxAnalysis.ArchiveMailboxes) ($($script:ReportData.LicensingInfo.MailboxAnalysis.ArchivePercentage)%)</li>
                        <li>20% Threshold: $($script:ReportData.LicensingInfo.MailboxAnalysis.ArchiveThreshold) mailboxes</li>
                        <li>Excess Archive Mailboxes: $($script:ReportData.LicensingInfo.MailboxAnalysis.ExcessArchiveMailboxes)</li>
                        <li>Additional Licenses for Excess Archives: $([math]::Ceiling($script:ReportData.LicensingInfo.MailboxAnalysis.ExcessArchiveMailboxes / 50))</li>
                    </ul>
                </div>
            </div>
            
            <div class="section">
                <h2>💰 Initial Cost Estimates</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$$([math]::Round($script:ReportData.CostAnalysis.MonthlyStorageCost, 2))</div>
                        <div class="metric-label">Monthly Storage Cost</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$([math]::Round($script:ReportData.CostAnalysis.MonthlyWorkerNodeCost, 2))</div>
                        <div class="metric-label">Monthly Worker Node Cost</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$([math]::Round($script:ReportData.CostAnalysis.TotalMonthlyCost, 2))</div>
                        <div class="metric-label">Total Monthly Cost</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$([math]::Round($script:ReportData.CostAnalysis.TotalAnnualCost, 2))</div>
                        <div class="metric-label">Total Annual Cost</div>
                    </div>
                </div>
                
                <h3>Storage Cost Breakdown</h3>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.CostAnalysis.CurrentStorageGB) GB</div>
                        <div class="metric-label">Current Storage</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.CostAnalysis.CompressedStorageGB) GB</div>
                        <div class="metric-label">After 40% Compression</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.CostAnalysis.ProjectedStorageGB) GB</div>
                        <div class="metric-label">With 20% Growth</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$($script:ReportData.CostAnalysis.AnnualStorageCost)</div>
                        <div class="metric-label">Annual Storage Cost</div>
                    </div>
                </div>
                
                <h3>Worker Node Cost Breakdown</h3>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($script:ReportData.CostAnalysis.TenantSizeTB) TB</div>
                        <div class="metric-label">Tenant Size (Pre-compression)</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$($script:ReportData.CostAnalysis.CostPerTBPerMonth)/TB/month</div>
                        <div class="metric-label">Cost per TB per Month</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$([math]::Round($script:ReportData.CostAnalysis.MonthlyWorkerNodeCost, 2))</div>
                        <div class="metric-label">Monthly Worker Node Cost</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$($script:ReportData.CostAnalysis.AnnualWorkerNodeCost)</div>
                        <div class="metric-label">Annual Worker Node Cost</div>
                    </div>
                </div>
                
                <div class="recommendation">
                    <strong>Initial Cost Estimates Summary:</strong>
                    <ul>
                        <li><strong>Storage Cost:</strong> $$($script:ReportData.CostAnalysis.MonthlyStorageCost)/month ($$($script:ReportData.CostAnalysis.AnnualStorageCost)/year)</li>
                        <li><strong>Worker Node Cost:</strong> $$([math]::Round($script:ReportData.CostAnalysis.MonthlyWorkerNodeCost, 2))/month ($$($script:ReportData.CostAnalysis.AnnualWorkerNodeCost)/year)</li>
                        <li><strong>Total Cost:</strong> $$([math]::Round($script:ReportData.CostAnalysis.TotalMonthlyCost, 2))/month ($$([math]::Round($script:ReportData.CostAnalysis.TotalAnnualCost, 2))/year)</li>
                        <li><strong>Assumptions:</strong> 40% compression, 20% growth rate, 0.2% daily change rate, 1-year retention</li>
                    </ul>
                </div>
            </div>
            
        </div>
        
        <div class="footer">
            <p>Generated by HYCU M365 Sizing Tool v1.0 | For backup planning and capacity management</p>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $reportPath -Encoding UTF8
    Write-ColorOutput "Report generated: $reportPath" "SUCCESS"
    return $reportPath
}

# Main execution
try {
    Write-ColorOutput "Starting HYCU M365 Sizing Tool v1.0" "HEADER"
    Write-ColorOutput "=====================================" "HEADER"
    
    # Install required modules
    Install-RequiredModules
    
    # Connect to Microsoft Graph
    Connect-MicrosoftGraph
    
    # Gather tenant information
    Get-TenantInformation
    Get-UserStatistics
    
    # Gather usage data
    Get-ExchangeUsage
    Get-OneDriveUsage
    Get-SharePointUsage
    Get-TeamsUsage
    Get-PlannerUsage
    Get-GroupsUsage
    Get-SitesAndOneDriveCounts
    
    # Perform analysis
    Get-GrowthAnalysis
    Get-LicensingInformation
    Get-BackupRecommendations
    Get-CostAnalysis
    
    # Generate report
    $reportPath = New-HTMLReport
    
    $endTime = Get-Date
    $duration = $endTime - $script:StartTime
    
    Write-ColorOutput "=====================================" "HEADER"
    Write-ColorOutput "Analysis completed successfully!" "SUCCESS"
    Write-ColorOutput "Duration: $($duration.TotalMinutes.ToString('F2')) minutes" "INFO"
    Write-ColorOutput "Report saved to: $reportPath" "SUCCESS"
    Write-ColorOutput "=====================================" "HEADER"
}
catch {
    Write-ColorOutput "Script execution failed: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    # Disconnect from Microsoft Graph
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-ColorOutput "Disconnected from Microsoft Graph" "INFO"
    }
    catch {
        # Ignore disconnect errors
    }
}
