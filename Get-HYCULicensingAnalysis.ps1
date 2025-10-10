#Requires -Version 5.1

<#
.SYNOPSIS
    HYCU Licensing Analysis - Comprehensive Microsoft 365 licensing analysis for HYCU backup licensing

.DESCRIPTION
    This script analyzes Microsoft 365 tenant licensing to determine HYCU backup licensing requirements,
    including user counts, license tiers, storage entitlements, and excess capacity calculations.

.PARAMETER TenantData
    JSON data from the main sizing script

.PARAMETER HYCUEntitlementPerUserGB
    HYCU entitlement per licensed user in GB (default: 50)

.PARAMETER ArchiveThresholdPercentage
    Archive mailbox threshold percentage (default: 20%)

.PARAMETER OutputPath
    Custom output directory for reports

.EXAMPLE
    .\Get-HYCULicensingAnalysis.ps1 -TenantData $sizingData

.EXAMPLE
    .\Get-HYCULicensingAnalysis.ps1 -TenantData $sizingData -HYCUEntitlementPerUserGB 100 -ArchiveThresholdPercentage 25

.NOTES
    Author: HYCU
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [hashtable]$TenantData,
    
    [Parameter(Mandatory = $false)]
    [int]$HYCUEntitlementPerUserGB = 50,
    
    [Parameter(Mandatory = $false)]
    [decimal]$ArchiveThresholdPercentage = 20.0,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "."
)

# Function to write colored output
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
    
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

# Function to analyze licensing requirements
function Get-LicensingRequirements {
    param(
        [hashtable]$TenantData,
        [int]$EntitlementPerUserGB,
        [decimal]$ArchiveThreshold
    )
    
    $totalLicensedUsers = $TenantData.LicensingInfo.TotalLicensedUsers
    $currentUsageGB = $TenantData.GrowthAnalysis.CurrentTotalSizeGB
    $totalHYCUEntitlementGB = $totalLicensedUsers * $EntitlementPerUserGB
    
    # Calculate excess capacity
    $excessCapacityGB = [math]::Max(0, $currentUsageGB - $totalHYCUEntitlementGB)
    $additionalLicensesNeeded = if ($excessCapacityGB -gt 0) { [math]::Ceiling($excessCapacityGB / $EntitlementPerUserGB) } else { 0 }
    
    # Analyze archive mailboxes
    $mailboxAnalysis = $TenantData.LicensingInfo.MailboxAnalysis
    $archiveThreshold = $mailboxAnalysis.TotalMailboxes * ($ArchiveThreshold / 100)
    $excessArchiveMailboxes = [math]::Max(0, $mailboxAnalysis.ArchiveMailboxes - $archiveThreshold)
    $additionalLicensesForArchives = if ($excessArchiveMailboxes -gt 0) { [math]::Ceiling($excessArchiveMailboxes / $EntitlementPerUserGB) } else { 0 }
    
    # Calculate total additional licenses needed
    $totalAdditionalLicenses = $additionalLicensesNeeded + $additionalLicensesForArchives
    
    return @{
        TotalLicensedUsers = $totalLicensedUsers
        CurrentUsageGB = $currentUsageGB
        HYCUEntitlementGB = $totalHYCUEntitlementGB
        ExcessCapacityGB = $excessCapacityGB
        AdditionalLicensesForCapacity = $additionalLicensesNeeded
        ArchiveAnalysis = @{
            TotalMailboxes = $mailboxAnalysis.TotalMailboxes
            ArchiveMailboxes = $mailboxAnalysis.ArchiveMailboxes
            ArchivePercentage = $mailboxAnalysis.ArchivePercentage
            ArchiveThreshold = $archiveThreshold
            ExcessArchiveMailboxes = $excessArchiveMailboxes
            AdditionalLicensesForArchives = $additionalLicensesForArchives
        }
        TotalAdditionalLicenses = $totalAdditionalLicenses
        LicensingCompliance = $excessCapacityGB -eq 0 -and $excessArchiveMailboxes -eq 0
    }
}

# Function to generate licensing recommendations
function Get-LicensingRecommendations {
    param(
        [hashtable]$LicensingRequirements
    )
    
    $recommendations = @()
    
    # Capacity-based recommendations
    if ($LicensingRequirements.ExcessCapacityGB -gt 0) {
        $recommendations += @{
            Type = "Capacity"
            Priority = "High"
            Title = "Storage Capacity Exceeded"
            Description = "Current usage exceeds HYCU entitlement by $($LicensingRequirements.ExcessCapacityGB) GB"
            Action = "Purchase $($LicensingRequirements.AdditionalLicensesForCapacity) additional HYCU licenses"
            Impact = "Prevents backup failures and ensures compliance"
        }
    }
    
    # Archive-based recommendations
    if ($LicensingRequirements.ArchiveAnalysis.ExcessArchiveMailboxes -gt 0) {
        $recommendations += @{
            Type = "Archive"
            Priority = "Medium"
            Title = "Archive Mailbox Threshold Exceeded"
            Description = "Archive mailboxes exceed 20% threshold by $($LicensingRequirements.ArchiveAnalysis.ExcessArchiveMailboxes) mailboxes"
            Action = "Purchase $($LicensingRequirements.ArchiveAnalysis.AdditionalLicensesForArchives) additional HYCU licenses for archive compliance"
            Impact = "Ensures proper licensing for archive mailboxes"
        }
    }
    
    # Growth-based recommendations
    $growthProjections = $TenantData.GrowthAnalysis.GrowthProjections
    $maxProjectedGrowth = ($growthProjections.Values | Measure-Object -Maximum).Maximum
    $projectedExcessCapacity = [math]::Max(0, $maxProjectedGrowth - $LicensingRequirements.HYCUEntitlementGB)
    
    if ($projectedExcessCapacity -gt 0) {
        $projectedAdditionalLicenses = [math]::Ceiling($projectedExcessCapacity / 50)
        $recommendations += @{
            Type = "Growth"
            Priority = "Low"
            Title = "Future Growth Planning"
            Description = "Projected growth may require additional licenses"
            Action = "Consider purchasing $projectedAdditionalLicenses additional HYCU licenses for future growth"
            Impact = "Prevents licensing issues as tenant grows"
        }
    }
    
    # Compliance recommendations
    if ($LicensingRequirements.LicensingCompliance) {
        $recommendations += @{
            Type = "Compliance"
            Priority = "Info"
            Title = "Licensing Compliance"
            Description = "Current licensing is compliant with HYCU requirements"
            Action = "No immediate action required"
            Impact = "Maintains current backup capabilities"
        }
    }
    
    return $recommendations
}

# Function to calculate cost impact
function Get-CostImpact {
    param(
        [hashtable]$LicensingRequirements,
        [decimal]$CostPerLicense = 0
    )
    
    if ($CostPerLicense -eq 0) {
        return @{
            AdditionalLicensesCost = 0
            CostPerLicense = 0
            TotalCostImpact = 0
            CostJustification = "Cost per license not specified"
        }
    }
    
    $additionalLicensesCost = $LicensingRequirements.TotalAdditionalLicenses * $CostPerLicense
    
    return @{
        AdditionalLicensesCost = $additionalLicensesCost
        CostPerLicense = $CostPerLicense
        TotalCostImpact = $additionalLicensesCost
        CostJustification = "Based on $($LicensingRequirements.TotalAdditionalLicenses) additional licenses at $$CostPerLicense each"
    }
}

# Function to generate HTML licensing report
function New-LicensingHTMLReport {
    param(
        [hashtable]$LicensingRequirements,
        [array]$Recommendations,
        [hashtable]$CostImpact
    )
    
    $reportPath = Join-Path $OutputPath "HYCU-Licensing-Analysis-$(Get-Date -Format 'yyyy-MM-dd-HHmm').html"
    
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>HYCU Licensing Analysis Report</title>
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
        .recommendation { background: #e8f5e8; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745; margin: 10px 0; }
        .warning { background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107; margin: 10px 0; }
        .error { background: #f8d7da; padding: 15px; border-radius: 8px; border-left: 4px solid #dc3545; margin: 10px 0; }
        .info { background: #d1ecf1; padding: 15px; border-radius: 8px; border-left: 4px solid #17a2b8; margin: 10px 0; }
        .footer { text-align: center; padding: 20px; color: #666; border-top: 1px solid #eee; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📋 HYCU Licensing Analysis</h1>
            <p>Microsoft 365 Licensing Analysis for HYCU Backup Licensing</p>
            <p>Generated on: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </div>
        
        <div class="content">
            <div class="section">
                <h2>📊 Licensing Overview</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.TotalLicensedUsers)</div>
                        <div class="metric-label">Total Licensed Users</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.HYCUEntitlementGB) GB</div>
                        <div class="metric-label">HYCU Entitlement</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.CurrentUsageGB) GB</div>
                        <div class="metric-label">Current Usage</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.TotalAdditionalLicenses)</div>
                        <div class="metric-label">Additional Licenses Needed</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>📈 Capacity Analysis</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.ExcessCapacityGB) GB</div>
                        <div class="metric-label">Excess Capacity</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.AdditionalLicensesForCapacity)</div>
                        <div class="metric-label">Licenses for Capacity</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.ArchiveAnalysis.ExcessArchiveMailboxes)</div>
                        <div class="metric-label">Excess Archive Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.ArchiveAnalysis.AdditionalLicensesForArchives)</div>
                        <div class="metric-label">Licenses for Archives</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>📋 Archive Mailbox Analysis</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.ArchiveAnalysis.TotalMailboxes)</div>
                        <div class="metric-label">Total Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.ArchiveAnalysis.ArchiveMailboxes)</div>
                        <div class="metric-label">Archive Mailboxes</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.ArchiveAnalysis.ArchivePercentage)%</div>
                        <div class="metric-label">Archive Percentage</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$($LicensingRequirements.ArchiveAnalysis.ArchiveThreshold)</div>
                        <div class="metric-label">20% Threshold</div>
                    </div>
                </div>
            </div>
            
            <div class="section">
                <h2>💡 Recommendations</h2>
"@

        foreach ($rec in $Recommendations) {
            $cardClass = switch ($rec.Priority) {
                "High" { "error" }
                "Medium" { "warning" }
                "Low" { "info" }
                "Info" { "recommendation" }
            }
            
            $html += @"
                <div class="$cardClass">
                    <strong>$($rec.Title)</strong> ($($rec.Type) - $($rec.Priority) Priority)
                    <p>$($rec.Description)</p>
                    <p><strong>Action:</strong> $($rec.Action)</p>
                    <p><strong>Impact:</strong> $($rec.Impact)</p>
                </div>
"@
        }

        $html += @"
            </div>
            
            <div class="section">
                <h2>💰 Cost Impact</h2>
                <div class="metric-grid">
                    <div class="metric-card">
                        <div class="metric-value">$$($CostImpact.AdditionalLicensesCost)</div>
                        <div class="metric-label">Additional Licenses Cost</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$($CostImpact.CostPerLicense)</div>
                        <div class="metric-label">Cost Per License</div>
                    </div>
                    <div class="metric-card">
                        <div class="metric-value">$$($CostImpact.TotalCostImpact)</div>
                        <div class="metric-label">Total Cost Impact</div>
                    </div>
                </div>
                <div class="info">
                    <strong>Cost Justification:</strong> $($CostImpact.CostJustification)
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p>Generated by HYCU M365 Sizing Tool v1.0 | For backup licensing analysis</p>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $reportPath -Encoding UTF8
    return $reportPath
}

# Main execution
try {
    Write-ColorOutput "Starting HYCU Licensing Analysis..." "HEADER"
    Write-ColorOutput "=================================" "HEADER"
    
    # Analyze licensing requirements
    $licensingRequirements = Get-LicensingRequirements -TenantData $TenantData -EntitlementPerUserGB $HYCUEntitlementPerUserGB -ArchiveThreshold $ArchiveThresholdPercentage
    
    # Generate recommendations
    $recommendations = Get-LicensingRecommendations -LicensingRequirements $licensingRequirements
    
    # Calculate cost impact
    $costImpact = Get-CostImpact -LicensingRequirements $licensingRequirements
    
    # Generate HTML report
    $reportPath = New-LicensingHTMLReport -LicensingRequirements $licensingRequirements -Recommendations $recommendations -CostImpact $costImpact
    
    # Output summary
    Write-ColorOutput "`nLicensing Analysis Complete:" "SUCCESS"
    Write-ColorOutput "Total Licensed Users: $($licensingRequirements.TotalLicensedUsers)" "INFO"
    Write-ColorOutput "HYCU Entitlement: $($licensingRequirements.HYCUEntitlementGB) GB" "INFO"
    Write-ColorOutput "Current Usage: $($licensingRequirements.CurrentUsageGB) GB" "INFO"
    Write-ColorOutput "Excess Capacity: $($licensingRequirements.ExcessCapacityGB) GB" "INFO"
    Write-ColorOutput "Additional Licenses Needed: $($licensingRequirements.TotalAdditionalLicenses)" "INFO"
    
    if ($licensingRequirements.LicensingCompliance) {
        Write-ColorOutput "Licensing Status: COMPLIANT" "SUCCESS"
    } else {
        Write-ColorOutput "Licensing Status: NON-COMPLIANT" "WARNING"
    }
    
    Write-ColorOutput "`nReport saved to: $reportPath" "SUCCESS"
    Write-ColorOutput "=================================" "HEADER"
    
    return @{
        LicensingRequirements = $licensingRequirements
        Recommendations = $recommendations
        CostImpact = $costImpact
        ReportPath = $reportPath
    }
}
catch {
    Write-ColorOutput "Licensing analysis failed: $($_.Exception.Message)" "ERROR"
    throw
}
