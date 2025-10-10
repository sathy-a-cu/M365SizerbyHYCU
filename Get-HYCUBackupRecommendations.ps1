#Requires -Version 5.1

<#
.SYNOPSIS
    HYCU Backup Recommendations Generator - Advanced backup planning based on M365 tenant analysis

.DESCRIPTION
    This script analyzes Microsoft 365 tenant data and generates comprehensive backup recommendations
    including RTO/RPO analysis, storage tiering, and disaster recovery planning.

.PARAMETER TenantData
    JSON data from the main sizing script

.PARAMETER BusinessCriticality
    Business criticality level: Low, Medium, High, Critical

.PARAMETER ComplianceRequirements
    Compliance requirements: None, GDPR, HIPAA, SOX, PCI-DSS

.PARAMETER Budget
    Annual backup budget in USD

.EXAMPLE
    .\Get-HYCUBackupRecommendations.ps1 -TenantData $sizingData -BusinessCriticality "High" -ComplianceRequirements "GDPR"

.NOTES
    Author: HYCU
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [hashtable]$TenantData,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Low", "Medium", "High", "Critical")]
    [string]$BusinessCriticality = "Medium",
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("None", "GDPR", "HIPAA", "SOX", "PCI-DSS")]
    [string]$ComplianceRequirements = "None",
    
    [Parameter(Mandatory = $false)]
    [decimal]$Budget = 0
)

# Function to calculate RTO/RPO requirements
function Get-RTORPORequirements {
    param(
        [string]$Criticality,
        [string]$Compliance
    )
    
    $rtoRpo = switch ($Criticality) {
        "Low" { @{ RTO = "24 hours"; RPO = "24 hours"; BackupFrequency = "Daily" } }
        "Medium" { @{ RTO = "8 hours"; RPO = "4 hours"; BackupFrequency = "Every 4 hours" } }
        "High" { @{ RTO = "4 hours"; RPO = "1 hour"; BackupFrequency = "Hourly" } }
        "Critical" { @{ RTO = "1 hour"; RPO = "15 minutes"; BackupFrequency = "Continuous" } }
    }
    
    # Adjust for compliance requirements
    if ($Compliance -ne "None") {
        $rtoRpo.RPO = "15 minutes"
        $rtoRpo.BackupFrequency = "Continuous"
    }
    
    return $rtoRpo
}

# Function to calculate storage tiering recommendations
function Get-StorageTieringRecommendations {
    param(
        [hashtable]$StorageData,
        [string]$Criticality
    )
    
    $recommendations = @{
        HotTier = @{
            Description = "Frequently accessed data"
            Percentage = switch ($Criticality) {
                "Low" { 20 }
                "Medium" { 30 }
                "High" { 40 }
                "Critical" { 50 }
            }
            StorageClass = "Premium SSD"
        }
        CoolTier = @{
            Description = "Infrequently accessed data"
            Percentage = switch ($Criticality) {
                "Low" { 50 }
                "Medium" { 40 }
                "High" { 35 }
                "Critical" { 30 }
            }
            StorageClass = "Standard SSD"
        }
        ArchiveTier = @{
            Description = "Long-term archival data"
            Percentage = switch ($Criticality) {
                "Low" { 30 }
                "Medium" { 30 }
                "High" { 25 }
                "Critical" { 20 }
            }
            StorageClass = "Archive Storage"
        }
    }
    
    return $recommendations
}

# Function to generate compliance-specific recommendations
function Get-ComplianceRecommendations {
    param(
        [string]$Compliance
    )
    
    $recommendations = switch ($Compliance) {
        "GDPR" {
            @{
                DataRetention = "7 years"
                EncryptionRequired = $true
                AuditLogging = $true
                DataResidency = "EU regions only"
                RightToErasure = $true
            }
        }
        "HIPAA" {
            @{
                DataRetention = "6 years"
                EncryptionRequired = $true
                AuditLogging = $true
                DataResidency = "US regions only"
                AccessControls = "Role-based access required"
            }
        }
        "SOX" {
            @{
                DataRetention = "7 years"
                EncryptionRequired = $true
                AuditLogging = $true
                ImmutableStorage = $true
                AccessControls = "Segregation of duties required"
            }
        }
        "PCI-DSS" {
            @{
                DataRetention = "3 years"
                EncryptionRequired = $true
                AuditLogging = $true
                NetworkSegmentation = $true
                AccessControls = "Multi-factor authentication required"
            }
        }
        default {
            @{
                DataRetention = "3 years"
                EncryptionRequired = $false
                AuditLogging = $false
            }
        }
    }
    
    return $recommendations
}

# Function to calculate backup costs
function Get-BackupCostAnalysis {
    param(
        [hashtable]$TenantData,
        [string]$Criticality,
        [string]$Compliance,
        [decimal]$Budget
    )
    
    $totalSizeGB = $TenantData.GrowthAnalysis.CurrentTotalSizeGB
    $userCount = $TenantData.TenantInfo.UserCounts.EnabledUsers
    
    # Base costs per GB per month
    $baseCostPerGB = switch ($Criticality) {
        "Low" { 0.05 }
        "Medium" { 0.08 }
        "High" { 0.12 }
        "Critical" { 0.18 }
    }
    
    # Compliance multiplier
    $complianceMultiplier = switch ($Compliance) {
        "None" { 1.0 }
        "GDPR" { 1.3 }
        "HIPAA" { 1.4 }
        "SOX" { 1.5 }
        "PCI-DSS" { 1.6 }
    }
    
    $monthlyStorageCost = $totalSizeGB * $baseCostPerGB * $complianceMultiplier
    $monthlyUserCost = $userCount * 2.50
    $monthlyTotal = $monthlyStorageCost + $monthlyUserCost
    $annualTotal = $monthlyTotal * 12
    
    # Calculate ROI if budget is provided
    $roi = if ($Budget -gt 0) {
        [math]::Round((($Budget - $annualTotal) / $Budget) * 100, 2)
    } else {
        0
    }
    
    return @{
        MonthlyStorageCost = [math]::Round($monthlyStorageCost, 2)
        MonthlyUserCost = [math]::Round($monthlyUserCost, 2)
        MonthlyTotal = [math]::Round($monthlyTotal, 2)
        AnnualTotal = [math]::Round($annualTotal, 2)
        ROI = $roi
        BudgetCompliance = if ($Budget -gt 0) { $annualTotal -le $Budget } else { $true }
    }
}

# Function to generate disaster recovery recommendations
function Get-DisasterRecoveryRecommendations {
    param(
        [string]$Criticality,
        [string]$Compliance
    )
    
    $drRecommendations = switch ($Criticality) {
        "Low" {
            @{
                PrimarySite = "Single region"
                SecondarySite = "None"
                FailoverTime = "24-48 hours"
                DataSync = "Daily"
                TestingFrequency = "Quarterly"
            }
        }
        "Medium" {
            @{
                PrimarySite = "Single region with ZRS"
                SecondarySite = "Cross-region replication"
                FailoverTime = "4-8 hours"
                DataSync = "Every 4 hours"
                TestingFrequency = "Monthly"
            }
        }
        "High" {
            @{
                PrimarySite = "Multi-AZ deployment"
                SecondarySite = "Active-passive cross-region"
                FailoverTime = "1-2 hours"
                DataSync = "Hourly"
                TestingFrequency = "Bi-weekly"
            }
        }
        "Critical" {
            @{
                PrimarySite = "Multi-AZ with GRS"
                SecondarySite = "Active-active cross-region"
                FailoverTime = "15-30 minutes"
                DataSync = "Real-time"
                TestingFrequency = "Weekly"
            }
        }
    }
    
    return $drRecommendations
}

# Function to generate backup schedule recommendations
function Get-BackupScheduleRecommendations {
    param(
        [hashtable]$TenantData,
        [string]$Criticality
    )
    
    $schedule = @{
        Exchange = switch ($Criticality) {
            "Low" { "Daily at 2 AM" }
            "Medium" { "Every 6 hours" }
            "High" { "Every 2 hours" }
            "Critical" { "Hourly" }
        }
        OneDrive = switch ($Criticality) {
            "Low" { "Daily at 3 AM" }
            "Medium" { "Every 4 hours" }
            "High" { "Every hour" }
            "Critical" { "Every 15 minutes" }
        }
        SharePoint = switch ($Criticality) {
            "Low" { "Daily at 4 AM" }
            "Medium" { "Every 6 hours" }
            "High" { "Every 2 hours" }
            "Critical" { "Hourly" }
        }
        Teams = switch ($Criticality) {
            "Low" { "Daily at 5 AM" }
            "Medium" { "Every 8 hours" }
            "High" { "Every 4 hours" }
            "Critical" { "Every 2 hours" }
        }
    }
    
    return $schedule
}

# Main execution
try {
    Write-Host "Generating HYCU Backup Recommendations..." -ForegroundColor Cyan
    Write-Host "=========================================" -ForegroundColor Cyan
    
    # Calculate RTO/RPO requirements
    $rtoRpo = Get-RTORPORequirements -Criticality $BusinessCriticality -Compliance $ComplianceRequirements
    
    # Calculate storage tiering
    $storageTiering = Get-StorageTieringRecommendations -StorageData $TenantData.StorageData -Criticality $BusinessCriticality
    
    # Get compliance recommendations
    $complianceRecs = Get-ComplianceRecommendations -Compliance $ComplianceRequirements
    
    # Calculate costs
    $costAnalysis = Get-BackupCostAnalysis -TenantData $TenantData -Criticality $BusinessCriticality -Compliance $ComplianceRequirements -Budget $Budget
    
    # Get DR recommendations
    $drRecommendations = Get-DisasterRecoveryRecommendations -Criticality $BusinessCriticality -Compliance $ComplianceRequirements
    
    # Get backup schedule
    $backupSchedule = Get-BackupScheduleRecommendations -TenantData $TenantData -Criticality $BusinessCriticality
    
    # Generate comprehensive recommendations
    $recommendations = @{
        BusinessCriticality = $BusinessCriticality
        ComplianceRequirements = $ComplianceRequirements
        RTORPO = $rtoRpo
        StorageTiering = $storageTiering
        ComplianceRecommendations = $complianceRecs
        CostAnalysis = $costAnalysis
        DisasterRecovery = $drRecommendations
        BackupSchedule = $backupSchedule
        GeneratedOn = Get-Date
    }
    
    # Output recommendations
    Write-Host "`nBackup Recommendations Generated:" -ForegroundColor Green
    Write-Host "Business Criticality: $BusinessCriticality" -ForegroundColor Yellow
    Write-Host "Compliance Requirements: $ComplianceRequirements" -ForegroundColor Yellow
    Write-Host "RTO: $($rtoRpo.RTO)" -ForegroundColor Yellow
    Write-Host "RPO: $($rtoRpo.RPO)" -ForegroundColor Yellow
    Write-Host "Backup Frequency: $($rtoRpo.BackupFrequency)" -ForegroundColor Yellow
    
    if ($costAnalysis.BudgetCompliance) {
        Write-Host "Budget Status: Within budget" -ForegroundColor Green
    } else {
        Write-Host "Budget Status: Over budget by $$($costAnalysis.AnnualTotal - $Budget)" -ForegroundColor Red
    }
    
    # Save recommendations to JSON file
    $outputPath = Join-Path $PSScriptRoot "HYCU-Backup-Recommendations-$(Get-Date -Format 'yyyy-MM-dd-HHmm').json"
    $recommendations | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputPath -Encoding UTF8
    
    Write-Host "`nRecommendations saved to: $outputPath" -ForegroundColor Green
    Write-Host "=========================================" -ForegroundColor Cyan
    
    return $recommendations
}
catch {
    Write-Error "Failed to generate backup recommendations: $($_.Exception.Message)"
    throw
}
