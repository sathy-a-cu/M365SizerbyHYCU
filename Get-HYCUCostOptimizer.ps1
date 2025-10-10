#Requires -Version 5.1

<#
.SYNOPSIS
    HYCU Cost Optimizer - Advanced cost analysis and optimization recommendations for M365 backup

.DESCRIPTION
    This script analyzes Microsoft 365 tenant costs and provides optimization recommendations
    including storage tiering, retention policies, and cost-saving strategies.

.PARAMETER TenantData
    JSON data from the main sizing script

.PARAMETER TargetBudget
    Target annual budget in USD

.PARAMETER OptimizationLevel
    Optimization aggressiveness: Conservative, Moderate, Aggressive

.PARAMETER IncludeROI
    Include ROI analysis for optimization recommendations

.EXAMPLE
    .\Get-HYCUCostOptimizer.ps1 -TenantData $sizingData -TargetBudget 50000 -OptimizationLevel "Moderate"

.NOTES
    Author: HYCU
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [hashtable]$TenantData,
    
    [Parameter(Mandatory = $false)]
    [decimal]$TargetBudget = 0,
    
    [Parameter(Mandatory = $false)]
    [ValidateSet("Conservative", "Moderate", "Aggressive")]
    [string]$OptimizationLevel = "Moderate",
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeROI
)

# Function to analyze current costs
function Get-CurrentCostAnalysis {
    param(
        [hashtable]$TenantData
    )
    
    $totalSizeGB = $TenantData.GrowthAnalysis.CurrentTotalSizeGB
    $userCount = $TenantData.TenantInfo.UserCounts.EnabledUsers
    
    # Current cost assumptions (these would be configurable)
    $currentStorageCostPerGB = 0.12  # $0.12 per GB per month
    $currentUserCost = 3.00  # $3.00 per user per month
    
    $monthlyStorageCost = $totalSizeGB * $currentStorageCostPerGB
    $monthlyUserCost = $userCount * $currentUserCost
    $monthlyTotal = $monthlyStorageCost + $monthlyUserCost
    $annualTotal = $monthlyTotal * 12
    
    return @{
        MonthlyStorageCost = [math]::Round($monthlyStorageCost, 2)
        MonthlyUserCost = [math]::Round($monthlyUserCost, 2)
        MonthlyTotal = [math]::Round($monthlyTotal, 2)
        AnnualTotal = [math]::Round($annualTotal, 2)
        CostPerGB = $currentStorageCostPerGB
        CostPerUser = $currentUserCost
    }
}

# Function to generate storage tiering optimization
function Get-StorageTieringOptimization {
    param(
        [hashtable]$TenantData,
        [string]$OptimizationLevel
    )
    
    $totalSizeGB = $TenantData.GrowthAnalysis.CurrentTotalSizeGB
    
    # Tier definitions with costs per GB per month
    $tiers = @{
        Hot = @{ Cost = 0.15; Description = "Frequently accessed data" }
        Cool = @{ Cost = 0.08; Description = "Infrequently accessed data" }
        Archive = @{ Cost = 0.02; Description = "Long-term archival data" }
    }
    
    # Optimization level determines tier distribution
    $tierDistribution = switch ($OptimizationLevel) {
        "Conservative" {
            @{ Hot = 40; Cool = 40; Archive = 20 }
        }
        "Moderate" {
            @{ Hot = 30; Cool = 35; Archive = 35 }
        }
        "Aggressive" {
            @{ Hot = 20; Cool = 30; Archive = 50 }
        }
    }
    
    $optimizedCost = 0
    $tierBreakdown = @{}
    
    foreach ($tier in $tierDistribution.Keys) {
        $percentage = $tierDistribution[$tier]
        $tierSizeGB = $totalSizeGB * ($percentage / 100)
        $tierCost = $tierSizeGB * $tiers[$tier].Cost
        $optimizedCost += $tierCost
        
        $tierBreakdown[$tier] = @{
            SizeGB = [math]::Round($tierSizeGB, 2)
            Cost = [math]::Round($tierCost, 2)
            Percentage = $percentage
            Description = $tiers[$tier].Description
        }
    }
    
    return @{
        TierBreakdown = $tierBreakdown
        OptimizedMonthlyCost = [math]::Round($optimizedCost, 2)
        OptimizedAnnualCost = [math]::Round($optimizedCost * 12, 2)
    }
}

# Function to generate retention policy optimization
function Get-RetentionPolicyOptimization {
    param(
        [hashtable]$TenantData,
        [string]$OptimizationLevel
    )
    
    $userCount = $TenantData.TenantInfo.UserCounts.EnabledUsers
    $totalSizeGB = $TenantData.GrowthAnalysis.CurrentTotalSizeGB
    
    # Retention policy recommendations based on optimization level
    $retentionPolicies = switch ($OptimizationLevel) {
        "Conservative" {
            @{
                Exchange = "7 years"
                OneDrive = "5 years"
                SharePoint = "7 years"
                Teams = "3 years"
                Description = "Standard compliance retention"
            }
        }
        "Moderate" {
            @{
                Exchange = "5 years"
                OneDrive = "3 years"
                SharePoint = "5 years"
                Teams = "2 years"
                Description = "Balanced retention policy"
            }
        }
        "Aggressive" {
            @{
                Exchange = "3 years"
                OneDrive = "2 years"
                SharePoint = "3 years"
                Teams = "1 year"
                Description = "Cost-optimized retention"
            }
        }
    }
    
    # Calculate storage reduction based on retention policy
    $storageReduction = switch ($OptimizationLevel) {
        "Conservative" { 0.10 }  # 10% reduction
        "Moderate" { 0.25 }       # 25% reduction
        "Aggressive" { 0.40 }     # 40% reduction
    }
    
    $reducedSizeGB = $totalSizeGB * (1 - $storageReduction)
    $sizeReductionGB = $totalSizeGB - $reducedSizeGB
    
    return @{
        RetentionPolicies = $retentionPolicies
        StorageReductionGB = [math]::Round($sizeReductionGB, 2)
        ReducedSizeGB = [math]::Round($reducedSizeGB, 2)
        ReductionPercentage = [math]::Round($storageReduction * 100, 1)
    }
}

# Function to generate compression and deduplication recommendations
function Get-CompressionOptimization {
    param(
        [hashtable]$TenantData
    )
    
    $totalSizeGB = $TenantData.GrowthAnalysis.CurrentTotalSizeGB
    
    # Compression ratios by data type
    $compressionRatios = @{
        Exchange = 0.30  # 70% compression
        OneDrive = 0.20  # 80% compression
        SharePoint = 0.25  # 75% compression
        Teams = 0.35  # 65% compression
    }
    
    # Calculate weighted average compression
    $averageCompression = 0.30  # 70% average compression
    $compressedSizeGB = $totalSizeGB * $averageCompression
    $compressionSavingsGB = $totalSizeGB - $compressedSizeGB
    
    return @{
        CompressionRatio = $averageCompression
        CompressedSizeGB = [math]::Round($compressedSizeGB, 2)
        CompressionSavingsGB = [math]::Round($compressionSavingsGB, 2)
        SavingsPercentage = [math]::Round((1 - $averageCompression) * 100, 1)
    }
}

# Function to calculate ROI for optimizations
function Get-ROIAnalysis {
    param(
        [hashtable]$CurrentCosts,
        [hashtable]$OptimizedCosts,
        [decimal]$ImplementationCost = 5000
    )
    
    $annualSavings = $CurrentCosts.AnnualTotal - $OptimizedCosts.OptimizedAnnualCost
    $roiPercentage = if ($ImplementationCost -gt 0) {
        [math]::Round(($annualSavings / $ImplementationCost) * 100, 2)
    } else {
        0
    }
    
    $paybackPeriod = if ($annualSavings -gt 0) {
        [math]::Round($ImplementationCost / $annualSavings, 1)
    } else {
        0
    }
    
    return @{
        AnnualSavings = [math]::Round($annualSavings, 2)
        ROIPercentage = $roiPercentage
        PaybackPeriodYears = $paybackPeriod
        ImplementationCost = $ImplementationCost
    }
}

# Function to generate cost optimization recommendations
function Get-CostOptimizationRecommendations {
    param(
        [hashtable]$TenantData,
        [string]$OptimizationLevel,
        [decimal]$TargetBudget
    )
    
    $recommendations = @()
    
    # Storage tiering recommendation
    $recommendations += @{
        Category = "Storage Tiering"
        Priority = "High"
        Description = "Implement intelligent storage tiering to move infrequently accessed data to lower-cost tiers"
        PotentialSavings = "20-40%"
        ImplementationEffort = "Medium"
    }
    
    # Compression recommendation
    $recommendations += @{
        Category = "Data Compression"
        Priority = "High"
        Description = "Enable advanced compression and deduplication to reduce storage footprint"
        PotentialSavings = "60-80%"
        ImplementationEffort = "Low"
    }
    
    # Retention policy recommendation
    $retentionSavings = switch ($OptimizationLevel) {
        "Conservative" { "10-15%" }
        "Moderate" { "20-30%" }
        "Aggressive" { "35-50%" }
    }
    
    $recommendations += @{
        Category = "Retention Policy"
        Priority = "Medium"
        Description = "Optimize retention policies based on compliance requirements"
        PotentialSavings = $retentionSavings
        ImplementationEffort = "Medium"
    }
    
    # Backup frequency optimization
    $recommendations += @{
        Category = "Backup Frequency"
        Priority = "Low"
        Description = "Optimize backup frequency based on data criticality and change rates"
        PotentialSavings = "10-20%"
        ImplementationEffort = "Low"
    }
    
    return $recommendations
}

# Main execution
try {
    Write-Host "Starting HYCU Cost Optimizer Analysis..." -ForegroundColor Cyan
    Write-Host "=======================================" -ForegroundColor Cyan
    
    # Analyze current costs
    $currentCosts = Get-CurrentCostAnalysis -TenantData $TenantData
    
    # Generate storage tiering optimization
    $tieringOptimization = Get-StorageTieringOptimization -TenantData $TenantData -OptimizationLevel $OptimizationLevel
    
    # Generate retention policy optimization
    $retentionOptimization = Get-RetentionPolicyOptimization -TenantData $TenantData -OptimizationLevel $OptimizationLevel
    
    # Generate compression optimization
    $compressionOptimization = Get-CompressionOptimization -TenantData $TenantData
    
    # Calculate optimized costs
    $optimizedCosts = @{
        OptimizedAnnualCost = $tieringOptimization.OptimizedAnnualCost
        StorageReductionGB = $retentionOptimization.StorageReductionGB
        CompressionSavingsGB = $compressionOptimization.CompressionSavingsGB
    }
    
    # Calculate ROI if requested
    $roiAnalysis = if ($IncludeROI) {
        Get-ROIAnalysis -CurrentCosts $currentCosts -OptimizedCosts $optimizedCosts
    } else {
        @{ AnnualSavings = 0; ROIPercentage = 0; PaybackPeriodYears = 0 }
    }
    
    # Generate recommendations
    $recommendations = Get-CostOptimizationRecommendations -TenantData $TenantData -OptimizationLevel $OptimizationLevel -TargetBudget $TargetBudget
    
    # Create comprehensive optimization report
    $optimizationReport = @{
        CurrentCosts = $currentCosts
        TieringOptimization = $tieringOptimization
        RetentionOptimization = $retentionOptimization
        CompressionOptimization = $compressionOptimization
        OptimizedCosts = $optimizedCosts
        ROIAnalysis = $roiAnalysis
        Recommendations = $recommendations
        TargetBudget = $TargetBudget
        BudgetCompliance = if ($TargetBudget -gt 0) { $optimizedCosts.OptimizedAnnualCost -le $TargetBudget } else { $true }
        GeneratedOn = Get-Date
    }
    
    # Output results
    Write-Host "`nCost Optimization Analysis Complete:" -ForegroundColor Green
    Write-Host "Current Annual Cost: $$($currentCosts.AnnualTotal)" -ForegroundColor Yellow
    Write-Host "Optimized Annual Cost: $$($optimizedCosts.OptimizedAnnualCost)" -ForegroundColor Yellow
    
    if ($roiAnalysis.AnnualSavings -gt 0) {
        Write-Host "Annual Savings: $$($roiAnalysis.AnnualSavings)" -ForegroundColor Green
        Write-Host "ROI: $($roiAnalysis.ROIPercentage)%" -ForegroundColor Green
        Write-Host "Payback Period: $($roiAnalysis.PaybackPeriodYears) years" -ForegroundColor Green
    }
    
    if ($TargetBudget -gt 0) {
        if ($optimizationReport.BudgetCompliance) {
            Write-Host "Budget Status: Within target budget" -ForegroundColor Green
        } else {
            Write-Host "Budget Status: Over target budget by $$($optimizedCosts.OptimizedAnnualCost - $TargetBudget)" -ForegroundColor Red
        }
    }
    
    # Save optimization report
    $outputPath = Join-Path $PSScriptRoot "HYCU-Cost-Optimization-$(Get-Date -Format 'yyyy-MM-dd-HHmm').json"
    $optimizationReport | ConvertTo-Json -Depth 10 | Out-File -FilePath $outputPath -Encoding UTF8
    
    Write-Host "`nOptimization report saved to: $outputPath" -ForegroundColor Green
    Write-Host "=======================================" -ForegroundColor Cyan
    
    return $optimizationReport
}
catch {
    Write-Error "Failed to generate cost optimization analysis: $($_.Exception.Message)"
    throw
}
