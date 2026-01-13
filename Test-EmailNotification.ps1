<#
.SYNOPSIS
Test script for Yardstick email notification functionality.

.DESCRIPTION
This script tests the email notification components to ensure they work properly
with the Outlook COM object and generate correctly formatted reports.
#>

param(
    [Switch]$TestOutlook,
    [Switch]$TestReport
)

# Import required modules
Import-Module "$PSScriptRoot\Modules\YardstickSupport.psm1" -Force
Import-Module powershell-yaml -Force

# Initialize logging
$Global:LOG_LOCATION = $PSScriptRoot
$Global:LOG_FILE = "EmailTest.log"
Write-Log -Init

Write-Log "Starting Yardstick Email Notification Test"

try {
    # Load preferences
    $prefs = Get-Content "$PSScriptRoot\Preferences.yaml" | ConvertFrom-Yaml
    Write-Log "Preferences loaded successfully"
    
    if ($TestOutlook) {
        Write-Log "Testing Outlook availability..."
        $outlookAvailable = Test-OutlookAvailability
        Write-Log "Outlook COM object available: $outlookAvailable"
    }
    
    if ($TestReport -or -not $TestOutlook) {
        Write-Log "Testing email report generation..."
        
        # Initialize application tracking
        Initialize-ApplicationTracker
        
        # Add some test data
        $acrobatDependents = @{
            "Visual C++ Runtime" = "Updated"
            "Fonts Package"      = "Skipped"
        }
        $chromeDependents = @{
            "Google Update Helper" = "Failed"
            "PowerShell Prereq"    = "Added"
        }
        Add-SuccessfulApplication -ApplicationId "TestApp1" -DisplayName "Adobe Acrobat Reader DC" -Version "2024.001.20643" -Action "Updated" -Dependents $acrobatDependents
        Add-SuccessfulApplication -ApplicationId "TestApp2" -DisplayName "Google Chrome" -Version "120.0.6099.216" -Action "Force Updated" -Dependents $chromeDependents
        Add-FailedApplication -ApplicationId "TestApp3" -DisplayName "Failed Application" -Version "1.0.0" -ErrorMessage "Download failed - HTTP 404" -FailureStage "Download"
        Add-FailedApplication -ApplicationId "TestApp4" -DisplayName "Another Failed App" -Version "2.0.0" -ErrorMessage "Upload to Intune failed" -FailureStage "Upload"
        
        # Test email report
        Write-Log "Generating test email report..."
        $testParams = "-All -NoInteractive"
        Send-YardstickEmailReport -Preferences $prefs -RunParameters $testParams
        
        Write-Log "Email test completed"
    }
}
catch {
    Write-Log "ERROR: Test failed: $_"
    throw
}

Write-Log "Email notification test completed successfully"
