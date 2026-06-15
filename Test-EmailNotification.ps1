<#
.SYNOPSIS
Test script for Yardstick email notification formatting.

.DESCRIPTION
Populates the application trackers with realistic data designed to exercise the
new email report formatting:
  - Long display names / versions that would overflow the report width
  - Dependent apps in every status, including "Failed (long error)" entries that
    should render in scrollable code boxes
  - Multi-line stack-trace style errors on failed apps

By default the email is opened in Outlook for visual inspection (not sent). Pass
-Send to actually deliver to the recipients in Preferences.yaml, or -SaveHtml to
also dump the rendered body to disk for browser preview.

.PARAMETER TestOutlook
Only check Outlook COM availability.

.PARAMETER Send
Actually send the email instead of opening it for preview.

.PARAMETER SaveHtml
Also write the rendered HTML to TestEmailPreview.html in the script directory.
#>

param(
    [Switch]$TestOutlook,
    [Switch]$Send,
    [Switch]$SaveHtml
)

Import-Module "$PSScriptRoot\Modules\YardstickSupport.psm1" -Force
Import-Module powershell-yaml -Force

$Global:LOG_LOCATION = $PSScriptRoot
$Global:LOG_FILE = "EmailTest.log"
Write-Host -Init

Write-Host "Starting Yardstick Email Notification Test"

try {
    $prefs = Get-Content "$PSScriptRoot\Preferences.yaml" | ConvertFrom-Yaml
    Write-Host "Preferences loaded successfully"

    if ($TestOutlook) {
        $outlookAvailable = Test-OutlookAvailability
        Write-Host "Outlook COM object available: $outlookAvailable"
        return
    }

    Initialize-ApplicationTracker

    # --- Successful apps -----------------------------------------------------

    Add-SuccessfulApplication `
        -ApplicationId "AcrobatReaderDC" `
        -DisplayName "Adobe Acrobat Reader DC" `
        -Version "2024.001.20643" `
        -Action "Updated" `
        -Dependents @{
            "Visual C++ Runtime" = "Updated"
            "Fonts Package"      = "Skipped (blacklisted)"
        }

    # Long-ish dependent error to verify the scrollable code box on a successful row.
    $longDepError = "The remote server returned an error: (504) Gateway Timeout. At line:1 char:1 + Add-IntuneWin32AppDependency -ID 00000000-aaaa-bbbb-cccc-111122223333 -Dependency ..."
    Add-SuccessfulApplication `
        -ApplicationId "GoogleChrome" `
        -DisplayName "Google Chrome" `
        -Version "120.0.6099.216" `
        -Action "Force Updated" `
        -Dependents @{
            "Google Update Helper"                                            = "Added"
            "PowerShell Prerequisite Module With An Unusually Long Name v2"   = "Updated"
            "Legacy Dependent App Pinned By Compliance"                       = "Skipped (link updates disabled for recipe)"
            "Broken Dependent App That Refuses To Re-link"                    = "Failed ($longDepError)"
        }

    # App with a very long display name + version to force horizontal overflow.
    Add-SuccessfulApplication `
        -ApplicationId "VeryWideApp" `
        -DisplayName "JetBrains IntelliJ IDEA Ultimate Edition With An Extremely Long Marketing Suffix" `
        -Version "2025.3.1.0-EAP-build-9876543210-x64-with-bundled-jbr" `
        -Action "Updated"

    # --- Failed apps ---------------------------------------------------------

    $multiLineError = @"
System.Net.WebException: The remote name could not be resolved: 'downloads.example.com'
   at System.Net.HttpWebRequest.GetResponse()
   at Invoke-WebRequest, <No file>: line 42
   at <ScriptBlock>, G:\Intune\Yardstick\Recipes\example.yaml: line 12
Inner exception:
   System.Net.Sockets.SocketException (11001): No such host is known.
"@

    Add-FailedApplication `
        -ApplicationId "BrokenDownloader" `
        -DisplayName "App With A Long Stack Trace Error" `
        -Version "3.14.159" `
        -ErrorMessage $multiLineError `
        -FailureStage "Download"

    Add-FailedApplication `
        -ApplicationId "ShortFailure" `
        -DisplayName "Simple Failure" `
        -Version "1.0.0" `
        -ErrorMessage "HTTP 404" `
        -FailureStage "Download"

    Add-FailedApplication `
        -ApplicationId "UploadFailure" `
        -DisplayName "Upload Failed." `
        -Version "2.0.0" `
        -ErrorMessage "Upload to Intune failed after 3 attempts: (413) Request Entity Too Large. Response body: {`"error`":{`"code`":`"RequestTooLarge`",`"message`":`"The package exceeds the maximum allowed size of 8 GB.`"}}" `
        -FailureStage "Upload"

    # --- Send / preview ------------------------------------------------------

    $testParams = "-All -NoInteractive (Test-EmailNotification.ps1)"
    $htmlOutPath = if ($SaveHtml) { Join-Path $PSScriptRoot "TestEmailPreview.html" } else { $null }

    if ($Send) {
        Write-Host "Sending test email report..."
        Send-YardstickEmailReport -Preferences $prefs -RunParameters $testParams -HtmlOutputPath $htmlOutPath
    } else {
        Write-Host "Opening test email report in Outlook for preview (use -Send to actually deliver)..."
        Send-YardstickEmailReport -Preferences $prefs -RunParameters $testParams -Preview -HtmlOutputPath $htmlOutPath
    }

    if ($SaveHtml) {
        Write-Host "Browser preview written to $htmlOutPath"
    }

    Write-Host "Email test completed"
}
catch {
    Write-Host "ERROR: Test failed: $_"
    throw
}
