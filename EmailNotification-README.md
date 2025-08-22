# Yardstick Email Notification Feature

## Overview

The Yardstick email notification feature provides automated reporting of application processing results via Microsoft Outlook. This feature tracks both successful and failed application updates and sends a comprehensive HTML-formatted email report.

## Prerequisites

- Microsoft Outlook installed and configured
- Outlook COM automation available (typically available with Outlook desktop installation)
- Proper email configuration in `Preferences.yaml`

## Configuration

### Preferences.yaml Settings

The following settings control email notifications:

```yaml
#####################################
# EMAIL NOTIFICATION SETTINGS
#####################################
emailNotificationEnabled: true                    # Enable/disable email notifications
emailRecipient: "admin@yourorganization.com"      # Recipient email address
emailSubject: "Yardstick Application Update Report" # Email subject line
emailSenderName: "Yardstick Automation"           # Display name for sender
emailSendFromAddress: "noreply@yourorganization.com" # From address (optional)
```

### Setting Descriptions

- **emailNotificationEnabled**: Boolean flag to enable or disable email notifications
- **emailRecipient**: The email address that will receive the reports
- **emailSubject**: The subject line for notification emails
- **emailSenderName**: Display name shown as the sender
- **emailSendFromAddress**: Optional "from" address (requires Outlook delegation rights)

## Features

### Application Tracking

The system automatically tracks:
- **Successful Applications**: Applications that were successfully processed and uploaded to Intune
- **Failed Applications**: Applications that encountered errors during processing

### Email Report Content

The HTML email report includes:
- **Executive Summary**: Count of successful vs failed applications
- **Successful Applications Table**: Detailed list with application names, versions, actions performed, and timestamps
- **Failed Applications Table**: Detailed list with application names, error messages, failure stages, and timestamps
- **Run Information**: Parameters used, execution time, and log file location

### Error Tracking Stages

Failed applications are categorized by failure stage:
- **Configuration**: Issues reading application recipe files
- **Pre-Download Script**: Errors in pre-download script execution
- **Download**: File download failures
- **Download Script**: Custom download script failures
- **Post-Download Script**: Post-download script failures
- **Intune Upload**: Failures during upload to Microsoft Intune
- **General Processing**: Unexpected errors during processing

## Usage

### Automatic Operation

Email notifications are sent automatically at the end of each Yardstick run when:
1. Email notifications are enabled in preferences
2. At least one application was processed (successful or failed)
3. Outlook is available

### Manual Testing

Use the included test script to validate email functionality:

```powershell
# Test Outlook availability only
.\Test-EmailNotification.ps1 -TestOutlook

# Test full email report generation
.\Test-EmailNotification.ps1 -TestReport

# Test both
.\Test-EmailNotification.ps1 -TestOutlook -TestReport
```

## New PowerShell Functions

### YardstickSupport.psm1 Functions

The following functions were added to the YardstickSupport module:

#### `Initialize-ApplicationTracker`
Initializes tracking arrays for successful and failed applications.

#### `Add-SuccessfulApplication`
Records a successful application processing event.

**Parameters:**
- `ApplicationId`: Application identifier
- `DisplayName`: Human-readable application name
- `Version`: Application version
- `Action`: Action performed (Updated, Force Updated, Repaired)

#### `Add-FailedApplication`
Records a failed application processing event.

**Parameters:**
- `ApplicationId`: Application identifier
- `DisplayName`: Human-readable application name (optional)
- `Version`: Application version (optional)
- `ErrorMessage`: Description of the error
- `FailureStage`: Stage where failure occurred

#### `Test-OutlookAvailability`
Tests if Microsoft Outlook COM object is available.

**Returns:** Boolean indicating availability

#### `Send-YardstickEmailReport`
Generates and sends the email report using Outlook COM.

**Parameters:**
- `Preferences`: Configuration hashtable from Preferences.yaml
- `RunParameters`: String describing Yardstick execution parameters

## Implementation Details

### COM Object Management

The system properly manages Outlook COM objects:
- Creates COM objects as needed
- Handles errors gracefully
- Performs garbage collection to prevent memory leaks
- Closes objects properly even on errors

### HTML Email Formatting

The email report uses modern HTML with:
- Responsive CSS styling
- Color-coded sections (success = green, failure = red)
- Tabular data presentation
- Professional corporate styling
- Emoji indicators for better visual recognition

### Error Handling

Comprehensive error handling includes:
- Graceful degradation when Outlook is unavailable
- Configuration validation before attempting to send
- Fallback logging when email fails
- Proper resource cleanup on all code paths

## Troubleshooting

### Common Issues

1. **"Outlook COM object not available"**
   - Ensure Outlook is installed
   - Verify Outlook is configured with an email account
   - Check if Outlook is running or can be started

2. **"Email setting 'X' not configured"**
   - Verify all required settings in Preferences.yaml
   - Check for typos in setting names

3. **"Failed to send email report"**
   - Check Outlook configuration
   - Verify network connectivity
   - Ensure proper permissions for COM automation

### Logging

All email-related activities are logged to the standard Yardstick log file:
- Email configuration validation
- Outlook availability checks
- Successful email transmissions
- Error details and troubleshooting information

## Security Considerations

- Email content may contain sensitive application information
- Consider email encryption for sensitive environments
- Validate recipient addresses to prevent information disclosure
- Monitor COM object usage for security compliance

## Integration with Yardstick Workflow

The email notification system integrates seamlessly:
1. **Initialization**: Application tracking is initialized at script start
2. **Processing**: Success/failure events are recorded throughout processing
3. **Completion**: Email report is generated and sent before cleanup
4. **Cleanup**: All resources are properly disposed

This feature provides valuable visibility into Yardstick operations without requiring manual log file review.
