# Yardstick Email Notification Feature

## Overview

The Yardstick email notification feature provides automated reporting of application processing results, delivered through Microsoft Outlook or directly to an SMTP server. This feature tracks both successful and failed application updates and sends a comprehensive HTML-formatted email report.

## Prerequisites

- Proper email configuration in `Local\preferences.yaml` (`.\Install-Yardstick.ps1 -Step Email` sets it up)
- For Outlook delivery (the default):
  - Microsoft Outlook installed and configured
  - Outlook COM automation available (typically available with Outlook desktop installation)
- For SMTP delivery: an SMTP server or relay that accepts mail from the Yardstick host

## Configuration

### Preference settings

The following settings control email notifications:

```yaml
#####################################
# EMAIL NOTIFICATION SETTINGS
#####################################
emailNotificationEnabled: true                    # Enable/disable email notifications
emailDeliveryMethod: outlook                      # "outlook" (default) or "smtp"
emailRecipient: "general@yourorganization.com"    # General list - updated software only
adminEmailRecipient: "admin@yourorganization.com" # Admin - full report incl. failures
emailSubject: "Yardstick Application Update Report" # Email subject line
adminEmailSubject: "Yardstick Application Update Report (Admin)" # Optional
emailSenderName: "Yardstick Automation"           # Display name for sender
emailSendFromAddress: "noreply@yourorganization.com" # From address (required for smtp)
smtpServer: "smtp.yourorganization.com"           # SMTP only
smtpPort: 25                                      # SMTP only (default 25)
smtpUseSsl: false                                 # SMTP only - STARTTLS
smtpCredentialTarget: "Yardstick:Smtp"            # SMTP only - Credential Manager target
```

### Setting Descriptions

- **emailNotificationEnabled**: Boolean flag to enable or disable email notifications
- **emailDeliveryMethod**: `outlook` (default when omitted) or `smtp`
- **emailRecipient**: The general list. Receives the general email, which lists only the software that was updated
- **adminEmailRecipient**: Receives the admin email (everything: run parameters, successes, failures, dependents, auto-update and backup status) and operational alerts such as client secret expiration. When omitted, a single admin-style full report goes to `emailRecipient` instead
- **emailSubject**: The subject line for notification emails
- **adminEmailSubject**: Subject of the admin email. Defaults to `emailSubject` followed by ` (Admin)` when `adminEmailRecipient` is set
- **emailSenderName**: Display name shown as the sender
- **emailSendFromAddress**: "From" address. Optional with Outlook (requires Outlook delegation rights); required with SMTP
- **smtpServer**: Host name of the SMTP server or relay
- **smtpPort**: SMTP port, default 25 (587 is typical for authenticated submission)
- **smtpUseSsl**: Use TLS (STARTTLS) for the SMTP connection
- **smtpCredentialTarget**: Windows Credential Manager target holding the SMTP account, default `Yardstick:Smtp`

## Delivery Methods

### General and admin emails

Each run sends up to two emails. The **admin** email goes to `adminEmailRecipient` and contains the full report. The **general** email goes to `emailRecipient` and lists only applications that were updated (`Updated` or `Force Updated`); maintenance sweeps, repairs and failures are left out, and nothing is sent when no application was updated. Both are sent with one confirmation when `emailPromptBeforeSending` is on.

### Outlook (default)

The report is created as an Outlook mail item over COM and sent from the profile Outlook is signed into. The header logo is attached inline. The HTML refers to it as `cid:<file name>`, and Outlook matches that to the attachment by name. No Content-ID property is set, because reading `Attachment.PropertyAccessor` would raise an Outlook security prompt.

Outlook asks for approval before any program sends mail ("A program is trying to send an e-mail message on your behalf") when it cannot confirm that the antivirus is up to date. It always asks on Windows Server, which has no Security Center to report the antivirus state. `.\Test-YardstickSetup.ps1` warns when this will happen. To turn the prompt off, run `.\Install-Yardstick.ps1 -Step Email` and accept the offer; Windows asks for administrator rights at the end of setup. That sets `ObjectModelGuard` (DWORD) to `2` under `HKLM\SOFTWARE\Microsoft\Office\16.0\Outlook\Security`, or under `HKLM\SOFTWARE\Wow6432Node\...` for 32-bit Office. In Outlook's Trust Center, this is **Programmatic Access > Never warn me about suspicious activity**. The setting applies to every program on the computer, not only Yardstick. If you would rather not change it, use SMTP delivery instead.

### SMTP

Set `emailDeliveryMethod: smtp` to send with `Send-MailMessage` straight to `smtpServer`. Outlook does not need to be installed, which makes this the better fit for scheduled runs on a server.

- **Authentication**: by default the message is relayed anonymously. If the server requires a login, store the account once on the Yardstick host:

  ```powershell
  .\Set-YardstickCredential.ps1 -Smtp          # prompt for user name and password
  .\Set-YardstickCredential.ps1 -Smtp -Show    # show the stored user name
  .\Set-YardstickCredential.ps1 -Smtp -Remove  # go back to anonymous relay
  ```

  Like the Intune credentials, the account is kept in Windows Credential Manager, encrypted for the Windows user that runs Yardstick, and is never written to `preferences.yaml`.
- **Logo**: `Send-MailMessage` cannot set a Content-ID on an attachment, and many clients block base64 images, so SMTP reports are sent without the header logo.
- **Preview**: SMTP has no compose window, so `-Preview` opens the rendered HTML in the default browser instead.
- Microsoft marks `Send-MailMessage` as obsolete because it cannot guarantee secure connections, but it still works in both Windows PowerShell 5.1 and PowerShell 7. Use `smtpUseSsl: true` whenever the server supports it.

Credential-expiration alerts (see the README) use the same delivery method as the run report.

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
3. Outlook is available (Outlook delivery), or `smtpServer` and `emailSendFromAddress` are set (SMTP delivery)

### Manual Testing

Use the included test script to validate email functionality:

```powershell
# Test Outlook availability only
.\Test-EmailNotification.ps1 -TestOutlook

# Check SMTP settings, stored credential and TCP connectivity only
.\Test-EmailNotification.ps1 -TestSmtp

# Build a sample report and open it for preview (Outlook, or the browser for SMTP)
.\Test-EmailNotification.ps1

# Build a sample report and actually send it
.\Test-EmailNotification.ps1 -Send
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

#### `Test-YardstickOutlookPromptExpected`
Returns `$true` when Outlook will ask for approval each time Yardstick sends an email. This is the case when `ObjectModelGuard` is `1`, or on Windows Server when it is not `2`. `Set-YardstickOutlookObjectModelGuard` sets it to `2` and needs an elevated session.

#### `Get-YardstickEmailDeliveryMethod`
Returns the configured `emailDeliveryMethod`, lowercased, or `outlook` when unset.

#### `Send-YardstickSmtpMessage`
Sends an HTML message with `Send-MailMessage` using the `smtp*` preferences and the stored SMTP credential, if there is one. Throws on failure.

**Parameters:**
- `Preferences`: Configuration hashtable from `Import-YardstickPreferences`
- `To`: One or more recipient addresses
- `Subject`: Message subject
- `HtmlBody`: HTML message body

#### `Send-YardstickEmailReport`
Generates the email report and sends it through the configured delivery method.

**Parameters:**
- `Preferences`: Configuration hashtable from `Import-YardstickPreferences`
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
   - Verify all required settings in `Local\preferences.yaml`
   - Check for typos in setting names

3. **"Failed to send email report"**
   - Check Outlook configuration
   - Verify network connectivity
   - Ensure proper permissions for COM automation

4. **"Email setting 'smtpServer' is required for SMTP delivery"**
   - Set `smtpServer` and `emailSendFromAddress` when `emailDeliveryMethod` is `smtp`

5. **SMTP send fails (connection refused, authentication required, relay denied)**
   - Run `.\Test-EmailNotification.ps1 -TestSmtp` to confirm the server and port are reachable
   - If the server requires a login, store one with `.\Set-YardstickCredential.ps1 -Smtp`
   - For port 587, set `smtpUseSsl: true`
   - The SMTP credential is readable only by the Windows user that stored it, so store it as the account that runs Yardstick

6. **Outlook asks "A program is trying to send an e-mail message on your behalf" for every report**
   - Outlook cannot confirm the antivirus is up to date. This always happens on Windows Server.
   - Run `.\Install-Yardstick.ps1 -Step Email -EmailMethod outlook -DisableOutlookPrompts`, accept the administrator prompt at the end, then restart Outlook. See [Outlook (default)](#outlook-default) for what this changes.

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
