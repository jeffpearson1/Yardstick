# Writing Yardstick Recipes

This guide provides comprehensive documentation for creating new application recipes in Yardstick.

## Quick Start

1. Copy the appropriate template from the `Templates/` folder:
   - `MSITemplate.yaml` for MSI installers
   - `EXETemplate.yaml` for EXE installers or portable applications
2. Rename the file to match your application name (e.g., `googlechrome.yaml`)
3. Fill in the required fields based on the sections below
4. Test your recipe with `.\Yardstick.ps1 -ApplicationId YourAppName -NoEmail`

## Recipe File Naming

The recipe filename (without `.yaml`) becomes the application ID. The filename must:
- Match the `id` field in the YAML file
- Use PascalCase convention (e.g., `GoogleChrome`, `7zip`)
- Be unique across all recipes

## Core Required Fields

### Basic Identification

**`id`** (string) - The application identifier. Must match the YAML filename without extension.
- Example: `id: GoogleChrome`

**`displayName`** (string) - Human-readable application name shown in Intune and Company Portal. Do not include version numbers.
- Example: `displayName: Google Chrome`

**`publisher`** (string) - The software vendor or publisher name.
- Example: `publisher: Google`

**`description`** (string) - Brief description of the application. Supports multi-line strings and Markdown.
- Example: `description: A web browser by Google`

**`iconFile`** (string) - Filename of the icon (PNG or JPG, max 512x512) stored in the `Icons/` folder.
- Example: `iconFile: GoogleChrome.png`

### Version and Download Configuration

**`version`** (string) - Application version. Must be set dynamically in `preDownloadScript` or `downloadScript`.
- Set via: `$version = "1.2.3"`
- Not typically defined in YAML unless using a static version

**`fileName`** (string) - Name of the downloaded file. Can be static in YAML or dynamic in scripts.
- Example: `fileName: googlechromestandaloneenterprise64.msi`
- Set dynamically via: `$fileName = "installer.msi"`

**`url`** (string) - Download URL. Can be static in YAML or dynamic in scripts.
- Static example: `url: https://example.com/download/app.msi`
- Dynamic example: Set `$url` in `preDownloadScript` or `downloadScript`

**`urlRedirects`** (boolean) - If `true`, follows redirect chain to get final URL. Useful for `/latest` style URLs.
- Example: `urlRedirects: true`
- Default: `false`

### File Type

**`fileType`** (string) - Type of installer package.
- Accepted values: `msi`, `EXE`, `zip`
- Example: `fileType: msi`

## Script Blocks

Scripts execute in the following order and allow dynamic configuration:

### `preDownloadScript`

Executes before any download processing. Use this to:
- Scrape version numbers from vendor websites
- Construct dynamic download URLs
- Set version-dependent variables

**Must always set:**
- `$version` - The application version
- `$fileDetectionVersion` - Version string for detection rules (often same as `$version`)

**Commonly sets:**
- `$url` - Download URL when not static
- `$fileName` - When filename includes version or is dynamic

**Example (7-Zip):**
```yaml
preDownloadScript: |
  $html = $(Invoke-WebRequest "https://www.7-zip.org/download.html").content
  $html -match "a\/7z[0-9]{1,6}-x64\.msi" | Out-Null
  $url = "https://www.7-zip.org/$($matches[0])"
  $fileName = $($matches[0] -split "/")[1]
  $html -match "[0-9]{2}\.[0-9]{2}" | Out-Null
  $version = $matches[0]
  $fileDetectionVersion = $version
```

### `downloadScript`

Custom download logic. Yardstick uses BITS transfer by default, but this script allows:
- Custom authentication
- Cookie handling
- Multi-file downloads
- Selenium-based automation

**Must set:**
- `$version`
- `$fileName`
- Perform the actual download to `$BuildSpace`

### `postDownloadScript`

Executes after download completes. Use for:
- Extracting archives
- Modifying installer files
- Adding additional files to the package
- Post-processing downloaded content

**Example:**
```yaml
postDownloadScript: |
  Expand-Archive -Path "$BuildSpace\$fileName" -DestinationPath "$BuildSpace\extracted"
  Copy-Item "C:\Scripts\config.xml" -Destination "$BuildSpace"
```

### `postRunScript`

Executes after the Intune upload completes. Use for:
- Cleanup operations
- Logging
- Notifications
- Any post-deployment tasks

## Installation and Uninstallation

### MSI Installers

**`installScript`** (string) - MSI install command.
- Standard: `installScript: msiexec /i <filename> /qn /norestart`
- The `<filename>` placeholder is automatically replaced

**`uninstallScript`** (string) - MSI uninstall command.
- Standard: `uninstallScript: msiexec /x <productcode> /qn /norestart`
- The `<productcode>` placeholder is automatically populated from the MSI

### EXE Installers

**`installScript`** (string) - EXE install command with silent switches.
- Example: `installScript: <filename> /S`
- Example: `installScript: <filename> /VERYSILENT /NORESTART`

**`uninstallScript`** (string) - EXE uninstall command.
- Example: `uninstallScript: '"C:\Program Files\App\uninstall.exe" /S'`
- Use proper escaping for paths with spaces

### PowerShell Install Scripts

**`powerShellInstallScript`** (string) - PowerShell script content for complex installations.
- Automatically creates `install.ps1` in the package
- Auto-generates `installScript` as: `powershell.exe -noprofile -executionpolicy bypass -file .\install.ps1`

**`powerShellUninstallScript`** (string) - PowerShell script content for complex uninstallations.
- Automatically creates `uninstall.ps1` in the package
- Auto-generates `uninstallScript` as: `powershell.exe -noprofile -executionpolicy bypass -file .\uninstall.ps1`

**Example (QGIS):**
```yaml
powershellInstallScript: | 
  Invoke-WebRequest -Uri "https://github.com/google/fonts/raw/main/OpenSans.ttf" -UseBasicParsing
  if (Test-Path -Path "$($env:ProgramData)\Microsoft\Windows\Start Menu\Programs\QGIS*") {
    Remove-Item -Path "$($env:ProgramData)\Microsoft\Windows\Start Menu\Programs\QGIS*" -Recurse -Force
  }
  Start-Process -FilePath "msiexec" -ArgumentList "/i `"<filename>`" INSTALLDIR=`"$env:ProgramFiles\QGIS`" /qn /norestart" -Wait
```

## Detection Rules

Yardstick supports four detection types: `msi`, `file`, `registry`, and `script`.

### MSI Detection

**`detectionType: msi`**

The simplest detection method. Automatically extracts and uses the MSI ProductCode.

**Required fields:**
- `detectionType: msi`

**Optional fields:**
- `softwareName` (string) - Display name filter for Add/Remove Programs

**Example:**
```yaml
detectionType: msi
```

### File Detection

**`detectionType: file`**

Detects based on file existence, version, date, or size.

**Required fields:**
- `detectionType: file`
- `fileDetectionPath` (string) - Full folder path where the file exists
- `fileDetectionName` (string) - Filename to detect
- `fileDetectionMethod` (string) - Detection method: `version`, `exists`, `dateModified`, `dateCreated`, `size`

**Version detection fields:**
- `fileDetectionVersion` (string) - Expected version (set in `preDownloadScript`)
- `fileDetectionOperator` (string) - Comparison operator: `equal`, `notEqual`, `greaterThan`, `greaterThanOrEqual`, `lessThan`, `lessThanOrEqual`

**Date detection fields:**
- `fileDetectionDateTime` (datetime) - Date/time to compare against
- `fileDetectionOperator` (string) - Same operators as version

**Size detection fields:**
- `fileDetectionValue` (integer) - File size in bytes
- `fileDetectionOperator` (string) - Same operators as version

**Example (Notepad++):**
```yaml
detectionType: file
fileDetectionPath: 'C:\Program Files\Notepad++'
fileDetectionName: notepad++.exe
fileDetectionMethod: version
fileDetectionVersion:  # Set in preDownloadScript
fileDetectionOperator: equal
```

### Registry Detection

**`detectionType: registry`**

Detects based on registry key/value existence or value content.

**Required fields:**
- `detectionType: registry`
- `registryDetectionMethod` (string) - Detection method: `exists`, `notExists`, `string`, `integer`, `version`
- `registryDetectionKey` (string) - Full registry path including hive

**Value-based detection fields:**
- `registryDetectionValueName` (string) - Name of the registry value
- `registryDetectionValue` (string/int) - Expected value
- `registryDetectionOperator` (string) - Comparison operator (same as file detection)

**Example (Firefox):**
```yaml
detectionType: registry
registryDetectionMethod: exists
registryDetectionKey: 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox <version> (x64 en-US)'
```

**Note:** The `<version>` placeholder in registry keys is automatically replaced.

### Script Detection

**`detectionType: script`**

Custom PowerShell detection script for complex scenarios.

**Required fields:**
- `detectionType: script`
- `detectionScript` (string) - PowerShell script content

**Optional fields:**
- `detectionScriptFileExtension` (string) - Script file extension (default from preferences)
- `detectionScriptRunAs32Bit` (boolean) - Run in 32-bit context (default from preferences)
- `detectionScriptEnforceSignatureCheck` (boolean) - Require signed script (default from preferences)

**Example:**
```yaml
detectionType: script
detectionScript: |
  $version = (Get-ItemProperty 'HKLM:\SOFTWARE\MyApp').Version
  if ($version -eq "1.2.3") { exit 0 } else { exit 1 }
```

## Installation Behavior Settings

These settings control how Intune handles installation, or they can be set as defaults in `preferences.yaml`.

**`installExperience`** (string) - Installation context.
- Values: `system`, `user`
- Default: From `preferences.yaml`
- Example: `installExperience: system`

**`restartBehavior`** (string) - Restart handling.
- Values: `allow`, `suppress`, `force`, `basedOnReturnCode`
- Default: From `preferences.yaml`
- Example: `restartBehavior: suppress`

**`maximumInstallationTimeInMinutes`** (integer) - Installation timeout.
- Default: From `preferences.yaml`
- Example: `maximumInstallationTimeInMinutes: 60`

**`allowUserUninstall`** (boolean) - Allow users to uninstall.
- Default: From `preferences.yaml`
- Example: `allowUserUninstall: false`

## System Requirements

**`minOSVersion`** (string) - Minimum Windows version required.
- Values: `W10_1607`, `W10_1703`, `W10_1709`, `W10_1803`, `W10_1809`, `W10_1903`, `W10_1909`, `W10_2004`, `W10_20H2`, `W10_21H1`, `W10_21H2`, `W10_22H2`, `W11_21H2`, `W11_22H2`, `W11_23H2`, `W11_24H2`
- Default: From `preferences.yaml`
- Example: `minOSVersion: W10_20H2`

**`architecture`** (string) - Target architecture.
- Values: `x64`, `x86`, `arm64`, `none`
- Default: From `preferences.yaml`
- Example: `architecture: x64`

**`is32BitApp`** (boolean) - Indicates if application is 32-bit.
- Default: From `preferences.yaml`
- Example: `is32BitApp: false`

## Deployment Configuration

**`scopeTags`** (array) - Intune scope tags for RBAC.
- Default: From `preferences.yaml`
- Example:
  ```yaml
  scopeTags:
    - My Scope Tag
    - My Scope Tag 2
  ```

**`owner`** (string) - Application owner.
- Default: From `preferences.yaml`
- Example: `owner: IT Department`

**`defaultDeploymentGroups`** (array) - Azure AD groups for automatic assignment.
- Default: From `preferences.yaml`
- Example:
  ```yaml
  defaultDeploymentGroups:
    - All Users
    - IT Pilot Group
  ```

**`availableGroups`** (array) - Groups for "Available" assignment intent.
- Default: From `preferences.yaml`

**`requiredGroups`** (array) - Groups for "Required" assignment intent.
- Default: From `preferences.yaml`

**`availableDateOffset`** (integer) - Days from now when app becomes available.
- Default: From `preferences.yaml`
- Example: `availableDateOffset: 0`

**`deadlineDateOffset`** (integer) - Days from now for required installation deadline.
- Default: From `preferences.yaml`
- Example: `deadlineDateOffset: 7`

## Version Management

**`numVersionsToKeep`** (integer) - Number of app versions to retain in Intune.
- Default: From `preferences.yaml`
- Example: `numVersionsToKeep: 2`
- Note: When `versionLock` is set, this is forced to 1

**`versionLock`** (string) - Lock to specific version pattern using wildcards.
- Pattern: Use `x` for wildcard positions
- Example: `versionLock: 3.11.x` (only deploy Python 3.11.x versions)
- Example: `versionLock: 2024.x.x` (only deploy 2024.x.x versions)

**`displayVersion`** (string) - Override version shown in Intune (optional).
- Useful when file version differs from marketing version

## Dependency Management

**`allowDependentLinkUpdates`** (boolean) - Allow updating dependency links when creating new versions.
- Default: From `preferences.yaml`
- Example: `allowDependentLinkUpdates: true`

**`dependentLinkUpdateEnabled`** (boolean) - Enable dependency link update feature.
- Default: From `preferences.yaml`

**`dependentLinkUpdateRetryCount`** (integer) - Number of retry attempts for dependency updates.
- Default: From `preferences.yaml` (typically 3)

**`dependentLinkUpdateRetryDelaySeconds`** (integer) - Delay between retry attempts.
- Default: From `preferences.yaml` (typically 5)

**`dependentLinkUpdateTimeoutSeconds`** (integer) - Timeout for dependency updates.
- Default: From `preferences.yaml` (typically 60)

**`dependentApplicationBlacklist`** (array) - Application IDs to exclude from dependency updates.
- Default: From `preferences.yaml`
- Example:
  ```yaml
  dependentApplicationBlacklist:
    - LegacyApp
    - TestApp
  ```

## Placeholders

Yardstick automatically replaces placeholders in install/uninstall scripts and detection rules:

- **`<filename>`** - Replaced with actual downloaded filename
- **`<version>`** - Replaced with application version
- **`<productcode>`** - Replaced with MSI ProductCode (MSI packages only)
- **`<fileDetectionVersion>`** - Replaced in detection rules

**Example:**
```yaml
installScript: msiexec /i <filename> /qn /norestart
uninstallScript: msiexec /x <productcode> /qn /norestart
```

Becomes:
```powershell
msiexec /i ChromeSetup.msi /qn /norestart
msiexec /x {8A69D345-D564-463C-AFF1-A69D9E530F96} /qn /norestart
```

## Complete Examples

### MSI Installer Example (Google Chrome)

```yaml
url: https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi
urlredirects: false
preDownloadScript: |
  $VURL = "https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions/all/releases?filter=endtime=none&order_by=fraction%20desc"
  $Content = Invoke-WebRequest $VURL
  $JSON = $Content.content | ConvertFrom-JSON
  $Version = $JSON.releases[0].version
  $fileDetectionVersion = $Version
installScript: msiexec /i <filename> /qn /norestart
uninstallScript: msiexec /x <productcode> /qn /norestart
fileName: googlechromestandaloneenterprise64.msi
iconFile: GoogleChrome.png
id: GoogleChrome
displayName: Google Chrome
publisher: Google
description: A web browser by Google
fileType: msi
detectionType: msi
numVersionsToKeep: 2
scopeTags:
  - My Scope Tag
  - My Scope Tag 2
```

### EXE Installer with File Detection Example (Notepad++)

```yaml
url:
urlredirects: false
preDownloadScript: |
  $html = Invoke-WebRequest "https://notepad-plus-plus.org/downloads/"
  $s = $($html.links | Select-String "Current Version") -replace ".{0,}Current Version "
  $version = ($s.Substring(0, $s.IndexOf(";")) -split "<")[0]
  $url = "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v$version/npp.$version.Installer.x64.exe"
  $fileName = "npp.$version.Installer.x64.exe"
  $fileDetectionVersion = $version
installScript: <filename> /S
uninstallScript: '"C:\Program Files\Notepad++\uninstall.exe" /S'
iconFile: npp.png
id: notepadplusplus
displayName: Notepad++
publisher: Don Ho
description: 'A free source code editor and Notepad replacement'
detectionType: file
fileDetectionPath: 'C:\Program Files\Notepad++'
fileDetectionName: notepad++.exe
fileDetectionMethod: version
fileDetectionOperator: equal
scopeTags:
  - My Scope Tag
  - My Scope Tag 2
```

### Registry Detection Example (Firefox)

```yaml
url: https://download.mozilla.org/?product=firefox-msi-latest-ssl&os=win64&lang=en-US
urlredirects: true
preDownloadScript: |
  $filename = $URL.split("/")[-1].replace("%20", " ")
  $version = ($filename.split(" ")[-1] -Split "\.msi")[0]
installScript: msiexec /i "<filename>" /qn /norestart
uninstallScript: |
  "%ProgramFiles%\Mozilla Firefox\uninstall\helper.exe" -ms
iconFile: firefox.png
id: firefox
displayName: Mozilla Firefox
publisher: The Mozilla Foundation
description: Mozilla Firefox is a free and open-source web browser
detectionType: registry
registryDetectionMethod: exists
registryDetectionKey: 'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Mozilla Firefox <version> (x64 en-US)'
numVersionsToKeep: 2
scopeTags:
  - My Scope Tag
  - My Scope Tag 2
```

## Testing Your Recipe

1. **Test the download and version detection:**
   ```powershell
   .\Yardstick.ps1 -ApplicationId YourAppName -NoEmail
   ```

2. **Test with force flag to recreate existing version:**
   ```powershell
   .\Yardstick.ps1 -ApplicationId YourAppName -Force -NoEmail
   ```

3. **Test repair functionality:**
   ```powershell
   .\Yardstick.ps1 -ApplicationId YourAppName -Repair -NoEmail
   ```

4. **Verify detection rules work correctly in Intune after deployment**

## Common Pitfalls

1. **Forgetting to set `$version`** - Always set in `preDownloadScript` or `downloadScript`
2. **Incorrect file detection version format** - Ensure it matches the actual file properties
3. **Missing icon files** - Icons must exist in `Icons/` folder
4. **Hardcoded paths** - Use placeholders instead of hardcoding filenames
5. **Case sensitivity** - YAML keys are case-sensitive; follow examples carefully
6. **Missing `fileDetectionVersion`** - Required for file-based detection
7. **Wrong detection type** - MSI packages should use `detectionType: msi`, not file detection

## Advanced Topics

### Multi-Architecture Support

For applications that support multiple architectures, create separate recipes:
- `appname.yaml` - Default (x64)
- `appname.arm64.yaml` - ARM64 version

Set the `architecture` field appropriately in each.

### Version Normalization

Some vendors use inconsistent version formats. Normalize in `preDownloadScript`:

```yaml
preDownloadScript: |
  $version = "24.08"
  # Convert to 4-part version for file detection
  $fileDetectionVersion = "$version.0.0"
```

### Custom Download Authentication

For applications requiring authentication, use `downloadScript`:

```yaml
downloadScript: |
  $headers = @{ Authorization = "Bearer $token" }
  Invoke-WebRequest -Uri $url -Headers $headers -OutFile "$BuildSpace\$fileName"
```

### Selenium-Based Downloads

For complex download scenarios (JavaScript-heavy pages, multi-step auth), see `Modules/Custom/AdobeDownloader.psm1` for Selenium patterns.

## Getting Help

- Review existing recipes in `Recipes/` folder for examples
- Check templates in `Templates/` folder
- Refer to main Yardstick documentation in `README.md`
- Review Copilot instructions in `.github/copilot-instructions.md`
