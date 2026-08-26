# IntuneWin32App Removal Handoff

## Executive summary

This branch replaces Yardstick's external `IntuneWin32App` dependency with repository-owned PowerShell modules that authenticate to and call Microsoft Graph directly.

The implementation covers the full workflow Yardstick currently uses:

- OAuth client-credential authentication and token caching
- Graph request serialization, paging, and transient retries
- Win32 app detection, requirement, icon, and application payloads
- Microsoft Win32 Content Prep Tool acquisition and package creation
- `.intunewin` metadata extraction and encrypted content upload
- Win32 app create, read, update, and delete operations
- Group, All Devices, and All Users assignments
- Dependencies and supersedence relationships
- Scope-tag resolution
- Production script, migration script, credential tool, test, CI, and documentation cutover

There are no remaining executable imports or calls to `IntuneWin32App`, `Connect-MSIntuneGraph`, `Invoke-MSGraphOperation`, `Clear-MsalTokenCache`, or the former `Connect-AutoMSIntuneGraph` wrapper.

## Current branch state

- The work is implemented in the current working tree and is not committed.
- The branch was created from `main` before this work began.
- No live Intune tenant was modified during implementation or testing.
- The final full Pester run passed 337 tests with zero failures.
- The CI-equivalent PSScriptAnalyzer run reported zero error-level findings. Warning-level findings remain non-blocking under the existing CI policy.

## New architecture

### `Modules/YardstickGraph.psm1`

This module owns Microsoft Graph authentication and transport. It exports two commands:

- `Connect-YardstickGraph`
- `Invoke-YardstickGraphRequest`

`Connect-YardstickGraph` performs the OAuth 2.0 client-credentials flow directly against:

```text
https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
```

It uses the existing Yardstick credential globals populated by `Initialize-YardstickIntuneCredential`:

- `$Global:TenantID`
- `$Global:ClientID`
- `$Global:ClientSecret`

The function also accepts those values explicitly for isolated use and testing. It requests the `https://graph.microsoft.com/.default` scope, caches the token, and preserves `$Global:Token` and `$Global:AuthenticationHeader` as compatibility state for existing credential and reporting code.

Token behavior:

- A normal connection reuses a token when it has more than 30 minutes remaining.
- `-Force` always obtains a fresh token.
- Requests force a refresh when the cached token has five minutes or less remaining.
- Missing credentials fail with a clear error instead of relying on external-module token helpers.

`Invoke-YardstickGraphRequest` supports `GET`, `POST`, `PATCH`, `PUT`, and `DELETE` against `v1.0` or `beta`. It:

- Accepts either a Graph-relative resource path or a full Graph URL.
- Converts non-string bodies to compressed JSON at depth 20.
- Adds the cached bearer token and JSON headers.
- Flattens OData `value` collections.
- Follows `@odata.nextLink` until all pages have been read.
- Correctly returns zero-, one-, and many-item collections without manufacturing phantom results.
- Retries network failures, HTTP 408, HTTP 429, and HTTP 5xx responses.
- Honors `Retry-After` when available and otherwise uses capped exponential backoff.

### `Modules/YardstickIntune.psm1`

This module owns the Intune Win32 application operations previously supplied by `IntuneWin32App`. It calls `Invoke-YardstickGraphRequest` for Graph operations and contains no external Intune module dependency.

Public commands are grouped below.

Payload builders:

- `New-YardstickWin32AppDetectionRuleFile`
- `New-YardstickWin32AppDetectionRuleMsi`
- `New-YardstickWin32AppDetectionRuleRegistry`
- `New-YardstickWin32AppDetectionRuleScript`
- `New-YardstickWin32AppRequirementRule`
- `New-YardstickWin32AppIcon`

Packaging and app operations:

- `New-YardstickWin32AppPackage`
- `Get-YardstickWin32App`
- `Add-YardstickWin32App`
- `Set-YardstickWin32App`
- `Remove-YardstickWin32App`

Assignments:

- `Get-YardstickWin32AppAssignment`
- `Add-YardstickWin32AppAssignment`
- `Add-YardstickWin32AppAssignmentGroup`
- `Add-YardstickWin32AppAssignmentAllDevices`
- `Add-YardstickWin32AppAssignmentAllUsers`
- `Remove-YardstickWin32AppAssignment`
- `Remove-YardstickWin32AppAssignmentGroup`
- `Remove-YardstickWin32AppAssignmentAllDevices`
- `Remove-YardstickWin32AppAssignmentAllUsers`

Relationships:

- `Get-YardstickWin32AppRelationship`
- `Get-YardstickWin32AppDependency`
- `Get-YardstickWin32AppSupersedence`
- `New-YardstickWin32AppDependency`
- `New-YardstickWin32AppSupersedence`
- `Add-YardstickWin32AppDependency`
- `Add-YardstickWin32AppSupersedence`
- `Remove-YardstickWin32AppDependency`
- `Remove-YardstickWin32AppSupersedence`

The builders preserve Yardstick's existing recipe-facing behavior, including:

- File, MSI, registry, and PowerShell-script detection rules
- Current Windows 10 and Windows 11 minimum release mappings
- `x86`, `x64`, ARM64, combined, and `none` architecture handling
- Optional resource requirements without strict-mode property failures
- PNG and JPEG MIME types for Company Portal icons
- Standard Intune return codes for success, reboot, and retry outcomes
- Maximum installation time and available-uninstall settings
- Scope-tag lookup by exact display name, with the default scope tag ID `0`

## Packaging and encrypted upload flow

`New-YardstickWin32AppPackage` uses Microsoft's official Win32 Content Prep Tool. The tool is not committed to this repository. On first use, Yardstick:

1. Looks for `Tools/IntuneWinAppUtil.exe`.
2. Validates an existing copy using SHA-256.
3. If absent, downloads the executable from Microsoft repository commit `1d6cfcbdf8c28edc596337031f74df951f38f718`.
4. Verifies SHA-256 `C1BA45B5CB939E84AF064BB7FF4B38FB3DFE33C8DC1078FD9B157672EAE671F6`.
5. Moves the validated download into `Tools`.
6. Invokes it with quoted source, setup, and output paths.

The expected output name is based on the setup file name without its extension, matching the Content Prep Tool's behavior.

`Add-YardstickWin32App` performs the complete Intune content transaction:

1. Opens the `.intunewin` ZIP container and reads `detection.xml`.
2. Builds and posts the `win32LobApp` resource.
3. Creates a beta content version.
4. Extracts the encrypted payload named by `ApplicationInfo.FileName` into a unique temporary directory.
5. Creates the `mobileAppContentFile` record with unencrypted and encrypted sizes.
6. Waits for `azureStorageUriRequestSuccess`.
7. Uploads the encrypted payload to the supplied Azure Storage SAS URI in 8 MiB blocks.
8. Renews the SAS URI when a long upload has used it for approximately seven minutes.
9. Commits the block list.
10. Posts the encryption key, MAC key, initialization vector, MAC, digest, digest algorithm, and profile identifier from `detection.xml`.
11. Waits for `commitFileSuccess`.
12. Patches `committedContentVersion` on the app.
13. Returns the completed app from Graph.

Failure handling:

- Temporary extracted content is removed in `finally`.
- If app creation succeeded but a later upload step fails, Yardstick attempts to delete the incomplete Intune app and rethrows the original error.
- Azure Storage block requests have their own retry loop.
- Graph calls use the shared paging, authentication, and retry transport.

## Assignment behavior

Assignments are created as native `mobileAppAssignment` payloads with Win32 assignment settings.

Supported targets:

- Included groups
- Excluded groups
- All Devices
- All licensed users

Supported settings include:

- `required`, `available`, and `uninstall` intents
- Notifications
- Availability and deadline times
- Local-time handling
- Delivery optimization priority
- Assignment filter ID and include/exclude mode

`Get-YardstickWin32AppAssignment` safely projects missing optional target or settings properties. This removes the former external-module edge cases where an empty OData envelope could become a phantom assignment or a single assignment could disappear because of PowerShell collection unrolling.

The generic `Remove-YardstickWin32AppAssignment` removes every assignment from one app selected by ID or display name. Target-specific removal functions retain Yardstick's existing migration behavior.

## Dependency and supersedence behavior

Intune's `updateRelationships` endpoint replaces the submitted relationship set. The native functions therefore preserve the other relationship category when changing one category:

- Adding or removing dependencies preserves forward supersedence entries.
- Adding or removing supersedence preserves forward dependency entries.
- Reverse/parent entries are not resubmitted as if they were owned by the queried app.

Relationship payloads continue to use Intune's beta endpoint. Existing Yardstick support functions still use `targetType` to distinguish forward (`child`) from reverse (`parent`) relationships.

## Production command cutover

The repository-wide rename follows this pattern:

| Previous command | Native command |
|---|---|
| `Connect-AutoMSIntuneGraph` | `Connect-YardstickGraph` |
| `Get-IntuneWin32App` | `Get-YardstickWin32App` |
| `Add-IntuneWin32App` | `Add-YardstickWin32App` |
| `Set-IntuneWin32App` | `Set-YardstickWin32App` |
| `Remove-IntuneWin32App` | `Remove-YardstickWin32App` |
| `New-IntuneWin32AppPackage` | `New-YardstickWin32AppPackage` |
| `New-IntuneWin32App*Rule*` | corresponding `New-YardstickWin32App*Rule*` command |
| `*-IntuneWin32AppAssignment*` | corresponding `*-YardstickWin32AppAssignment*` command |
| `*-IntuneWin32AppDependency` | corresponding `*-YardstickWin32AppDependency` command |
| `*-IntuneWin32AppSupersedence` | corresponding `*-YardstickWin32AppSupersedence` command |

`Yardstick.ps1`, `Migrate-ToSupersedenceModel.ps1`, and `Set-YardstickCredential.ps1` now import the local Graph and Intune modules before their consumers.

`Modules/YardstickSupport.psm1` also imports both sibling modules so it remains usable when imported directly by tests or maintenance scripts.

The old support-module implementations of `Connect-AutoMSIntuneGraph` and `Invoke-YardstickGraphRequest` were removed to prevent command shadowing and duplicated transport behavior.

## Dependency and CI changes

`Test-Prerequisites` now treats only `powershell-yaml` as a required PowerShell Gallery module. Selenium and TUN.CredentialManager remain optional for recipes that need them.

The GitHub Actions test job no longer:

- Downloads the custom IntuneWin32App fork ZIP
- Expands it into the user's module directory
- Copies it into the runner's PowerShell module path

The CI lint list now includes:

- `Modules/YardstickGraph.psm1`
- `Modules/YardstickIntune.psm1`

README and contributor instructions describe the repository-owned implementation and no longer direct operators to install the external module or MSAL.PS.

## Files added

- `Modules/YardstickGraph.psm1`
- `Modules/YardstickIntune.psm1`
- `Tests/YardstickGraph.Tests.ps1`
- `Tests/YardstickIntune.Tests.ps1`
- `HANDOFF.md`

## Files updated

- `Yardstick.ps1`
- `Migrate-ToSupersedenceModel.ps1`
- `Set-YardstickCredential.ps1`
- `Modules/YardstickSupport.psm1`
- `Modules/YardstickCredential.psm1`
- `Test-EmailNotification.ps1`
- `Tests/Set-YardstickSupersedence.Tests.ps1`
- `.github/workflows/test.yml`
- `.github/copilot-instructions.md`
- `README.md`

Most changes outside the two new modules are import updates, command renames, removal of the old Graph wrappers, test mock renames, and documentation cleanup.

## Test coverage added

`Tests/YardstickGraph.Tests.ps1` covers:

- Client-credential token requests
- Compatibility authentication-header publication
- Valid-token reuse
- Missing-credential errors
- OData paging
- Empty collection normalization
- JSON request-body serialization

`Tests/YardstickIntune.Tests.ps1` covers:

- Architecture and Windows release requirement mapping
- MSI detection payloads
- App-body construction with absent optional requirements
- JPEG MIME preservation
- Single-assignment projection with missing settings
- Removing every assignment from an app
- Preserving supersedence when dependencies are replaced

The existing supersedence and migration test suite was updated to mock the new command names and import the local modules explicitly.

## Verification performed

The final implementation was validated with:

```powershell
Invoke-Pester -Path Tests
```

Final result:

```text
Passed: 337
Failed: 0
Skipped: 0
```

Additional checks performed:

- Parsed every `.ps1` and `.psm1` file with the PowerShell AST parser.
- Imported both new modules successfully.
- Verified 31 native Graph/Intune commands are exported in total.
- Compared production command parameters against the imported native command signatures.
- Ran the same PSScriptAnalyzer file set and severity policy used by CI.
- Confirmed zero PSScriptAnalyzer error-level findings.
- Confirmed zero executable legacy references to the external module and its helper commands.
- Ran `git diff --check`; only repository line-ending conversion notices were emitted, with no whitespace errors.

## Live validation still required

Unit and integration-style tests use mocks and do not create tenant resources. Before merging, perform a controlled canary against a non-critical Intune application or test tenant.

Recommended canary sequence:

1. Run `Set-YardstickCredential.ps1` validation to confirm direct OAuth succeeds with the existing app registration.
2. Package a small EXE recipe whose source and output paths contain spaces.
3. Confirm the Content Prep Tool is downloaded and hash-validated, or pre-stage the exact validated executable in `Tools`.
4. Upload the package and verify Intune reaches `publishingState = published`.
5. Confirm display name, version, install/uninstall commands, icon, architecture, minimum OS, return codes, scope tags, and maximum run time.
6. Add and read back one group assignment, one All Devices or All Users assignment, and an assignment filter if available.
7. Create a dependency and supersedence relationship, then verify that changing one does not remove the other.
8. Exercise removal on the canary app and confirm assignments and relationships are cleaned up as expected.
9. Review Graph errors for any tenant-specific permission gaps, especially beta content, relationship, assignment-filter, and scope-tag endpoints.

## Operational notes and known boundaries

- The first package operation requires access to `raw.githubusercontent.com` unless the validated Content Prep Tool already exists in `Tools`.
- A mismatched existing or downloaded Content Prep Tool is rejected rather than executed.
- Content-version uploads and app relationships still rely on Microsoft Graph beta endpoints, matching the capabilities Yardstick needs.
- The current app registration and Credential Manager flow are reused; no new secret storage mechanism was introduced.
- The implementation intentionally retains `$Global:Token` and `$Global:AuthenticationHeader` compatibility state while removing the external authentication library.
- Warning-level PSScriptAnalyzer findings include existing naming/compatibility patterns and do not fail the current CI job.
- The working tree should be reviewed and committed after the live canary succeeds.

## Suggested reviewer focus

Reviewers should pay particular attention to:

- `Add-YardstickWin32App` and its cleanup path
- Azure block upload and SAS renewal behavior for large packages
- Architecture and minimum Windows release values sent to Graph
- Scope-tag lookup behavior in tenants with duplicate or renamed tags
- Relationship preservation around `updateRelationships`
- Assignment settings for exclusions, filters, deadlines, and All Devices/All Users targets
- Required Microsoft Graph application permissions in the target tenant

