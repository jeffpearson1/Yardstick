# Installing Yardstick

The quick version is in the [README](../README.md#getting-started): download Yardstick,
run `.\Install-Yardstick.ps1`, then run one recipe. This page explains what the
installer does, how to do each part by hand, and how to keep an installation healthy.

- [Requirements](#requirements)
- [What the installer does](#what-the-installer-does)
- [Microsoft Graph permissions](#microsoft-graph-permissions)
- [Creating the app registration yourself](#creating-the-app-registration-yourself)
- [Preferences](#preferences)
- [Email](#email)
- [Running on a schedule](#running-on-a-schedule)
- [Checking an installation](#checking-an-installation)
- [Upgrading](#upgrading)
- [Troubleshooting](#troubleshooting)


## Requirements

- Windows 10 or 11, or Windows Server 2019 or later.
- [PowerShell 7](https://aka.ms/powershell) (`pwsh`). Windows PowerShell 5.1 is not
  supported; the installer tells you how to get PowerShell 7 if you start it from 5.1.
- A normal PowerShell window, signed in as the account that will run Yardstick.
  Yardstick does not need administrator rights, and the installer refuses to start
  from a window opened with **Run as administrator** (see
  [Administrator rights](#administrator-rights)).
- The `powershell-yaml` module. The installer installs it for the current user.
- Enough free disk space in the Yardstick folder to download and package installers
  (a few GB covers most recipe sets).
- An Intune tenant, and for the first setup an account that can create app
  registrations and grant admin consent - Global Administrator or Privileged Role
  Administrator. If you don't have one, an administrator can
  [create the app registration](#creating-the-app-registration-yourself) for you.

Optional:

- `Selenium` and `TUN.CredentialManager`, plus Firefox, for the few recipes that drive a
  browser or sign in to a vendor portal. The recipes need Selenium 4.0.0-preview3 or later,
  which is only published as a prerelease; the installer installs it with prerelease
  versions allowed (`Install-PSResource Selenium -Prerelease`).
- `Pester` 5 or later, to run the unit tests in `Tests\Unit`. Windows includes Pester
  3.4, which is too old. The installer offers to install a current version alongside
  it; by hand, run `Install-PSResource Pester` or
  `Install-Module Pester -Scope CurrentUser -Force -SkipPublisherCheck`.
- `Microsoft.Graph.Authentication`. If it is installed, the installer signs you in
  through the browser; otherwise it uses a device code.


## What the installer does

```powershell
.\Install-Yardstick.ps1
```

The installer runs these steps in order. Each one checks what is already in place, so
it is safe to run again; it only asks about things that are missing or that you want
to change.

| Step | What it does |
|---|---|
| `Prerequisites` | Installs `powershell-yaml`, offers the optional modules, unblocks files from a zip download, checks that you can change every file in the Yardstick folder, checks the execution policy and downloads the Microsoft Win32 Content Prep Tool (`IntuneWinAppUtil.exe`). |
| `Preferences` | Creates `Local\preferences.yaml` from `Config\Preferences.example.yaml`, creates the working folders, and asks for the app owner, the minimum Windows version and the optional Backup and SoftwareDropbox folders. |
| `AppRegistration` | Creates an Entra app registration called Yardstick, grants it the [Graph permissions](#microsoft-graph-permissions), creates a client secret and stores it in Windows Credential Manager. Or stores the details of an app you already have. |
| `IntuneDefaults` | Picks the scope tags for new apps and a test group that receives every new version as Available. |
| `Email` | Optional email report after each run, through Outlook or SMTP. |
| `ScheduledTask` | Optional daily scheduled task. |
| `Verify` | Runs `Test-YardstickSetup.ps1`. |

Run a single step with `-Step`, for example to renew the client secret:

```powershell
.\Install-Yardstick.ps1 -Step AppRegistration
.\Install-Yardstick.ps1 -Step Email, ScheduledTask
```

`Get-Help .\Install-Yardstick.ps1 -Full` lists every parameter.

### Administrator rights

Run the installer as yourself, from a normal PowerShell window. Yardstick runs
without administrator rights, and a file created from an elevated window belongs to
the Administrators group: you can read it afterwards, but not change it, so later runs
fail to update their logs, working folders or preferences.

A few changes do need administrator rights:

- giving you back control of files in the Yardstick folder that an elevated session
  created (see [Troubleshooting](#troubleshooting))
- turning off Outlook's programmatic access warnings (`Email` step)
- creating the scheduled task, on computers that only let administrators create a task
  that runs while nobody is signed in (`ScheduledTask` step)

The installer collects these and makes them together at the end, before `Verify`, so
Windows asks for administrator rights once. A standard user enters an administrator's
name and password at that prompt. If you decline, the summary lists what was not done;
run the step again later.

There is one exception: if the installer cannot write to the Yardstick folder itself,
`Local` or `Tools`, the following steps would fail, so it asks straight away in the
`Prerequisites` step. A run that also needs one of the other changes then asks twice.

### Scripted installs

`-NonInteractive` never prompts. Values come from parameters, the existing preferences
file or the built-in defaults, and a step that is missing information is skipped
rather than failing:

```powershell
$secret = Read-Host -AsSecureString 'Client secret'
.\Install-Yardstick.ps1 -NonInteractive -UseExistingApp `
    -TenantId $tenantId -ClientId $clientId -ClientSecret $secret -ClientSecretExpiresOn 2027-10-01 `
    -Owner 'IT Services' -MinOSVersion W10_22H2 -TestGroup 'Intune App Testers' `
    -EmailMethod smtp -EmailRecipient it@example.com -EmailFrom yardstick@example.com -SmtpServer smtp.example.com -SmtpPort 587 -SmtpUseSsl
```

Creating a new app registration always needs an interactive sign-in, so scripted
installs use `-UseExistingApp`.


## Microsoft Graph permissions

Yardstick signs in as its own app (client credentials), so the app registration needs
these Microsoft Graph **application** permissions, with admin consent:

| Permission | Needed | Why |
|---|---|---|
| `DeviceManagementApps.ReadWrite.All` | Yes | Create, update, assign and delete Win32 apps in Intune. |
| `DeviceManagementRBAC.Read.All` | Yes | Resolve Intune scope tag names to IDs. |
| `Group.Read.All` | No | Lets the installer look up the test deployment group by name. Without it you enter the group's object ID; Yardstick itself only uses group IDs. |
| `Application.Read.All` | No | Discover the client secret's expiry date when none was recorded at setup. |

The installer requests the two required permissions. Add
`-IncludeOptionalPermissions` to request all four. A `.ReadWrite.All` permission the
app already has counts for the matching `.Read.All` one.

Assignment filters need no extra permission: when Yardstick replaces an app version it
copies the existing filters by ID, which `DeviceManagementApps.ReadWrite.All` covers.
An app registration from an earlier version of the installer may also hold
`DeviceManagementConfiguration.Read.All`; Yardstick no longer uses it, so you can remove it.

`Test-YardstickSetup.ps1` reads the permissions that were actually granted from the
access token, and names any that are missing.


## Creating the app registration yourself

Use this when your account cannot create app registrations, or your organisation
creates them through a separate process.

1. In the [Entra admin center](https://entra.microsoft.com), open
   **Identity > Applications > App registrations** and choose **New registration**.
   Name it (for example `Yardstick`), keep **Accounts in this organizational directory
   only**, leave the redirect URI empty and register it.
2. Note the **Application (client) ID** and **Directory (tenant) ID** on the Overview page.
3. Under **API permissions**, choose **Add a permission > Microsoft Graph > Application
   permissions** and add the [permissions above](#microsoft-graph-permissions). Then
   choose **Grant admin consent for _your organisation_**.
4. Under **Certificates & secrets**, create a **New client secret**. Copy the **Value**
   (not the Secret ID) straight away - it is only shown once - and note the expiry date.
5. On the machine that runs Yardstick, signed in as the account that will run it:

   ```powershell
   .\Install-Yardstick.ps1 -Step AppRegistration -UseExistingApp
   ```

   Enter the tenant ID, client ID, secret and expiry date when asked. The installer
   signs in to confirm they work and reports any missing permission.

`Set-YardstickCredential.ps1` does the same as step 5, and can also show, refresh or
remove the stored credential:

```powershell
.\Set-YardstickCredential.ps1 -Show               # tenant, client and expiry (secret masked)
.\Set-YardstickCredential.ps1 -RefreshExpiration  # re-read the expiry date from Graph
.\Set-YardstickCredential.ps1 -Remove             # delete the stored credential
```

### Where the credentials are kept

The tenant ID, client ID and secret are stored in Windows Credential Manager,
encrypted for the Windows account that stored them. They are never written to
`preferences.yaml`. This has two consequences:

- **Store them as the account that runs Yardstick.** If a scheduled task runs as a
  service account, sign in as that account and run
  `.\Install-Yardstick.ps1 -Step AppRegistration` there too.
- Each machine needs its own copy.

If an older `preferences.yaml` still contains `TenantID`, `ClientId` and
`ClientSecret`, the next Yardstick run moves them into Credential Manager and reminds
you to delete them from the file.

### Client secret expiry

Every run checks how long the client secret has left. Inside
`credentialExpirationWarningDays` (default 30) it writes a warning and emails
`adminEmailRecipient` (or `emailRecipient`), at most once every
`credentialExpirationEmailIntervalHours` (default 24).

Yardstick knows the expiry date if the installer created the secret, or if you entered
it. Otherwise it reads it from Graph, which needs `Application.Read.All`.

To renew the secret, run `.\Install-Yardstick.ps1 -Step AppRegistration` and choose to
create the app registration. It finds the existing app by name, adds a new secret and
stores it; delete the old secret in Entra afterwards. If you manage the app yourself,
create a new secret in Entra and run the step with `-UseExistingApp`.


## Preferences

Settings live in `Local\preferences.yaml`, which the installer creates from
`Config\Preferences.example.yaml`. The `Local` folder is excluded from Git, so a
`git pull` never overwrites your settings.

Every setting has a built-in default, so the file only needs the values you want to
change. The example file lists them all with explanations. In particular:

- **Folders** (`Temp`, `BuildSpace`, `Published`, `Recipes`, `Icons`, `Tools`,
  `Scripts`, ...) default to folders inside the Yardstick directory. Relative paths are
  resolved against it; absolute and UNC paths are used as they are. Working folders
  are created when Yardstick starts.
- **`defaultDeploymentGroups`** and **`defaultGroupsSkipAutoUpdates`** take Entra group
  object IDs, not names.
- **`SoftwareDropbox`** and **`SoftwareArchive`** are for recipes whose installers must
  be downloaded by hand. See [SoftwareDropbox.md](SoftwareDropbox.md).
- **`Backup`** keeps a copy of every `.intunewin` file Yardstick uploads (the newest
  `backupVersionsToKeep` per recipe), so a bad release can be traced to the exact
  package. Leave it blank to turn this off.

Yardstick checks the file when it starts and reports invalid values - for example an
unknown `emailDeliveryMethod`, a group name where an ID is expected, or a leftover
`<REPLACE_...>` placeholder - before it does anything else.

Recipes that sign in to a vendor portal read a Windows credential from Credential
Manager. Create the credential each recipe names, as the account that runs Yardstick.


## Email

Yardstick can send two reports after each run: a short list of updated software to
`emailRecipient`, and a full report with errors and warnings to
`adminEmailRecipient`. The `Email` step sets this up, and
`.\Test-EmailNotification.ps1` sends a test message.

| Method | Use it when |
|---|---|
| `smtp` | Yardstick runs on a schedule or on a server. Needs an SMTP server; a sign-in, if the server needs one, is stored in Credential Manager. |
| `outlook` | Yardstick runs interactively on a machine with Outlook. Outlook needs a signed-in user, so it usually fails from a scheduled task. |

See [EmailNotifications.md](EmailNotifications.md) for every setting.


## Running on a schedule

The `ScheduledTask` step creates the task `\Yardstick\Yardstick`, which runs daily at a
time you choose:

```
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File <Yardstick folder>\Yardstick.ps1 -All -NoInteractive
```

It runs `-Group <name>` instead of `-All` if you pick a recipe group from
`RecipeGroups.yaml`. `-NoInteractive` makes sure the run never waits for input.

Things to know:

- The task runs whether or not anyone is signed in, so Task Scheduler asks for the
  account's password.
- It must run as the account that stored the Intune credentials (see
  [Where the credentials are kept](#where-the-credentials-are-kept)).
- Some computers only let administrators create a task that runs while nobody is
  signed in. The installer then creates it with administrator rights at the end of
  setup (see [Administrator rights](#administrator-rights)). The task still runs as
  your account, without administrator rights.
- If Yardstick is on a network share, the task account needs access to it, and must use
  the UNC path; mapped drive letters are not available to scheduled tasks.
- Use SMTP rather than Outlook for email.

To create the task by hand, use the same command line, set **Start in** to the
Yardstick folder, and choose **Run whether user is logged on or not**.


## Checking an installation

```powershell
.\Test-YardstickSetup.ps1
```

The health check reports OK, Warn or Fail for:

- PowerShell version and modules
- `Local\preferences.yaml`
- that the working folders exist and are writable, and that the account can change
  every file in the Yardstick folder
- the stored Intune credential, the secret's expiry, sign-in, the granted Graph
  permissions and a test call to Intune
- `IntuneWinAppUtil.exe`
- the email settings
- the scheduled task and its last result
- recipes whose icon is missing from the Icons folder

It exits with code 1 when anything fails, so it can be used as a pre-flight check.
`-Offline` skips the checks that sign in to Microsoft Graph, and `-PassThru` returns
the results as objects.

Icons are not included for licensing reasons. `Get-YardstickIcon.ps1` helps find
and save the ones your recipes need.


## Upgrading

```powershell
git pull
.\Install-Yardstick.ps1
```

Your settings in `Local` are not touched by `git pull`. Re-running the installer adds
anything a new version needs and skips what is already done; the last step runs the
health check. If you downloaded a zip instead, extract the new version over the old
folder (keeping `Local`, `Icons` and `Artifacts`) and run the installer.

**Upgrading to the module layout.** From this version the folder is a PowerShell module:
`Yardstick.psm1` and `Private\Invoke-YardstickRun.ps1` hold the code, and `Yardstick.ps1`
becomes a short wrapper with the same parameters. Scheduled tasks and anything else that
calls `.\Yardstick.ps1` keep working unchanged, but the three files must be updated
together: a `Yardstick.ps1` copied on its own will not run. `git pull`, extracting the
whole zip, and `Publish-Yardstick.ps1 -Runtime` all do that.

The scripts always load the current code. If you use the module commands
(`Invoke-Yardstick`, `Test-YardstickSetup`, ...) in a PowerShell window that was open
before an update, run `Import-Module .\Yardstick.psd1 -Force` first, or open a new
window.


## Troubleshooting

**"No Yardstick preferences file was found"** - `Local\preferences.yaml` is missing.
Run the installer, or copy `Config\Preferences.example.yaml` there.

**"The PowerShell module 'powershell-yaml' is not installed"** - run
`Install-Module powershell-yaml -Scope CurrentUser`, or the installer.

**"No Intune credentials found in Windows Credential Manager"** - the credential
belongs to a different Windows account, or was never stored. Run `.\Install-Yardstick.ps1 -Step AppRegistration` as the
account that runs Yardstick.

**The device-code sign-in is blocked** - some Conditional Access policies block the
device-code flow. Install `Microsoft.Graph.Authentication`
(`Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`) and run the step
again to sign in through the browser instead, or
[create the app registration yourself](#creating-the-app-registration-yourself).

**"Could not grant admin consent"** - your account can create the app but not consent
to its permissions. The installer prints an admin-consent link; a Global Administrator
or Privileged Role Administrator opens it and accepts. Then run
`.\Test-YardstickSetup.ps1`.

**A missing permission straight after setup** - changes in Entra can take a few
minutes to reach the token. Wait, then run `.\Test-YardstickSetup.ps1` again.

**Scripts are blocked or "not digitally signed"** - files from a zip download are
marked as coming from the internet. The installer's `Prerequisites` step unblocks them,
or run `Get-ChildItem -Recurse | Unblock-File` in the Yardstick folder.

**"Run the installer from a normal PowerShell window"** - the installer was started
with **Run as administrator**. Close that window and run it from a normal one; it asks
for administrator rights itself when a step needs them.

**"Access to the path ... is denied" for a log or working file** - setup or Yardstick
was once run as administrator, and the files it created belong to the Administrators
group. `.\Test-YardstickSetup.ps1` lists them under Permissions. Run
`.\Install-Yardstick.ps1 -Step Prerequisites` and accept the administrator prompt: it
gives your account full control of the Yardstick folder and everything in it.
