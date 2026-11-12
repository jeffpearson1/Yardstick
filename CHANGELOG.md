# Changelog

All notable changes to Yardstick are recorded here. Versions follow
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- **Recipe evidence check.** A recipe can only be merged into `main` with proof that
  it passed a live install test with exactly the contents being merged. Yardstick
  Development Studio commits that proof as `Tests/RecipeEvidence/<id>.json`, and
  the new **Recipe evidence** pull request check
  (`Tests/RecipeHarness/Test-RecipeEvidence.ps1`) requires it for every added,
  changed or renamed recipe, along with a passing `Test-RecipeSchema.ps1`.
  - The `evidence-waived` label turns evidence failures into warnings for recipes
    that can't be live-tested on a build machine. Schema failures still fail.
  - `Test-RecipeSchema.ps1` has a new `-SkipIconCheck` switch for checkouts
    without icons.

  See [Recipe evidence](Docs/RecipeEvidence.md).

- **Recipe repositories and sync.** With the new `syncRecipes: true` preference,
  every run brings its recipes up to date from GitHub before it reads any.
  - **`recipeRepository`** (default `jeffpearson1/Yardstick@main`) is synced into
    `Recipes`.
    - Replaced and removed files are backed up to `Artifacts\RecipeSyncBackups`.
    - A recipe edited here after it was synced is kept, with a warning.
  - **`additionalRecipeRepositories`** links more repositories. Each is mirrored into
    `Local\RecipeRepos` and searched after `Recipes`, so a recipe in `Recipes` wins on
    a duplicate id.
  - **Only changed files are downloaded,** through the GitHub API, so the stored
    token reaches private repositories.
  - **`requireRecipeEvidence: true`** only takes recipes whose
    [evidence](Docs/RecipeEvidence.md) proves they passed a live test with exactly
    those contents.
  - **`Test-YardstickSetup.ps1`** reports when each repository was last synced.

  See [Recipe repositories](Docs/Installation.md#recipe-repositories).

- **Recipe variants: more than one app from a recipe.** An entry in
  `Local\recipeOverrides.yaml` that sets `variantOf: <recipe id>` builds a second app
  from that recipe, such as a lab Zoom next to the standard one. The variant starts
  from the recipe with the recipe's own overrides applied, and its entry is layered
  on top, so it only sets what differs; it needs a `displayName` of its own. `-All`
  runs each variant straight after its recipe, and `-ApplicationId`, recipe groups
  and `Migrate-ToSupersedenceModel.ps1` take its id. `Test-YardstickSetup.ps1` lists
  the variants and warns about any that cannot be built. This replaces the old
  gitignored `Recipes\<recipe>.common.yaml` copies. See
  [Variants](Docs/WritingRecipes.md#variants-more-than-one-app-from-a-recipe).

- **Authenticated GitHub API.** Store a read-only fine-grained GitHub token with
  `.\Set-YardstickCredential.ps1 -GitHub` (kept in Credential Manager as
  `Yardstick:GitHub`) to raise GitHub's limit from 60 to 5,000 API calls an hour.
  Every `api.github.com` call, from `GithubDownloader` and from recipes, now goes
  through the new `Invoke-YardstickGitHubApi`, which:
  - sends the token to `api.github.com` only, and identifies Yardstick to GitHub;
  - carries on anonymously if GitHub rejects the token, and warns when the token
    is within 14 days of expiring.

  Each run logs the calls left, `Test-YardstickSetup.ps1` checks the token, and
  `Test-RecipeSchema` warns about scripts that call the API directly. The 23 recipes
  that did are switched over.

- **Yardstick waits as long as GitHub asks.** When GitHub throttles a call, Yardstick
  makes no more calls to GitHub (the API, `raw.githubusercontent.com` or a release
  download), from any recipe, until the wait GitHub asked for has passed: the
  `Retry-After` time, the reset time once the rate limit is used up, or otherwise a
  minute, doubling each time GitHub throttles again within 15 minutes of the last wait
  ending. The log records GitHub's explanation and request ID.
  - A recipe that needs GitHub during the wait is set aside, the run carries on with
    the others, and the recipe runs again at the end once the wait has passed. A
    recipe is set aside at most three times, and a run waits for GitHub for at most
    75 minutes in all. The email report shows each recipe once.
  - The recipe harness waits in place instead, for up to five minutes at a time.
  - The new `Invoke-YardstickGitHubRaw` reads `raw.githubusercontent.com` under the
    same rules, without the token. The capcut, djiterra, kindle and snagit recipes,
    which read winget-pkgs manifests there, use it, and `Test-RecipeSchema` warns
    about scripts that read the host directly.

- **Development-only instances.** `.\Install-Yardstick.ps1 -DevelopmentOnly` (or
  answering yes to *Is this a development-only computer?*) sets `developmentOnly: true`
  and sets up a computer for writing and testing recipes with the recipe harness. It
  never signs in to Intune or Graph: the installer skips the `AppRegistration`,
  `IntuneDefaults`, `Email` and `ScheduledTask` steps, `Test-YardstickSetup.ps1` skips
  the matching checks, and `Yardstick.ps1` refuses to run. It can still sync the shared
  icons and recipe overrides.
- **Shared folders in the installer.** The `Preferences` step asks whether to sync
  icons and recipe overrides with a shared folder, and for both locations
  (`-SharedIcons`, `-SharedRecipeOverrides`; a folder, or any path without a `.yaml`
  or `.yml` extension, means the `recipeOverrides.yaml` in it, even if the folder
  doesn't exist yet). A mapped drive letter is stored as the share it maps to, since
  the scheduled task doesn't see drive mappings. Relative and `%VARIABLE%` paths are
  resolved against the Yardstick folder, as a run resolves them, and stored as given.
  Passing the locations already set, in any of those forms, doesn't turn
  `syncSharedContent` back on after it was turned off.
- **`-SoftwareDropbox` in the installer**, which also works for development-only
  instances and with `-NonInteractive`, so the recipe harness can test
  `manualDownload` recipes there. `''` clears it. `Invoke-RecipeInstallTest.ps1` now
  resolves `SoftwareDropbox` as a run does, so relative and `%VARIABLE%` paths work.
- **The installer can be driven by another program.** For the Yardstick Development
  Studio's setup wizard, which runs it without a console:
  - `-InstallOptionalModule` installs the named optional modules (Selenium,
    TUN.CredentialManager, Pester) without asking, and `-SetExecutionPolicy` sets
    RemoteSigned for the current user when the policy blocks Yardstick's scripts.
    Both work with `-NonInteractive`, which on its own still answers no.
  - `-ResultPath` writes the outcome of every step, and every status line, to a JSON
    file, even when a step fails.
  - `-Status` reports the state of setup as JSON without changing anything, before
    or after setup has run. It includes `preferences.softwareDropbox`.

  See [Running the installer from another program](Docs/Installation.md#running-the-installer-from-another-program).

### Changed

- **Shared icons and recipe overrides sync in both directions.** Icons and override
  fields added or changed on this machine are now written to the share, as well as
  taken from it, and the recipe harness scripts (`Invoke-RecipeLab.ps1`,
  `Test-RecipeSchema.ps1`, `Invoke-RecipeInstallTest.ps1`) sync too.
  - Icons: the newer copy wins; each icon is checked to be a complete PNG or JPEG
    before it is copied; nothing is deleted.
  - Overrides: merged field by field against the value at the last sync (kept in
    `Local\recipeOverrides.shared.sync.json`). A field changed on both sides is
    reported as a conflict and the local value applies; a local deletion is never
    pushed. Entries are added to the shared file in place, keeping its comments and
    order.
  - Nothing is written unless both files parse, every write is read back, and the
    shared file is changed under a lock, only if no one else changed it meanwhile,
    keeping `recipeOverrides.yaml.bak`. A share that is unreachable, malformed, or
    suddenly empty is left alone and the run carries on with the last synced copy.
  - Recipe ids or fields that differ only in case are now reported as malformed rather
    than one silently replacing the other.
  - **Before upgrading,** check `Local\recipeOverrides.yaml` on every machine that
    syncs. Every value in it that differs from the shared file is now written to the
    share (or reported as a conflict), so an old full copy of the shared file would
    replace newer shared values with out-of-date ones. Delete such a copy, and remove
    any value that should not reach every machine.
- A missing `SharedRecipeOverrides` file in a reachable folder is no longer a warning;
  the first sync creates it. A `SharedRecipeOverrides` folder means the
  `recipeOverrides.yaml` file in it, with a warning from `Test-YardstickSetup.ps1`.
  Setup used to store a folder that didn't exist yet, or wasn't reachable, as the
  file, and every sync then failed. Such settings now work, and the next
  `.\Install-Yardstick.ps1 -NonInteractive` run corrects them.
- **Custom schedule days.** The installer's `ScheduledTask` step offers Every day,
  Weekdays (Monday to Friday) or Custom, which asks which days to run
  (`-TaskSchedule Daily|Weekdays|Custom`, `-TaskDays`). `-TaskSchedule MondayToFriday`
  and `MondayToThursday` from 1.1.0 still work. `Register-YardstickScheduledTask` takes
  `-DaysOfWeek` instead of `-Schedule`.
- **`versionLock` matches whole numbers.** A lock no longer matches a version whose
  number merely starts with the locked one: `3.1` now matches `3.1.4` but not `3.12.0`,
  and `8` no longer matches `80.1`. No committed recipe uses `versionLock`; check any
  lock in `Local\recipeOverrides.yaml` that relied on the old prefix matching.
- `tableaudesktop` is in the `NoGitHub` recipe group; it does not use GitHub.
- **A stuck recipe no longer ends the run.** When a recipe stage makes no progress for
  `recipeStageTimeoutMinutes`, the watchdog now abandons that recipe and moves on to
  the next. It interrupts whatever the stage is waiting on (a download, a sleep, a
  process), the recipe is recorded as failed with "Abandoned after N minutes without
  progress in stage ...", and the email report lists it with the other failures.
  Before, the watchdog ended the whole run, so every recipe after the stuck one was
  skipped. A few calls cannot be interrupted, such as `Start-Process -Wait` or a native
  call that blocks. If the stage still has not moved five minutes after the
  interrupt, the watchdog ends the run as before.
- **Recipes download with `Invoke-YardstickBitsDownload`.** The 13 recipes that called
  `Start-BitsTransfer` themselves (endnote, grammarly, ltspice, marcedit, mobaxterm,
  nvivo, postshot, the three Vernier apps, both Visual C++ redistributables and
  winscp) now use the helper. It gives up on a transfer that moves no data and tells
  the watchdog when a slow download is still moving. The helper takes `-Dynamic`,
  uses `bitsStallTimeoutMinutes` when a recipe calls it, and `Test-RecipeSchema` warns
  about a synchronous `Start-BitsTransfer` in a recipe script.

### Fixed

- **Moved group assignments are removed from the old version again.** Removing a migrated group
  assignment from the previous version failed whenever that app also had an All
  Devices or All Users assignment, leaving the group assigned to both versions.
- **Filtered All Devices and All Users assignments move to the new version.** They were
  left on the old version, which was then kept, because filters once could not be
  copied. The filter now goes with them.
- **Assignments keep their Delivery Optimization priority and restart settings** when
  they move to a new version; they were reset to the defaults.
- The installer's execution policy check no longer passes just because the installer
  was started with `pwsh -ExecutionPolicy Bypass`. It now checks the policy a normal
  window gets, so a CurrentUser or LocalMachine policy that blocks Yardstick's
  scripts is found, and it says when Group Policy sets the policy. Setting
  RemoteSigned from such a session is no longer reported as failed.
- The installer accepts paths in quotes, as Explorer's *Copy as path* produces, for
  every folder and file it asks for, and preferences paths written in quotes resolve
  correctly.
- A failed Intune upload is now retried properly. The retry used to add a second
  content version to the same app, which Intune refuses until the first is committed,
  so it always failed. It now removes the half-made app, waits a minute, and starts
  again with a new one. The retry also covers Intune failing or timing out the final
  commit, which happens when Azure storage is busy.
- FTK Imager: reads the version from Exterro's new download file name
  (`Exterro_FTK_Imager_8.3.1.6.zip`).
- A web download in a recipe that stops partway now fails after
  `webRequestTimeoutSeconds` with nothing received, instead of holding the run until
  the stage watchdog. `-TimeoutSec` only covered the wait for the server to start
  answering, so a vendor host that sent half an installer and then went quiet on an
  open connection blocked the run with no limit. Yardstick now also sets
  `-OperationTimeoutSeconds`, which bounds the gap between reads, so a slow download
  that is still moving is never cut off.

### Removed

- `Get-YardstickIcon.ps1`, the icon picker, which no longer worked. Icon tooling is
  moving to the Yardstick dev studio; until then, add icons to the `Icons` folder by
  hand. `Test-YardstickSetup.ps1` still lists recipes whose icon is missing.

## [1.1.0] - 2026-10-26

### Added

- **Shared icons and recipe overrides.** With `syncSharedContent: true`, each run first
  copies new and changed icons from `SharedIcons` into the Icons folder (it never
  deletes any) and caches the `SharedRecipeOverrides` file in
  `Local\recipeOverrides.shared.yaml`. `Local\recipeOverrides.yaml` is layered on top,
  field by field, so a value set locally wins. If the share can't be reached, or the
  shared file is malformed, the run warns and uses the last copy it synced.
  `Test-YardstickSetup.ps1` reports the sync status.
- **`minRunInterval` recipe field.** Sets the minimum number of whole days between runs
  of a recipe, for vendor sites that throttle repeated requests. A run counts as soon as
  the recipe is about to contact the vendor, even if it then fails. A recipe run too
  soon is skipped like an up-to-date one, and `-Force` ignores the interval. Last-run
  times are kept per machine in `Local\recipeRunHistory.json`.
- PuTTY now runs at most once a week (`minRunInterval: 7`).
- **`Get-YardstickChromeUserAgent`** returns the user agent of the current stable
  Chrome, read once per run from Google's release feed (and estimated from the last
  known release if the feed can't be reached). Recipes for sites that only serve
  browsers should call it rather than paste in a browser string. See
  [Browser user agent](Docs/WritingRecipes.md#browser-user-agent).
- **Weekday schedules.** The installer's `ScheduledTask` step can run Yardstick every
  day, Monday to Friday, or Monday to Thursday (`-TaskSchedule`).

### Changed

- `Get-RedirectedUrl`, `Get-YardstickIcon.ps1` and the Articulate recipes present
  themselves as the current Chrome. Before, they sent fixed, years-old Chrome versions
  (PowerShell's built-in Chrome agent claims Chrome 7), which some sites, Dropbox
  among them, refuse.
- The README logo adapts to dark mode and is smaller.

### Fixed

- A scheduled task created with Microsoft Store PowerShell stopped working when the
  Store updated PowerShell, because it pointed into a folder named for the old version.
  The task now starts PowerShell through its app execution alias, which survives
  updates. `Test-YardstickSetup.ps1` reports a task whose `pwsh.exe` no longer exists,
  and the `ScheduledTask` step offers to replace it. If you use Store PowerShell, run
  `.\Install-Yardstick.ps1 -Step ScheduledTask` once to move an existing task over.

## [1.0.0] - 2026-10-21

The first versioned release. Earlier builds were unversioned (the module manifest said
`0.1.0`). Upgrading from one of them? Read
[Upgrading to 1.0](Docs/Installation.md#upgrading-to-10) first, then run
`.\Test-YardstickSetup.ps1`.

### Highlights

- **An importable PowerShell module.** `Yardstick.psd1` exports `Invoke-Yardstick`,
  `Install-Yardstick` and `Test-YardstickSetup`. `Yardstick.ps1` still runs Yardstick
  and calls `Invoke-Yardstick`.
- **A guided installer.** `Install-Yardstick.ps1` installs the dependencies, creates the
  Entra app registration and its Graph permissions, writes `Local\preferences.yaml`,
  sets up email and registers the scheduled task. `Test-YardstickSetup.ps1` checks an
  existing installation.
- **Native Microsoft Graph integration.** Packaging, upload, assignments, dependencies
  and supersedence go straight to Graph, so the IntuneWin32App module is no longer
  required. Packaging and blob uploads have timeouts.
- **Supersedence, auto-update and retention.** Each new version supersedes the older
  ones, turns on auto-update for available assignments, moves required assignments
  forward, and prunes versions beyond `numVersionsToKeep`. A pinned DETECT anchor
  version lets supersedence reach devices still running a much older release.
- **Maintenance sweeps.** `-Repair` reconciles naming, assignments, retention and
  supersedence for apps that have no new version.
- **Email reports** through Outlook or SMTP, with separate admin and general reports.
  The installer can stop Outlook's security prompts, and `Send-YardstickLogReport.ps1`
  rebuilds and sends a report from `YLog.log`.
- **250 recipes**, including the JetBrains IDEs.
- **Recipe features:**
  - inheritance with `base:`;
  - multiple detection rules;
  - a software dropbox for licensed media (`manualDownload`);
  - generated PowerShell install scripts that run 64-bit;
  - site values kept out of shared recipes in `Local\recipeOverrides.yaml`.
- **A local recipe test harness** under `Tests\RecipeHarness`. It checks the schema,
  downloads and packages the app, then installs, detects, uninstalls and checks for
  leftovers, without contacting Intune.
- **Promotion and housekeeping scripts:** `Publish-Yardstick.ps1` promotes a development
  folder to production, and `Sync-RecipeGroup.ps1` keeps a recipe group in step with
  its folder.

### Changed

- Recipes no longer carry scope tags or organization-specific values. Recipes that need
  your values ship with `<YOUR_...>` placeholders and fail at the Configuration stage
  until you set them in `Local\recipeOverrides.yaml`.
- `dependentApplicationBlacklist` is now `dependentApplicationBlocklist`. The old name
  still works, with a warning.
- The working folders moved to `Artifacts\BuildSpace`, `Artifacts\Packages` and
  `Artifacts\Temp`.
- `Publish-Yardstick.ps1` needs `-ProductionPath`, and leaves production's own
  `RecipeGroups.yaml` and `Yardstick.Project.psd1` alone.
- `postRunScript` runs for every recipe, including failed ones.
- Running `Yardstick.ps1` with no `-ApplicationId`, `-Group` or `-All` stops with an
  error instead of doing nothing.

### Removed

- `availableGroups` and `requiredGroups` have no effect, and a run warns about recipes
  that set them. Use `defaultDeploymentGroups` for available assignments; add required
  assignments in Intune.

[Unreleased]: https://github.com/jeffpearson1/Yardstick/compare/v1.1.0...HEAD
[1.1.0]: https://github.com/jeffpearson1/Yardstick/releases/tag/v1.1.0
[1.0.0]: https://github.com/jeffpearson1/Yardstick/releases/tag/v1.0.0
