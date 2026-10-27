# Changelog

All notable changes to Yardstick are recorded here. Versions follow
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- **Development-only instances.** `.\Install-Yardstick.ps1 -DevelopmentOnly` (or
  answering yes to *Is this a development-only computer?*) sets `developmentOnly: true`
  and sets up a computer for writing and testing recipes with the recipe harness. It
  never signs in to Intune or Graph: the installer skips the `AppRegistration`,
  `IntuneDefaults`, `Email` and `ScheduledTask` steps, `Test-YardstickSetup.ps1` skips
  the matching checks, and `Yardstick.ps1` refuses to run. It can still sync the shared
  icons and recipe overrides.
- **Shared folders in the installer.** The `Preferences` step asks whether to sync
  icons and recipe overrides with a shared folder, and for both locations
  (`-SharedIcons`, `-SharedRecipeOverrides`; a folder means the
  `recipeOverrides.yaml` in it).

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
  the first sync creates it. A path that is a folder is an error.
- **Custom schedule days.** The installer's `ScheduledTask` step offers Every day,
  Weekdays (Monday to Friday) or Custom, which asks which days to run
  (`-TaskSchedule Daily|Weekdays|Custom`, `-TaskDays`). `-TaskSchedule MondayToFriday`
  and `MondayToThursday` from 1.1.0 still work. `Register-YardstickScheduledTask` takes
  `-DaysOfWeek` instead of `-Schedule`.

### Fixed

- The installer accepts paths in quotes, as Explorer's *Copy as path* produces, for
  every folder and file it asks for, and preferences paths written in quotes resolve
  correctly.

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
