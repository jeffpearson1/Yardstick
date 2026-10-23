# Changelog

All notable changes to Yardstick are recorded here. Versions follow
[Semantic Versioning](https://semver.org/).

## [Unreleased]

### Added

- `minRunInterval` recipe field: the minimum number of days between runs of a recipe,
  for vendor sites that throttle repeated requests. Last-run times are kept in
  `Local\recipeRunHistory.json`, and `-Force` ignores the interval.

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

[1.0.0]: https://github.com/jeffpearson1/Yardstick/releases/tag/v1.0.0
