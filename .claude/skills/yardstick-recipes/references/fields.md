# Recipe Fields

Full reference: `Docs\WritingRecipes.md`. This file covers only the choices that
are easy to get wrong.

## Required by the schema validator

`id` (must match filename exactly, lowercase), `displayName`, `detectionType`,
`iconFile` (must exist in `Icons\`), `description`, `publisher`, one of
`installScript` / `powerShellInstallScript`, one of `uninstallScript` /
`powerShellUninstallScript`.

Each may be satisfied either as a YAML key *or* by assigning the matching variable
(`$registryDetectionKey = ...`) inside a script block — the validator checks both.

## Required at runtime but not schema-checked

| Variable | Where it may be set | Notes |
|---|---|---|
| `$version` | YAML `version`, or `preDownloadScript` **only** | Validated immediately after `preDownloadScript`. Setting it in `downloadScript` is too late and is the single most common recipe bug. |
| `$fileName` | YAML, `preDownloadScript`, or `downloadScript` | Must name the *final* installer. `postDownloadScript` may reassign it (see extraction patterns). |
| `$url` | YAML, `preDownloadScript`, or `downloadScript` | Not needed if `downloadScript` performs the download itself, or if `manualDownload: true`. |

## Script blocks

Order: `preDownloadScript` → (version validation) → `downloadScript` →
`postDownloadScript` → … packaging/upload … → `postRunScript`.

All run with `-NoNewScope`: they share one scope, see each other's variables, and
writing `$version = ...` writes Yardstick's own `$Script:Version`. That power cuts
both ways — assigning `$temp`, `$buildSpace`, `$scripts`, `$published`, `$recipes`,
`$icons`, `$tools`, `$secrets`, `$prefs`, or `$applications` clobbers the runner's
folder variables. The schema validator warns on this; rename your locals.

Working directory for `downloadScript` and `postDownloadScript` is
`$BuildSpace\$id\$version`. Write every intermediate file there so `Invoke-Cleanup`
can sweep it.

Available in all script blocks: `$BuildSpace`, `$Temp`, `$Scripts`, `$Published`,
`$Backup`, `$Recipes`, `$Icons`, `$Tools`, `$Secrets`. Plus, for manual-download
recipes, `$dropboxPath` and `$dropboxFiles`.

`postRunScript` runs outside the try/catch — errors there do not fail the recipe.

## Choosing where the version comes from

Escalate only as far as you must:

1. Static `url` + `urlRedirects: true` — the vendor 302s to a versioned artifact,
   and `preDownloadScript` parses the resolved `$url`. (firefox, amazoncorretto21)
2. Static `url` + `preDownloadScript` scraping a vendor version API or landing
   page. (googlechrome, 7zip, dragonframe)
3. `preDownloadScript` using a shared downloader module — `GithubDownloader`,
   `JetbrainsDownloader`. Prefer these over hand-rolled scraping.
4. `postDownloadScript` reading the version off the downloaded artifact, when the
   vendor publishes no version anywhere. Note this still requires *some* `$version`
   set in `preDownloadScript` — the usual trick is to stage the file to `$Temp`
   during `preDownloadScript` and read `VersionInfo.ProductVersion` there. (marcedit)
5. `downloadScript` — full custom flow: authentication, cookies, hash-validated
   multi-file payloads, Selenium. (keepass, Adobe recipes)

`urlRedirects` is the canonical casing — `Yardstick.ps1` reads
`$Parameters.urlRedirects`, and the schema validator warns on any other casing.
PowerShell hashtable lookup is case-insensitive, so a miscased key still resolves
at runtime, but it will draw a warning on every schema run.

## Detection type decision

| Situation | Use |
|---|---|
| MSI with a stable ProductCode | `detectionType: msi` — ProductCode is extracted automatically by `Get-MsiProductCode`. Never hardcode a GUID. |
| App drops a known versioned EXE at a known path | `detectionType: file` + `fileDetectionPath`, `fileDetectionName`, `fileDetectionMethod: version`, `fileDetectionOperator` |
| Stable ARP key or registry value | `detectionType: registry` + `registryDetectionMethod`, `registryDetectionKey` |
| Dynamic product codes, multiple registrations, version format needing normalization, or a portable/hand-rolled install | `detectionType: script` + `detectionScript` |

`fileDetectionVersion` defaults to `$version`. Set it only when the installed
file's version format differs — commonly 4-part padding:
`$fileDetectionVersion = [VersionPro]::new($version).ToString(4)`.

A `detectionScript` must `Write-Output 'Detected'; exit 0` on match and
`Write-Output 'Not Detected'; exit 1` otherwise. Prefer `-ge` comparisons against
the packaged `<version>` so supersedence can still see older installs. `sharex.yaml`,
`marcedit.yaml`, and `sysinternals.yaml` are good models.

Note that `msi` / `file:version` / `registry:version` detection types participate
in the `{DETECT}` anchor mechanism that pulls stale endpoints forward; script and
existence-only detection do not.

## Placeholders

`<filename>`, `<version>`, `<productcode>` are expanded by
`Update-ScriptPlaceholders` in `installScript`, `powerShellInstallScript`,
`uninstallScript`, `powerShellUninstallScript`, `detectionScript`, and
`registryDetectionKey`. Use them instead of interpolating manually — the recipe
**fails** if any placeholder is still unreplaced after the three expansion passes.

`<productcode>` is populated only for MSI payloads.

## Install script style

- Simple MSI: `installScript: msiexec /i <filename> /qn /norestart` /
  `uninstallScript: msiexec /x <productcode> /qn /norestart`.
- Simple EXE: `installScript: <filename> /S` (or the vendor's documented switch).
- Anything needing ordering, prerequisites, conditional logic, or exit-code
  mapping: use `powerShellInstallScript` / `powerShellUninstallScript`. Yardstick
  generates `install.ps1` / `uninstall.ps1` and wires up the invocation. Inside
  those scripts, `$PSScriptRoot` is the package root — that is how you reach
  prerequisites you packaged alongside the installer.
- Accept the benign MSI exit codes explicitly: `0, 1641, 3010` for install,
  plus `1605, 1614` for uninstall-when-absent, `1638` for "already installed".

## Behaviour and lifecycle fields

`installExperience` (`system`/`user`), `restartBehavior` (`suppress` is the norm),
`minOSVersion` (`W10_20H2` is the common floor), `architecture`,
`maximumInstallationTimeInMinutes` (raise it for long installers),
`numVersionsToKeep`, `versionLock` (dotted pattern with `x` wildcards,
e.g. `3.11.x`), `scopeTags`, `dependentApplicationBlacklist`.

`fileType` (`msi`/`exe`/`zip`) and `softwareName` pass schema validation and appear
throughout the existing recipes, but nothing in `Yardstick.ps1` or the modules
reads them — they annotate the payload for human readers and nothing more. Match
the surrounding recipes for consistency; don't expect either to change behavior.

Supersedence and auto-update (`supersedence`, `uninstallPreviousVersion`,
`autoUpdateOnAssignment`, `groupSkipAutoUpdates`) all default from
`preferences.yaml`; set them per-recipe only for a deliberate reason. Note
`groupSkipAutoUpdates` **unions** with the preference rather than overriding it.

## Inheritance

`base: <other-recipe-id>` shallow-merges the child over the base. Single level
only — a base recipe may not itself have a `base`. Typical use is an arm64 or
site-licensed variant overriding `id`, `displayName`, and `preDownloadScript`.

## Pitfalls

1. `$version` set in `downloadScript` — validated too early; always `preDownloadScript`.
2. `$fileName` not updated after `postDownloadScript` changes the payload.
3. Detection version format mismatch (3-part vs 4-part) — pad with `[VersionPro]`.
4. Missing icon in `Icons\`.
5. Hardcoded paths, GUIDs, or versions where a placeholder belongs.
6. Assigning a reserved runner variable inside a pre/download/post script.
7. Hard-coded versioned download URL that will 404 next release.
8. `detectionType: file` pointed at a file the installer only sometimes writes.
