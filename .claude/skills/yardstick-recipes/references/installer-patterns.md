# Installer Patterns

Each pattern names a recipe you can read in full. Paths are relative to
`G:\Intune\YardstickDev`.

## Plain MSI

The easy case. Static or scraped `url`, MSI detection, ProductCode placeholder.

```yaml
installScript: msiexec /i <filename> /qn /norestart
uninstallScript: msiexec /x <productcode> /qn /norestart
fileType: msi
detectionType: msi
```

Reference: `Recipes\googlechrome.yaml`, `Recipes\7zip.yaml`, `Recipes\zoomrooms.yaml`.
Do not add extra detection metadata unless the vendor ships multiple MSI payloads.

## Vendor MSI needing extra properties

Chained or transformed MSIs often need properties (`INSTALLDIR`, `ALLUSERS=1`,
`NOAUTOUPDATE=1`). Keep the placeholders; add the properties to the argument list.
Reference: `Recipes\qgis.yaml`.

## InnoSetup / NSIS EXE

InnoSetup: `<filename> /VERYSILENT /SUPPRESSMSGBOXES /NORESTART`, uninstall via the
`unins000.exe` in the install directory or the ARP `UninstallString` /
`QuietUninstallString`. NSIS: `/S`. Always confirm against vendor docs — some
installers ignore or actively refuse the conventional switch.

Reference: `Recipes\Development\sharex.yaml` (InnoSetup, with an uninstall script
that falls back from install-dir discovery to ARP parsing).

## Bootstrapper extraction (Advanced Installer and friends)

An EXE bundle whose bootstrapper deletes its temp MSI before msiexec opens it
cannot be deployed as-is. Extract at *build* time and ship the loose payload:

```yaml
postDownloadScript: |
  $wrapper = Join-Path $PWD $fileName
  $process = Start-Process -FilePath $wrapper -ArgumentList '/extract' -WorkingDirectory $PWD.Path -Wait -PassThru
  if ($process.ExitCode -ne 0) { throw "extraction exited with $($process.ExitCode)." }
  # verify every expected artifact exists, then:
  $fileName = 'setup.msi'
  Remove-Item -LiteralPath $wrapper -Force
```

Preserve the exact relative layout the MSI expects (external CABs, prerequisite
subfolders). Advanced Installer MSIs extracted this way usually need
`SETUPEXEDIR="<package root>"` passed to msiexec, otherwise custom actions fail.

Reference: `Recipes\Development\dragonframe.yaml` (validated live, zero residue),
`Recipes\Development\marcedit.yaml`.

Note `Start-Process -Wait` on a GUI-subsystem extractor also tracks its extraction
child; invoking the EXE directly often returns before extraction finishes.

## Packaging prerequisites

When the vendor bootstrapper would download prerequisites at install time, make the
package self-contained instead: fetch them in `preDownloadScript`/`downloadScript`
(hash-verified against vendor metadata where available), then install them in
order from `powerShellInstallScript` via `$PSScriptRoot`.

```yaml
powerShellInstallScript: |
  $redist = Join-Path $PSScriptRoot 'Visual C++ Redistributable for Visual Studio 2019\vc_redist.x64.exe'
  $process = Start-Process -FilePath $redist -ArgumentList '/install','/quiet','/norestart' -Wait -PassThru
  if ($process.ExitCode -notin 0, 1638, 1641, 3010) { throw "prerequisite exited with $($process.ExitCode)." }
```

Reference: `Recipes\Development\marcedit.yaml` packages x64 + x86 .NET Desktop
Runtime (resolved from Microsoft's `releases.json` and SHA512-verified) plus
WebView2 before its MSI.

Alternatively, express the prerequisite as a separate recipe and wire it as an
Intune dependency — that is the right call when the prerequisite is independently
useful. Reference: `azurevirtualdesktopagent.yaml` + `azurevirtualdesktopagentbootloader.yaml`.

## ZIP / portable, no installer

There is nothing to run, so you own the whole install: expand to Program Files,
add to machine PATH if appropriate, and **publish your own ARP entry** so the app
is detectable and looks normal to users. Mirror it exactly in the uninstall script.

Reference: `Recipes\Development\sysinternals.yaml` (also derives its version from
the ZIP's `Last-Modified` header, since the vendor publishes no version at all),
`Recipes\Development\ffmpeg.yaml`.

Detection for these is always `detectionType: script` against the ARP key you
created *and* a marker file on disk.

## GitHub releases

Use the shared downloader rather than scraping HTML:

```yaml
url: https://github.com/<owner>/<repo>
preDownloadScript: |
  using module .\Modules\Custom\GithubDownloader.psm1
  $Script:downloader = [GithubDownloader]::new($url, "<asset-name-regex>", "<version-regex>")
  $downloader.Update()
  $version  = $downloader.LatestVersion
  $url      = $downloader.URL
  $fileName = $downloader.filename
```

Set `$downloader.Prerelease = $true` only if the vendor ships through prereleases.
Reference: `Recipes\powershell7.yaml`, `Recipes\dbeaver.yaml`, `Recipes\bluej.yaml`,
`Recipes\djv.yaml`. Calling the GitHub REST API directly is also acceptable when
you need asset-selection logic the module does not express — see `sharex.yaml`.

## JetBrains IDEs

Recipes live in `Recipes\Jetbrains\`. In `preDownloadScript`:
`Get-JetbrainsAppDownloadLink -ProductName "<Name>" -Channel "release" -Latest`,
then assign `$version`, `$fileName`, `$url`. The standard pattern writes a
`silent.config` in `postDownloadScript` and installs with
`<filename> /S /CONFIG=silent.config`.
Reference: `Recipes\Jetbrains\pycharm.yaml`, `Recipes\Jetbrains\intellijidea.yaml`.

## Selenium / JS-gated downloads

Last resort, only when the vendor blocks BITS or hides the artifact behind
JavaScript. Launch headless Firefox in `downloadScript` with `Start-SeDriver`,
target `"$buildSpace\$id\$version"` as the download directory, poll for the file,
then `$Driver.Close()`. Reference: `Recipes\keepass.yaml`; the Adobe recipes wrap
this inside `Modules\Custom\AdobeDownloader.psm1` to stay declarative.

## Manual / licensed media (software dropbox)

For vendors whose installer is behind a signed-in licensed download, or whose
license forbids repackaging the public edition.

```yaml
manualDownload: true
manualDownloadFolder: <id>   # only when the folder name differs from id
preDownloadScript: |
  $installer = $dropboxFiles |
    Where-Object { $_.Name -match '<expected-name-regex>' } |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $installer) { throw "No licensed installer found in $dropboxPath." }
  $version  = $Matches['version']
  $fileName = $installer.Name
  $url      = $installer.FullName
```

`SoftwareDropbox` and `SoftwareArchive` come from `preferences.yaml`, never from
the recipe. After a successful upload Yardstick moves the payload to the archive,
which empties the dropbox — **an empty dropbox is the normal steady state**, not a
failure. Reference: `Recipes\Development\freefilesync.yaml`; operator workflow in
`Docs\SoftwareDropbox.md`.

## Uninstall cleanup

Vendor uninstallers routinely leave the final EXE/PDB, an all-users Start Menu
folder, or an install directory behind. Clean up machine-wide application
artifacts in `powerShellUninstallScript`; deliberately preserve per-user settings
and working data. Reference: the tail of `Recipes\Development\marcedit.yaml`.

Parse `UninstallString` defensively — it may be quoted, carry arguments, or be an
msiexec GUID command:

```powershell
$exe = if ($raw -match '^\s*"([^"]+)"') { $Matches[1] }
  elseif ($raw -match '^\s*(.+?\.exe)(\s|$)') { $Matches[1] }
  else { throw "Could not parse uninstall command '$raw'." }
```

Prefer `QuietUninstallString` when present. Stop the app's processes first if it
holds file locks.

## Installers that cannot be automated

Some genuinely cannot, and the right answer is to say so rather than ship a broken
recipe. Nmap's free Windows installer refuses `/S` unless Npcap is present, and
silent Npcap requires the paid OEM edition — so it was dropped in favour of
Sysinternals. FreeFileSync's public edition has no supported silent install, so the
recipe demands licensed Business media. Document the reason in a comment at the top
of the recipe, as those two do.
