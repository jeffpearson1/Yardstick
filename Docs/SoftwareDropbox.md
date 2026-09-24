# The Yardstick Software Dropbox

## Overview

Most Yardstick recipes fetch their installer themselves, from a URL or a
`downloadScript`. Some vendors make that impossible: the download is gated behind a
signed-in, licensed account, and no amount of scripting gets past an entitlement
check. CLO 3D is the reference case — the version list is public, but the download
endpoint returns 401 without a licensed account.

The software dropbox is the escape hatch for those applications. A human fetches the
installer once and drops it into a folder; Yardstick packages whatever is sitting
there and publishes it to Intune exactly like any other recipe. Everything downstream
of the download — versioning, detection rules, supersedence, assignments, retention,
backup, email reporting — is unchanged.

Recipes that use this are called **manual-drop recipes**, and they are marked with
`manualDownload: true`.

## The two folders

| Folder | Set in | Purpose |
|--------|--------|---------|
| `SoftwareDropbox` | `preferences.yaml` | Where an operator stages an installer waiting to be packaged |
| `SoftwareArchive` | `preferences.yaml` | Where Yardstick moves the payload after Intune accepts the upload |

Both are roots, not per-application folders. Yardstick appends the recipe id itself:

```
<SoftwareDropbox>\
  clo3d\                        <- drop the installer here
    CLO_Windows_NetworkOnlineAuth_x64_2026.1.188.58419.exe

<SoftwareArchive>\
  clo3d\
    2026.1.188.58419\           <- moved here automatically after a successful run
      CLO_Windows_NetworkOnlineAuth_x64_2026.1.188.58419.exe
    2026.1.180.57712\           <- the previous drop
      ...
```

A network share is fine for both, and is usually what you want — the dropbox is the
handoff point between whoever has the vendor login and the machine that runs
Yardstick.

## Configuration

In `preferences.yaml`:

```yaml
SoftwareDropbox: \\fileserver\share\Intune\SoftwareDropbox
SoftwareArchive: \\fileserver\share\Intune\Apps
```

Leave both blank if you run no manual-drop recipes. These paths live **only** in
`preferences.yaml` — a recipe never names a dropbox path.

If a recipe sets `manualDownload: true` and `SoftwareDropbox` is blank, that recipe
fails immediately with `Recipe sets manualDownload but no SoftwareDropbox path is
configured in preferences.yaml.`, reported under the failure stage **Manual Drop**.

If `SoftwareArchive` is blank, packaging and upload still work, but the payload is
left in the dropbox with a warning in the log. See
[Archiving is what stops a republish](#archiving-is-what-stops-a-republish) for why
that matters.

## Staging a new version (the operator workflow)

1. Download the installer from the vendor with whatever licensed account is required.
2. Confirm the target folder is empty. `<SoftwareDropbox>\<recipe id>` should contain
   nothing but a `.gitkeep`, if anything — Yardstick empties it after every successful
   publish, so anything left over is either a failed run or a drop nobody finished.
3. Copy the installer in. Create the folder if it does not exist.
4. Run the recipe:

   ```powershell
   .\Yardstick.ps1 -AppId clo3d
   ```

   Or just let the next scheduled `-All` run pick it up — a staged payload is
   packaged on the next run that touches the recipe, no separate command needed.
5. Check the log or the email report. On success the payload has moved to
   `<SoftwareArchive>\<recipe id>\<version>\`, and the dropbox folder is empty again.

You can stage more than one file. Everything in the folder — files and subfolders
alike — is copied into the buildspace, so an installer plus a licence file, a config
XML or a transform all work. The recipe decides which of those is the installer.

## What happens during a run

1. **Before the pre-download script**, Yardstick reads the dropbox folder and exposes
   what it finds to the recipe's script blocks.
2. If the folder is missing or empty, the recipe is **skipped** — this is the normal
   steady state, not an error (see below).
3. The `preDownloadScript` inspects the staged files and sets `$version` and
   `$fileName`.
4. Version validation runs as usual. A version already in Intune, or one excluded by
   `versionLock`, stops the run here.
5. The entire dropbox folder is copied (recursively) into
   `<BuildSpace>\<id>\<version>` in place of a download.
6. Packaging, detection rules and the Intune upload proceed normally.
7. **After Intune accepts the upload**, the payload is moved to
   `<SoftwareArchive>\<id>\<version>`, emptying the dropbox.

Step 7 runs only on success. A failed run leaves the payload exactly where it is, so
you can fix the problem and re-run without staging the file again.

## Writing a manual-drop recipe

Two fields control this:

| Field | Default | Description |
|-------|---------|-------------|
| `manualDownload` | `false` | `true` makes the recipe take its payload from the dropbox instead of a URL or `downloadScript` |
| `manualDownloadFolder` | the recipe `id` | Only set this when the dropbox subfolder should differ from the recipe id — for example when two recipes share one staged payload |

Inside `preDownloadScript`, `downloadScript` and `postDownloadScript`, two extra
variables are available:

| Variable | Description |
|----------|-------------|
| `$dropboxPath` | The staging folder, `<SoftwareDropbox>\<id>` |
| `$dropboxFiles` | `FileInfo` / `DirectoryInfo` objects for everything staged there (`.gitkeep` excluded) |

A manual-drop recipe still has to set `$version` and `$fileName` in its
`preDownloadScript`, the same as any other recipe — Yardstick has no way to guess
them from a filename. The usual pattern is to pick the installer out of
`$dropboxFiles`, parse the version out of its name, and fall back to the file's
version resource:

```yaml
manualDownload: true
preDownloadScript: |
  $installer = @($dropboxFiles | Where-Object Extension -eq '.exe' | Sort-Object LastWriteTime -Descending)[0]
  if (-not $installer) { throw "No .exe staged in $dropboxPath" }
  $fileName = $installer.Name
  if ($fileName -match '(\d+\.\d+\.\d+\.\d+)') {
    $version = $matches[1]
  } else {
    $version = $installer.VersionInfo.ProductVersion
  }
  if (-not $version) { throw "Could not determine a version for $fileName" }
```

Throwing from the pre-download script is the right move when the staged files are not
what the recipe expects — it fails that one recipe with a clear message rather than
publishing a broken package.

`Recipes/clo3d.yaml` is the complete worked example, including a detection rule that
reads the *installed* binary's version resource rather than the download's build
number. See [WritingRecipes.md](WritingRecipes.md) for the full recipe reference.

## An empty dropbox is the normal state

This is the part worth internalising: **the dropbox being empty is how Yardstick knows
there is nothing new to publish.** There is no version check against a vendor site for
a manual-drop recipe, because there is no vendor site it can reach. The presence of a
file is the signal.

So on a normal `-All` run, a manual-drop recipe with an empty dropbox logs

```
Nothing staged in the software dropbox for clo3d. Skipping update.
```

and moves on. That is success, not a warning. Under `-Repair` the same recipe still
gets its full maintenance sweep — retention pruning, `(N-x)` renaming, assignment
migration, supersedence — so nothing rots just because no one has staged a new build
in months.

### Archiving is what stops a republish

Because presence is the signal, leaving a payload in the dropbox after a successful
publish would make every subsequent run try to package it again. Those runs would not
actually republish — version validation catches that the version is already in Intune
and skips — but you would get a permanent "already in the repo" skip on every run, and
you would lose the ability to tell "nothing new" from "something staged and waiting".

That is why `SoftwareArchive` should be set whenever you use the dropbox. Without it,
Yardstick logs a warning and leaves the payload in place, and clearing the folder
becomes your job.

## Common situations

| Situation | What happens |
|-----------|--------------|
| Dropbox folder missing or empty | Recipe is skipped (or maintenance-swept under `-Repair`). Normal. |
| `SoftwareDropbox` not configured | Recipe fails, stage **Manual Drop**, with an explicit message. |
| `SoftwareArchive` not configured | Upload succeeds; payload stays in the dropbox with a logged warning. |
| Staged version is already in Intune | Skipped as "already in the repo", and the payload is **not** archived — clear it by hand, or re-run with `-Force` to republish. |
| Staged version is older than what is published | Skipped as older. Payload stays put. |
| Run fails anywhere before upload | Payload stays in the dropbox. Fix and re-run; no need to re-stage. |
| Archiving fails (share offline, permissions) | Logged as a warning only. The application update still counts as successful; clear the dropbox by hand. |
| Same version archived twice (a `-Force` re-run, or a corrected payload) | The second archive folder gets a timestamp suffix rather than overwriting the first. |
| Several files staged together | All of them, including subfolders, are copied into the buildspace and packaged. |

## Adding a new manual-drop application

1. Create `<SoftwareDropbox>\<recipe id>\` on the share.
2. Write the recipe with `manualDownload: true` and a `preDownloadScript` that derives
   `$version` and `$fileName` from `$dropboxFiles`.
3. Add the icon to the icon cache, as with any recipe.
4. Stage the installer and run `.\Yardstick.ps1 -AppId <recipe id>` to verify.
5. Tell whoever holds the vendor login where to drop future installers — that person
   does not need Yardstick access, only write access to one folder.
