---
name: yardstick-recipes
description: Author, update, and prove Yardstick application recipes (Recipes/*.yaml) that package Windows software for Intune. Use whenever the request is to add/package/deploy an application to Intune via Yardstick, write or fix a recipe YAML, choose a detection rule or silent install switch, debug a failing download/version scrape, or run the recipe test harness (Test-RecipeSchema.ps1, Invoke-RecipeLab.ps1, Invoke-RecipeInstallTest.ps1). Triggers on "recipe", "Yardstick", "package <app> for Intune", "add <app> to Intune", "detection rule", "silent install", "intunewin", "software dropbox".
---

# Writing Yardstick Recipes

A recipe is one YAML file that teaches Yardstick how to find, download, package,
install, detect, and uninstall one Windows application. Yardstick then uploads it
to Intune as a Win32 app and manages versions, assignments, and supersedence.

Your job is to produce a recipe that **provably works**, not one that merely parses.

## Scope and boundaries

- Work in `G:\Intune\YardstickDev`. New and in-progress recipes go in
  `Recipes\Development\` (Git-ignored — never treat an empty `git status` as
  evidence that no work exists; inspect the files).
- `G:\Intune\Yardstick` is the **production** deployment target. Never write
  recipes, test output, or experiments there. Promotion happens only through
  `Publish-Yardstick.ps1`, and only when the user asks.
- Never run `Yardstick.ps1` itself for testing — it uploads to the live Intune
  tenant. Use the harness under `Tests\RecipeHarness\`, which never contacts Graph.
- Never substitute a public/unlicensed installer for licensed deployment media.
  Licensed media belongs in the software dropbox (`manualDownload: true`).
- Never execute an unknown downloaded payload just to see what it does. Inspect it
  statically first (signature, version info, MSI tables, extraction switches).

## Workflow

### 1. Research the vendor first

Read `references/fields.md` and `references/installer-patterns.md` before writing
YAML. Then find, from **authoritative vendor sources only**:

- a stable download URL or a stable landing page / API that exposes the current one
  (prefer these over a hard-coded versioned URL, which goes stale in weeks),
- the documented silent install switches and their exit codes,
- the documented silent uninstall path,
- prerequisites the installer expects to already be present,
- where the product registers itself (ARP key, install path, file version).

Look at sibling recipes before inventing a pattern. `Recipes\` and
`Recipes\Development\` hold ~200 worked examples; one of them almost certainly
already solved this installer family.

### 2. Download once and analyse statically

Fetch the installer a single time and inspect it: Authenticode signature, SHA256,
`VersionInfo.ProductVersion`, archive contents, MSI `ProductCode`/`UpgradeCode`/
`Property`/`LaunchCondition`, and whether an EXE wrapper supports `/extract`.
This is where you learn the real version format and the real detection target.

### 3. Write the recipe

Start from `Templates\MSITemplate.yaml` or `Templates\EXETemplate.yaml`. Fill in
only what the vendor actually needs; everything else inherits from
`Local\preferences.yaml`. The filename, the `id` field, and the display-name prefix
must line up — `id` must match the filename exactly (case-sensitive, lowercase).

Ensure the icon exists in `Icons\` under the exact `iconFile` name. If none exists,
`Tests\RecipeHarness\Export-ExeIcon.ps1 -ExePath <installer> -OutputPath Icons\<id>.png`
pulls one out of the executable.

### 4. Validate schema

```powershell
& 'G:\Intune\YardstickDev\Tests\RecipeHarness\Test-RecipeSchema.ps1' -Id <id>
```

Run it for every recipe you changed. Fix warnings too — a casing warning means
Yardstick may read a different field than you think you set.

### 5. Dry run (safe, no install)

Batch every changed recipe into **one** invocation:

```powershell
& 'G:\Intune\YardstickDev\Tests\RecipeHarness\Invoke-RecipeInstallTest.ps1' `
    -Id <id1>, <id2> -WhatIf -NoElevate -DiscardInstaller
```

`-WhatIf` still resolves redirects, runs your pre/post-download scripts, downloads
the payload, checks signature and metadata, builds the install/detection scripts,
and runs pre-install detection. It suppresses only the install and uninstall
commands. This catches most recipe bugs without touching the machine.

### 6. Live test — only with explicit user authorization

This installs and uninstalls real software on the current machine. **Ask first.**

```powershell
& 'G:\Intune\YardstickDev\Tests\RecipeHarness\Invoke-RecipeInstallTest.ps1' `
    -Id <id1>, <id2> -DiscardInstaller -InstallTimeoutMinutes 90 `
    -DetectionSettleSeconds 180 -ResidueSettleSeconds 60
```

Pass **all** system-context recipes to one invocation so the batch produces a
single UAC prompt. Never start a second harness while an elevated child is running,
and never start a batch while an unrelated installer is active.

### 7. Read the evidence, then report

Read `summary.md` and each recipe's `result.json` under
`Artifacts\RecipeHarness\_install-tests\<timestamp>\`. See
`references/test-harness.md` for phase and exit-code interpretation.

## Completion criteria

A recipe is **validated** only when all of these passed in one live run:

schema → download → pre-install detection (Not Detected) → install →
post-install detection (Detected) → uninstall → post-uninstall detection
(Not Detected) → residue check.

Anything short of that is reported as a stated limitation, never as success.
"Schema passes" and "it downloaded" are not completion. If a recipe fails for an
environmental reason (unsupported test OS, licensed media not staged, another MSI
holding the installer lock), say exactly that rather than editing the recipe to
chase a symptom.

## Final report

Always close with:

- recipe path(s) and resolved version(s)
- the authoritative vendor source used for URL, version, and switches
- schema result, `-WhatIf` result, live result **or** the explicit limitation
- any warnings and any residue left behind, with your read on whether it matters
- the path to the harness summary directory

## References

- `references/fields.md` — required fields, script-block order and scoping,
  detection-type decision table, placeholders, supersedence, common pitfalls.
- `references/installer-patterns.md` — MSI, EXE, InnoSetup/NSIS, bootstrapper
  extraction, ZIP/portable, prerequisites, GitHub/JetBrains downloaders, manual
  licensed media, dynamic detection, uninstall cleanup — each anchored to a real recipe.
- `references/test-harness.md` — harness parameters, phases, output layout,
  installer-lock and exit-code handling, evidence rules.
- Full field documentation: `Docs\WritingRecipes.md`.
- Dropbox operator workflow: `Docs\SoftwareDropbox.md`.
- Harness handoff notes and known-good reference recipes: `Tests\RecipeHarness\BOT-HANDOFF.md`.
