# Bot Handoff: Yardstick Recipe Authoring and Test Skill

## Purpose

This folder contains the standalone test harness used to research, author, and prove Yardstick application recipes. This document is the handoff for the next agent that turns the established workflow into a reusable Codex skill.

Do not modify generated `Artifacts\RecipeHarness\_install-tests\<timestamp>\summary.*` files. They are immutable test evidence. Put durable workflow guidance here and in `README.md`.

## Current Layout

- Yardstick workspace: `G:\Intune\YardstickDev`
- Development recipes: `G:\Intune\YardstickDev\Recipes\Development`
- Icons: `G:\Intune\YardstickDev\Icons`
- Runtime/support module: `G:\Intune\YardstickDev\Modules\YardstickSupport.psm1`
- Authoring reference: `G:\Intune\YardstickDev\Docs\WritingRecipes.md`
- Test harness: `G:\Intune\YardstickDev\Tests\RecipeHarness`
- Test evidence: `G:\Intune\YardstickDev\Artifacts\RecipeHarness\_install-tests`
- Software dropbox configuration: `G:\Intune\YardstickDev\Local\preferences.yaml`
- Production deployment target: `G:\Intune\Yardstick` (never use for recipe experimentation or harness output)

The recipes in `Recipes\Development` are commonly ignored by Git while under development. Do not use an empty `git diff` or `git status` result as evidence that no work exists; inspect the files directly.

Defaults now resolve from `Yardstick.Project.psd1`. Prefer explicit `-RecipePath` values, especially for batches; override the other roots only when deliberately testing another workspace.

## Canonical Workflow

1. Read `Docs\WritingRecipes.md`, this handoff, and the harness `README.md` before acting.
2. Inspect related recipes and base recipes before inventing a new pattern.
3. Research only authoritative vendor sources for current download URLs, versions, supported silent switches, prerequisites, and uninstall behavior. Prefer stable landing pages or vendor APIs over hard-coded version URLs.
4. Download once for static analysis when practical. Check Authenticode signatures, file/product versions, archive contents, MSI properties/product codes/upgrade codes, and wrapper extraction switches. Never execute an unknown payload merely to discover its behavior.
5. Create or update the YAML in `Recipes\Development` and ensure the referenced icon exists in `Icons`.
6. Run `Test-RecipeSchema.ps1` for every changed recipe.
7. Batch all safe dry runs into one `Invoke-RecipeInstallTest.ps1 -WhatIf -NoElevate` invocation. `-WhatIf` still resolves URLs, downloads payloads, executes recipe pre/post-download logic, validates signatures and metadata, builds scripts, and checks pre-install detection.
8. When the user authorizes real testing, pass all system-context recipe paths to one harness invocation. The harness delegates them to one elevated child, producing one UAC prompt for the batch.
9. Read the final `summary.md` and each recipe's `result.json`. A successful recipe must pass schema, download, pre-detection, install, post-install detection, uninstall, post-uninstall detection, and residue checks.
10. Delete only temporary analysis copies created for the task. Keep harness summaries and logs.

Recommended commands for the current layout:

```powershell
$harness = 'G:\Intune\YardstickDev\Tests\RecipeHarness'
$yardstick = 'G:\Intune\YardstickDev'
$recipes = "$yardstick\Recipes\Development"
$icons = "$yardstick\Icons"
$targets = @(
    "$recipes\example-one.yaml"
    "$recipes\example-two.yaml"
)

foreach ($target in $targets) {
    & "$harness\Test-RecipeSchema.ps1" -RecipePath $target -RecipesPath $recipes
}

& "$harness\Invoke-RecipeInstallTest.ps1" `
    -RecipePath $targets -WhatIf -NoElevate -DiscardInstaller `
    -RecipesPath $recipes -IconsPath $icons -YardstickDevPath $yardstick

# User-authorized live test: one UAC prompt for the whole system-context batch.
& "$harness\Invoke-RecipeInstallTest.ps1" `
    -RecipePath $targets -DiscardInstaller -InstallTimeoutMinutes 90 `
    -DetectionSettleSeconds 180 -ResidueSettleSeconds 60 `
    -RecipesPath $recipes -IconsPath $icons -YardstickDevPath $yardstick
```

## Important Harness and Installer Lessons

- Do not run a live batch while another installer is active. Windows Installer error `1618` means another MSI transaction owns the global lock. Let the unrelated installer and the current harness exit before retrying; do not immediately alter a recipe based on a collision.
- Keep the parent harness process alive. Output from the elevated child may not stream until it exits, but progress is visible under the current artifact directory.
- Do not start a second harness while the first elevated child is still running.
- Avoid `-Force` unless the user explicitly accepts touching software that predates the test. The default pre-detection guard protects existing installations.
- `-DiscardInstaller` removes retained payloads, not evidence. Use it for large batches.
- Exit `1603` requires evidence. Inspect verbose MSI logs and MSI database tables such as `LaunchCondition`, `Property`, `CustomAction`, and `InstallExecuteSequence`; do not assume every `1603` has the same cause.
- Vendor EXE bundles may be better packaged by extracting their signed MSI, external CAB files, and prerequisites. Preserve the exact relative layout expected by the MSI.
- An extracted Advanced Installer MSI may require bootstrapper properties such as `SETUPEXEDIR`. Dragonframe is the validated example.
- A ZIP or wrapper may have `UnknownError` signature status even when its inner executable/MSI is correctly signed. Record the warning and validate the actual executable payload.
- `postDownloadScript` may change `$fileName`; all later placeholders and metadata must refer to the final installer.
- Prefer MSI detection/uninstall when a stable product code is available. Use script detection for products with dynamic product codes, multiple registrations, or vendor-specific footprints.
- Detection must be false before installation, true after installation, and false after uninstall. Residue warnings should be understood, not hidden.
- Manual-download recipes must use `manualDownload: true`, `manualDownloadFolder`, and the configured `SoftwareDropbox`. Never substitute an unlicensed public installer for licensed deployment media.
- Some installers spawn UI despite nominal silent switches. Observe processes and windows during a test and verify the command from vendor documentation.

## High-Value Reference Recipes and Evidence

- `dragonframe.yaml`: dynamic scrape of the official downloads page; extracts an Advanced Installer bundle into MSI, CAB, and Visual C++ prerequisite; supplies required `SETUPEXEDIR`; uses MSI detection and uninstall. Version `2026.07.4` passed the complete live cycle with zero residue. Evidence: `Artifacts\RecipeHarness\_install-tests\20260918-103757\summary.md`.
- `marcedit.yaml`: managed/composite installer; extracts the x64 MSI; packages x64 and x86 .NET Desktop Runtime prerequisites plus WebView2; uses script detection and custom cleanup. Version `7.8.25` passed the complete live cycle with zero residue. Evidence: `Artifacts\RecipeHarness\_install-tests\20260918-102905\summary.md`.
- `androidstudio.yaml`: useful complex EXE reference. Version `261.26222.65.2614.16204760` passed the complete live cycle. Evidence is under the latest `20260917-*` passing run.
- `azurevirtualdesktopagent.yaml` and `azurevirtualdesktopagentbootloader.yaml`: paired Intune Remote Help prerequisites sourced from Microsoft redirects. The Agent version `1.0.15294.200` passed live testing; both recipes passed WhatIf together. The Bootloader has not completed a standalone live cycle and should be deployed with the Agent as its Intune dependency.
- `freefilesync.yaml`: deliberately requires licensed Business Edition media from the software dropbox because the public edition does not support the required silent deployment. It has no live result without staged licensed media.
- `treesize.yaml`: kept as a completed recipe, but live installation failed on this Windows Server test box because that environment is unsupported. The user intends to test it elsewhere; do not “fix” it from this server-only failure.

Recent live outcomes that a future skill should preserve as examples of honest reporting:

- Pass: Calibre `9.14.0`, NordVPN `8.11.1.0`, Cryptomator `1.19.3`.
- Pass with warning: StartAllBack `3.9.25` left `C:\Program Files\StartAllBack`; WD Security and WD Drive Utilities `2.1.6.511` passed while the outer ZIP signature produced a download warning.
- Environment/vendor failure requiring future work: Logitech Options+ `2.6.944893` returned install exit `1008` and its uninstall process faulted.
- WD Discovery did open a window during research, which the user closed, but no finished `wddiscovery.yaml` exists. Do not claim it was completed.
- Dell Display Manager was explicitly removed from scope.

## Proposed Codex Skill Shape

The next agent should use the `skill-creator` skill and build a focused Yardstick recipe-authoring skill, not a generic Windows packaging skill.

Suggested contents:

- `SKILL.md`: triggers, scope, safety boundaries, canonical workflow, batching/UAC rules, and completion criteria.
- `references/yardstick-fields.md`: route to `Docs\WritingRecipes.md`; summarize only the field combinations and placeholder rules the agent must choose correctly.
- `references/installer-patterns.md`: MSI, EXE, archive, bootstrapper extraction, manual licensed media, dependencies, dynamic detection, and cleanup patterns, anchored to the reference recipes above.
- `references/test-harness.md`: schema, WhatIf, full live cycle, output interpretation, installer-lock handling, and evidence requirements.
- Optional `scripts/` wrapper: accept explicit recipe paths and run schema/dry/live batches against the current configurable roots. It must never silently add `-Force`, `-SkipUninstall`, or real installation.

The skill should require a concise final report containing recipe paths, resolved versions, authoritative vendor sources, schema results, WhatIf results, live results or clearly stated test limitations, warnings/residue, and artifact-summary paths. It should never mark a recipe complete merely because schema or download stages pass.
