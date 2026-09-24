# Yardstick Recipe Test Harness

PowerShell harnesses for validating Yardstick recipe YAML from schema checks through local installation, detection, uninstallation, and residue testing.

The harness is part of the `G:\Intune\YardstickDev` monorepo:

- Harness source: `Tests\RecipeHarness`
- Development recipes: `Recipes\Development`
- Icons: `Icons`
- Runtime/support code: project root
- Generated evidence: `Artifacts\RecipeHarness`
- Machine-local preferences: `Local\preferences.yaml`

Defaults are resolved from `Yardstick.Project.psd1`. Every path can still be overridden through parameters.

## Requirements

- Windows PowerShell or PowerShell 7.
- `powershell-yaml` version 0.4.12 or newer.
- `IntuneWin32App` version 1.5.x only for `Invoke-RecipeLab.ps1 -Stage Package`.
- `BitsTransfer`.
- The project runtime with `Modules\YardstickSupport.psm1`.

The install harness runs real installers on the local machine. Use a disposable Windows VM when possible.

## Quick Start

```powershell
$project = 'G:\Intune\YardstickDev'
$harness = "$project\Tests\RecipeHarness"
$recipe = "$project\Recipes\Development\obsidian.yaml"

# Schema and icon validation only
& "$harness\Test-RecipeSchema.ps1" -RecipePath $recipe

# Schema, scripts, download, signature, hash, and installer metadata
& "$harness\Invoke-RecipeLab.ps1" -RecipePath $recipe -Stage Download

# Package locally as .intunewin; no Graph or Intune upload
& "$harness\Invoke-RecipeLab.ps1" -RecipePath $recipe -Stage Package

# Full preparation and detection preview without install/uninstall commands or UAC
& "$harness\Invoke-RecipeInstallTest.ps1" -RecipePath $recipe `
    -WhatIf -NoElevate -DiscardInstaller

# User-authorized live install/detect/uninstall/residue cycle
& "$harness\Invoke-RecipeInstallTest.ps1" -RecipePath $recipe -DiscardInstaller
```

## Common Parameters

- `-Id <id>` selects `<RecipesPath>\<id>.yaml`.
- `-RecipePath <path>` tests explicit recipe paths and accepts arrays in the install harness.
- `-YardstickDevPath <path>` overrides the runtime root.
- `-RecipesPath <path>` overrides the recipe directory.
- `-IconsPath <path>` overrides the icon directory.
- `-ArtifactsPath <path>` overrides the evidence directory.

`Invoke-RecipeInstallTest.ps1` also supports:

- `-All` to test every YAML file in `-RecipesPath`.
- `-IncludeReinstall` to run a second installation before uninstalling.
- `-SkipUninstall` for one-off debugging; this leaves software installed.
- `-DiscardInstaller` to remove payloads while retaining logs, hashes, and results.
- `-Force` to continue when detection already passes before installation.
- `-NoElevate` to skip system-context recipes instead of launching an elevated child.

Manual-download recipes read `SoftwareDropbox` from `Local\preferences.yaml`. The harness copies staged media into its artifact directory and never archives, moves, or modifies the dropbox source.

## Output

Download/package runs write to:

```text
Artifacts\RecipeHarness\<id>\<timestamp>\
```

Install-test runs write to:

```text
Artifacts\RecipeHarness\_install-tests\<timestamp>\
```

Each install-test run includes `summary.md`, `summary.json`, `harness.transcript.log`, and per-recipe `result.json`, command files, detection scripts, and raw stdout/stderr logs. The install harness exits `1` when any recipe fails.

## Safety Notes

- The harness never contacts Graph or Intune.
- `Invoke-RecipeInstallTest.ps1` does install and uninstall applications locally.
- Start live tests unelevated. System-context recipes are delegated to one elevated child.
- Batch system-context recipe paths in one invocation to minimize UAC prompts.
- `-WhatIf` still performs schema, download, extraction, metadata, signature, and pre-detection checks; it suppresses only install and uninstall commands.
- Avoid `-Force` unless the user explicitly accepts touching a pre-existing installation.
- Never start a live batch while another installer or harness child is active. MSI exit `1618` is an environmental collision.
- Investigate MSI exit `1603` with verbose logs and MSI database conditions instead of guessing.

## Bot Checklist

1. Read `BOT-HANDOFF.md` and `Docs\WritingRecipes.md`.
2. Confirm each recipe and icon exists.
3. Run schema validation for every changed recipe.
4. Batch `-WhatIf -NoElevate` checks.
5. Ensure no unrelated installer is active before a user-authorized live batch.
6. Batch system-context live tests to minimize UAC prompts.
7. Read `summary.md`, `summary.json`, and each `result.json`.
8. Call a recipe fully validated only after install, detection, uninstall, post-uninstall detection, and residue phases pass—or state the exact limitation.
