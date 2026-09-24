# Test Harness

`G:\Intune\YardstickDev\Tests\RecipeHarness\`. None of it contacts Graph or Intune.
Defaults resolve from `Yardstick.Project.psd1`, so `-Id <id>` alone works against
`Recipes\Development\`; override roots only when deliberately testing elsewhere.

Requirements: `powershell-yaml` ≥ 0.4.12, `BitsTransfer`, and — only for
`Invoke-RecipeLab.ps1 -Stage Package` — `IntuneWin32App` 1.5.x.

## The three scripts

### Test-RecipeSchema.ps1

Schema, `id`-matches-filename (case-sensitive), icon existence, enum validity,
casing warnings, reserved-variable warnings. Resolves `base:` inheritance first.
Seconds to run; run it on every changed recipe.

```powershell
& "$harness\Test-RecipeSchema.ps1" -Id <id>
& "$harness\Test-RecipeSchema.ps1" -RecipePath <path>
```

### Invoke-RecipeLab.ps1

Single-recipe, staged, no install.

- `-Stage Schema` (default) — delegates to `Test-RecipeSchema.ps1`.
- `-Stage Download` — plus scripts, download, signature, hash, installer metadata.
- `-Stage Package` — plus local `.intunewin` packaging.

Output: `Artifacts\RecipeHarness\<id>\<timestamp>\`.

Use this when iterating on one recipe's download/version logic — it is faster and
noisier than the install harness, and `-Stage Package` is the only way to prove
the `.intunewin` actually builds.

### Invoke-RecipeInstallTest.ps1

The real proof. Phases, in order:

| # | Phase | What it proves |
|---|---|---|
| 1 | Schema | same rules as above |
| 2 | Download | redirects, pre/download/post scripts, installer present, Authenticode, SHA256, MSI ProductCode |
| 3 | PreDetect | detection reports **Not Detected** before install |
| 4 | Install | `installScript` / `powerShellInstallScript` through cmd.exe |
| 5 | PostDetect | detection reports **Detected** |
| 6 | Reinstall | optional idempotency pass (`-IncludeReinstall`) |
| 7 | Uninstall | `uninstallScript` / `powerShellUninstallScript` |
| 8 | PostUninstall | detection reports **Not Detected** again |
| 9 | Residue | registry and directory leftovers |

Selection: `-Id <id>...`, `-RecipePath <path>...` (arrays accepted), or `-All`.

Key switches:

- `-WhatIf` — runs phases 1-3 and builds the scripts, but suppresses only the
  install and uninstall commands. Safe. Always do this before a live run.
- `-NoElevate` — skip system-context recipes rather than launching an elevated
  child. Pair with `-WhatIf` for the dry batch.
- `-DiscardInstaller` — delete payloads after each recipe, keep all evidence. Use
  it for any batch; installers add up to gigabytes.
- `-IncludeReinstall` — prove idempotency.
- `-InstallTimeoutMinutes` (default 30), `-DetectionSettleSeconds` (default 120),
  `-ResidueSettleSeconds` (default 30) — raise for slow installers; silent
  installers frequently return before they have finished writing.
- `-Force` — continue when the app is already detected pre-install. **Only with
  explicit user consent**; without it the harness refuses to touch software it did
  not install.
- `-SkipUninstall` — leaves software installed. One-off debugging only.

Exit code `1` when any recipe failed.

## Batching and UAC

Pass every system-context recipe to **one** invocation. The harness delegates the
whole set to a single elevated child, producing one UAC prompt. Start unelevated.

Do not start a second harness while the first elevated child is running. Output
from the child may not stream until it exits — progress is visible in the artifact
directory in the meantime. Keep the parent process alive.

## Output

```
Artifacts\RecipeHarness\<id>\<timestamp>\                  # lab runs
Artifacts\RecipeHarness\_install-tests\<timestamp>\        # install-test runs
    summary.md
    summary.json
    harness.transcript.log
    <id>\result.json, command files, detection scripts, raw stdout/stderr
```

These are immutable test evidence. Read them; never edit them. Delete only
temporary analysis copies you created yourself.

## Reading failures

- **MSI 1618** — another Windows Installer transaction holds the global lock. This
  is environmental. Let the other installer and the harness exit, then retry. Do
  **not** change the recipe in response to a 1618.
- **MSI 1603** — generic failure with many causes. Get evidence before theorising:
  verbose MSI log, and the MSI's `LaunchCondition`, `Property`, `CustomAction`, and
  `InstallExecuteSequence` tables. Advanced Installer payloads extracted from a
  bootstrapper commonly need `SETUPEXEDIR`.
- **Signature `UnknownError` on a ZIP or wrapper** — often expected; the inner
  EXE/MSI may be correctly signed. Record the warning and validate the real payload.
- **PostDetect fails but the app is clearly installed** — usually a version-format
  mismatch (3-part vs 4-part) or a detection path the installer writes only
  conditionally. Raise `-DetectionSettleSeconds` before assuming a logic bug.
- **Install spawns UI despite a silent switch** — watch processes and windows
  during the run and re-check the switch against vendor documentation.
- **Residue warnings** — understand them, don't suppress them. Some are acceptable
  (per-user settings by design); machine-wide program directories are not.

## Reporting honestly

Historical examples worth imitating, from `BOT-HANDOFF.md`:

- Pass: Calibre 9.14.0, NordVPN 8.11.1.0, Cryptomator 1.19.3.
- Pass with warning: StartAllBack 3.9.25 left `C:\Program Files\StartAllBack`
  behind; WD Security / WD Drive Utilities 2.1.6.511 passed while the outer ZIP
  signature produced a download warning.
- Environment/vendor failure, future work: Logitech Options+ 2.6.944893 returned
  install exit 1008 and its uninstaller faulted.
- TreeSize live install failed on the Windows Server test box because that
  environment is unsupported — the recipe was *not* "fixed" from a server-only
  failure.

Never call a recipe complete because schema or download passed.
