# Native-bitness guard — outstanding test coverage

Tracks validation still owed for the change that makes every generated
`install.ps1` / `uninstall.ps1` re-launch itself under native PowerShell
(`Add-NativeBitnessGuard` in `Yardstick.ps1`, mirrored in
`Tests/RecipeHarness/RecipeTestSupport.psm1`).

**Status: the mechanism is proven; per-recipe coverage is not.**

## What has already been validated

| Claim | Evidence |
| --- | --- |
| The guard re-launches natively from a 32-bit host | Generated `.cmd` run through `SysWOW64\cmd.exe` reported `Is64BitProcess=True` |
| It restores the native registry view | Same run counted **229** `Uninstall` subkeys (native) rather than **411** (WOW6432Node) |
| It is inert under a 64-bit caller | Identical results through `System32\cmd.exe` |
| It propagates the child exit code | Deliberate `exit 42` surfaced as 42 from both hosts |
| It preserves the working directory | `.\payload-marker.txt` resolved from both hosts — install scripts reference the payload relatively |
| It does not regress a real recipe | `r.yaml` 4.6.1, `-AgentBitness x86`: schema → download → PreDetect → install → PostInstallDetect → uninstall → PostUninstallDetect → residue 0/0, all Pass |
| It is actually injected | Read back from the generated `uninstall.ps1` in the harness buildspace |

Note the `r.yaml` run demonstrates **no regression**, not guard necessity: R's
uninstaller has an `unins000.exe` fallback built from `$env:ProgramFiles`, which
is inherited verbatim rather than rewritten by WOW64, so that run could have
succeeded without the guard. The 229-vs-411 result is the isolating evidence.

## What is still untested

The guard now applies to **87 recipes** using `powerShellUninstallScript` and
**38** using `powerShellInstallScript`. Exactly one of them (`r`) has been run.

### Priority 1 — known regression risk

The guard *changes* behaviour for any recipe that implicitly relied on running
32-bit. `[Environment]::GetFolderPath('ProgramFiles')` is the dangerous call: it
returns `C:\Program Files (x86)` in a 32-bit process and `C:\Program Files` in a
64-bit one. A static sweep of all `powerShell*Script` recipes found one user:

- **`actilife.yaml`** — uses it in two places in `powerShellUninstallScript`:
  - `$driverDir = Join-Path ([Environment]::GetFolderPath('ProgramFiles')) 'ActiGraph\Drivers'`
  - the `ActiGraph` entry in `$pruneRoots`

  ActiLife appears to be an x86-registered app (the uninstaller fallback searches
  `ProgramFilesX86` first). If its drivers live under `C:\Program Files (x86)`,
  then post-change these two lookups point at the wrong directory. Both are
  `Test-Path`-guarded and the final verdict is ARP-based, so **the uninstall
  would still report success while silently skipping the driver uninstall and
  leaving the directory behind.** The harness Residue phase should catch it.

  Test: `-Id actilife`, compare `-AgentBitness x86` against the pre-change
  behaviour, and check the Residue phase for a leftover `ActiGraph` directory.
  If confirmed, pin both lines to `GetFolderPath('ProgramFilesX86')` rather than
  reverting the guard.

Recipes naming `ProgramFilesX86` / `Program Files (x86)` *explicitly* are not at
risk — that resolves identically in both bitnesses. For the record, those are:
`axurerp`, `brave`, `epanet`, `epsonscanperfectionv600`, `evernote`,
`ghostscript`, `klitecodecpack`, `marcedit`, `mathtype`, `meetingowl`, `mgear`,
`openscad` (plus `actilife`).

### Priority 2 — recipes that previously could not be tested at all

These hand-write a `%SystemRoot%\sysnative\...` command line, which fails under a
64-bit caller, so they were untestable in the harness before `-AgentBitness`
existed. They now run, and their hand-rolled sysnative wrapper is redundant with
the injected guard:

`qgis`, `virtualbox`, `eclipse`, `endnote`, `ltspice`, `vivaldi`, `clo3d`

`clo3d`, `eclipse`, `virtualbox`, `vivaldi` set an explicit `uninstallScript`, so
their command line is unchanged; only the `.ps1` content gained the preamble.

Test: run each at the `x86` default. On success, consider simplifying them the
way `r.yaml` was simplified — drop the hand-rolled sysnative wrapper and let the
guard do it, so they stay testable under both hosts.

### Priority 3 — the remaining bulk

The other ~75 `powerShellUninstallScript` recipes. Most are expected to be
unaffected, but the `x86` default is new, so this is the first time any of them
will be exercised the way Intune actually runs them. Batch them:

```powershell
# System-context recipes must go in ONE invocation — single UAC prompt.
& 'Tests\RecipeHarness\Invoke-RecipeInstallTest.ps1' `
    -Id <ids...> -DiscardInstaller -InstallTimeoutMinutes 90 `
    -DetectionSettleSeconds 180 -ResidueSettleSeconds 60
```

Expect two distinct failure shapes, and do not conflate them:

- **Fails at `x86`, passes at `-AgentBitness x64`** → a genuine 32-bit
  dependency the guard did not resolve. Investigate the recipe.
- **Fails at both** → unrelated to this change (vendor URL drift, licensed media
  not staged, unsupported test OS).

## Known cosmetic effects

- The `.ps1` in the buildspace is no longer byte-identical to the recipe YAML.
  Any error message reporting a line number is offset by the preamble.
- `Set-Content -Encoding UTF8` under Windows PowerShell 5.1 writes a BOM ahead of
  the preamble. `powershell.exe -File` handles it; noted so nobody treats it as
  corruption.

## Unrelated hazard found while doing this work

`Publish-Yardstick.ps1` line 103 lists `Yardstick.ps1` in the `-Runtime` publish
set, implying a dev → production flow. At the time of writing,
`YardstickDev\Yardstick.ps1` was **3,627 diff lines and 4 days behind**
production `HEAD`. Running `-Runtime` from that tree would regress production.
Reconcile the two copies, or stop publishing `Yardstick.ps1` that way, before
anyone uses `-Runtime` again.
