# Recipe evidence

A recipe can only be merged into `main` if the pull request carries proof that
it passed a live install test with exactly the contents being merged. That proof
is an **evidence file**. A pull request check, **Recipe evidence**, reads it.

## The evidence file

Yardstick Development Studio writes `Tests/RecipeEvidence/<id>.json` when you
publish a recipe, and commits it next to the recipe:
- `<id>` is the recipe file name without `.yaml`, so `Recipes/Jetbrains/clion.yaml`
  has `Tests/RecipeEvidence/clion.json`.
- The folder is outside `Recipes/`, so nothing that scans recipes picks it up.
- Nothing in Yardstick's runtime, the recipe harness or `Publish-Yardstick.ps1`
  reads `Tests/`, so evidence never reaches production.

The file is UTF-8 JSON:

```json
{
  "schemaVersion": 1,
  "recipeId": "7zip",
  "recipePath": "Recipes/7zip.yaml",
  "recipeSha256": "dac9a5b6e5a249a1a674bd890731d93ada96e9fcb4b480a3659416695f52486b",
  "validated": true,
  "verdict": "Validated",
  "reason": "",
  "testedAt": "2026-11-09T10:00:00Z",
  "studioVersion": "0.5.0",
  "installExperience": "system",
  "installer": { "fileName": "7z2408-x64.msi", "version": "24.08", "sha256": "…", "signer": "Igor Pavlov" },
  "phases": [
    { "name": "Schema", "status": "Pass", "detail": "" },
    { "name": "Download", "status": "Pass", "detail": "" },
    { "name": "PreDetect", "status": "Pass", "detail": "" },
    { "name": "Install", "status": "Pass", "detail": "" },
    { "name": "PostInstallDetect", "status": "Pass", "detail": "" },
    { "name": "Uninstall", "status": "Pass", "detail": "" },
    { "name": "PostUninstallDetect", "status": "Pass", "detail": "" },
    { "name": "Residue", "status": "Warn", "detail": "1 leftover registry key" }
  ]
}
```

- **`recipePath`:** the recipe's path in the repository, with forward slashes.
- **`recipeSha256`:** the recipe hash (see below) at the time the live test passed.
- **`verdict`:**
  - `Validated` or `ValidatedWithWarnings` when the test passed;
  - `NotValidated` otherwise.
- **`validated`:** true only for the two passing verdicts. If you publish
  without a passing test, the Studio still writes the file, with `validated: false`.
  `reason` then says why.
- **`phases`:** copied from the harness's `summary.json`
  ([Invoke-RecipeInstallTest.ps1](../Tests/RecipeHarness/README.md#output)).
- **`installer` and `installExperience`:** for reviewers. The check ignores them.

### The recipe hash

1. Read the recipe file as bytes.
2. Drop a leading UTF-8 byte order mark (`EF BB BF`).
3. Replace every CRLF with LF, leaving a lone CR alone.
4. Take the SHA-256 and write it as lowercase hex.

Line endings don't change the hash, so git's `core.autocrlf` conversion makes no
difference. `id: 7zip` followed by LF or by CRLF both hash to
`dac9a5b6e5a249a1a674bd890731d93ada96e9fcb4b480a3659416695f52486b`.
`Get-RecipeContentHash` in `Modules/YardstickRecipeEvidence.psm1` computes it.

## What the check enforces

The **Recipe evidence** workflow (`.github/workflows/recipe-evidence.yml`) runs on
every pull request to `main`. It runs `Tests/RecipeHarness/Test-RecipeEvidence.ps1`.

**Which recipes it checks:** every `.yaml` or `.yml` file under `Recipes/` that the
pull request adds, modifies or renames into place, compared with the merge base. It
skips:
- `Recipes/Development/`;
- any `Disabled` folder;
- `*.common.yaml`;
- deleted recipes.

**Each such recipe fails unless all of these hold:**
1. `Tests/RecipeEvidence/<id>.json` exists in the pull request and is valid JSON.
2. `schemaVersion` is `1`.
3. `recipeId` is the file name without its extension, and `recipePath` is the
   recipe's path. Both are case-sensitive.
4. `recipeSha256` is the hash of the recipe as committed in the pull request.
5. `validated` is `true`, and `verdict` is `Validated` or `ValidatedWithWarnings`.
6. All eight phases are listed, and every one is `Pass` or `Warn`:
   - Schema, Download, PreDetect, Install, PostInstallDetect, Uninstall,
     PostUninstallDetect and Residue;
   - names and statuses are not case-sensitive;
   - `PostDetect` and `PostUninstall` are accepted as the older names of the two
     detection phases.
7. `Test-RecipeSchema.ps1` passes. The icon check is skipped because icons are not
   in Git. A `base:` recipe is read as committed in the pull request too, and must
   be in it.

**Warnings only:**
- an evidence file whose recipe the pull request deletes;
- an evidence file that matches no recipe.

The job summary has one row per recipe: result, verdict, test time and the reasons
it failed.

## Fixing a failing check

| Message | Fix |
|---|---|
| No evidence file | Test the recipe in Yardstick Development Studio and publish it from there. |
| Hash mismatch | The recipe changed after it was tested. Re-run the live test in the Studio and publish again. A line-ending change alone never causes this. |
| validated is not true / verdict / Phase … | The live test did not pass. Fix the recipe until it does, then publish again. |
| recipePath | The recipe moved since it was tested. Publish it again from its new path. |
| Schema: … | Fix the recipe. Run `.\Tests\RecipeHarness\Test-RecipeSchema.ps1 -RecipePath <file>` to see the errors. |

You can run the check before you push. It reads the committed recipe and evidence,
not your working copy, so commit first:

```powershell
git fetch origin
.\Tests\RecipeHarness\Test-RecipeEvidence.ps1 -BaseRef origin/main
```

## Waiving the evidence: `evidence-waived`

Some recipes can't be live-tested on a build machine, such as licensed software or
`manualDownload` media. For those, someone with write access adds the
**`evidence-waived`** label to the pull request.
- Evidence failures then become warnings, and the check passes.
- Schema failures still fail it.

Only people with write access can add labels, so the label is the control. Say why
in the pull request.

## Evidence in production: `requireRecipeEvidence`

When Yardstick [syncs recipes from GitHub](Installation.md#recipe-repositories), set
`requireRecipeEvidence: true` to apply the same rules on the way in. A new or
changed recipe is only taken when its evidence file in the same repository passes
rules 1 to 6 above.
- The recipe is not schema-checked here; the run checks it when it loads it.
- Otherwise the copy already here stays, and the run's log says why.
- A recipe merged with `evidence-waived` has no passing evidence, so it is not
  taken. Copy it into `Recipes` by hand if you want it.

## What this does and doesn't prove

- **The evidence is self-reported by the Studio.** CI can't re-run a live install
  test, and the files are not signed.
- **The check proves** the pull request carries Studio evidence for exactly these
  recipe contents. It is a process gate, not a security boundary.
- **Reviewers and CODEOWNERS** remain the control.

**Not covered:**
- **Inherited recipes:** a recipe with `base:` inherits fields from another recipe
  file. The hash covers only the recipe's own file. Changing a base recipe needs
  evidence for the base, but not for the recipes that inherit from it.
- **Site-specific overrides:** `Local/recipeOverrides.yaml` is not in Git, so it
  isn't covered either.
