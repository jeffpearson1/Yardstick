BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"

    # Clear-RecipeVariableConstraint operates on script scope, and Pester gives each
    # It block its own scope, so calling it directly from a test would not reproduce
    # how it behaves in the runner. Instead each scenario runs inside a real .ps1
    # (genuine script scope) and reports back what happened.
    $Global:Fixture = Join-Path $TestDrive 'constraint-scenario.ps1'
    Set-Content -Path $Global:Fixture -Value @'
param([Parameter(Mandatory)][string]$Scenario, [string]$YardstickPath)

function Write-Log { param([string]$Message) }

# Clear-RecipeVariableConstraint lives in Yardstick.ps1, which cannot be dot-sourced
# (its top level runs a whole update pass), so extract just the function via the AST.
$ast = [System.Management.Automation.Language.Parser]::ParseFile($YardstickPath, [ref]$null, [ref]$null)
$fn = $ast.Find({ param($n)
    $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $n.Name -eq 'Clear-RecipeVariableConstraint' }, $true)
if (-not $fn) { throw "Clear-RecipeVariableConstraint not found in $YardstickPath" }
. ([scriptblock]::Create($fn.Extent.Text))

# A recipe using the left-side cast that caused the original outage, and one that
# reuses the same variable name for an incompatible value.
$poisoning = { [xml]$manifest = '<AppInstaller><Version>1.0</Version></AppInstaller>' }
$victim    = { $manifest = @{ version = '1.11.7' } }

switch ($Scenario) {
    'NoReset' {
        Invoke-Command -ScriptBlock $poisoning -NoNewScope
        try { Invoke-Command -ScriptBlock $victim -NoNewScope; 'no-error' }
        catch { $_.Exception.Message }
    }
    'WithReset' {
        Clear-RecipeVariableConstraint
        Invoke-Command -ScriptBlock $poisoning -NoNewScope
        Clear-RecipeVariableConstraint
        try { Invoke-Command -ScriptBlock $victim -NoNewScope; $manifest.GetType().Name }
        catch { "threw: $($_.Exception.Message)" }
    }
    'RepeatedIterations' {
        Clear-RecipeVariableConstraint
        for ($i = 0; $i -lt 3; $i++) {
            Clear-RecipeVariableConstraint
            try { Invoke-Command -ScriptBlock $poisoning -NoNewScope } catch { return "threw on poisoning pass $i" }
            Clear-RecipeVariableConstraint
            try { Invoke-Command -ScriptBlock $victim -NoNewScope } catch { return "threw on victim pass $i" }
        }
        'all-passed'
    }
    'PreferenceVariables' {
        # PowerShell's preference variables are type-constrained by design. When
        # Yardstick.ps1 runs as a top-level script they live in ITS script scope,
        # and they are materialised lazily -- so they can be absent when the
        # baseline is taken and present afterwards. That combination (constrained,
        # in script scope, not in the baseline) is what made an earlier version of
        # this function delete them and change the runner's behaviour mid-run.
        # Recreate that exact condition here.
        Clear-RecipeVariableConstraint
        [System.Management.Automation.ActionPreference]$ErrorActionPreference = 'Continue'
        [System.Management.Automation.ActionPreference]$ProgressPreference = 'SilentlyContinue'
        [System.Management.Automation.ActionPreference]$WarningPreference = 'Continue'
        [System.Text.Encoding]$OutputEncoding = [System.Text.Encoding]::UTF8
        [System.Management.Automation.DefaultParameterDictionary]$PSDefaultParameterValues = @{}

        Invoke-Command -ScriptBlock $poisoning -NoNewScope
        Clear-RecipeVariableConstraint

        $missing = foreach ($n in 'ErrorActionPreference', 'ProgressPreference',
            'WarningPreference', 'OutputEncoding', 'PSDefaultParameterValues') {
            if (-not (Get-Variable $n -Scope Script -ErrorAction SilentlyContinue)) { $n }
        }
        if ($missing) { "deleted: $($missing -join ',')" }
        elseif ($ErrorActionPreference -ne 'Continue') { "value changed to $ErrorActionPreference" }
        else { 'all-preserved' }
    }
    'FirstCallRemovesNothing' {
        Invoke-Command -ScriptBlock $poisoning -NoNewScope
        Clear-RecipeVariableConstraint      # this is the baseline call
        if ($null -ne $manifest) { 'preserved' } else { 'removed' }
    }
    'UntypedStateUntouched' {
        Clear-RecipeVariableConstraint
        $plainValue = 'keep me'
        $accumulator = [System.Collections.Generic.List[string]]::new()
        $accumulator.Add('first')
        Invoke-Command -ScriptBlock $poisoning -NoNewScope
        Clear-RecipeVariableConstraint
        $accumulator.Add('second')
        "$plainValue|$($accumulator -join ',')"
    }
}
'@
    $Global:YardstickPath = (Resolve-Path "$PSScriptRoot\..\Yardstick.ps1").Path

    function Invoke-Scenario {
        param([string]$Scenario)
        & $Global:Fixture -Scenario $Scenario -YardstickPath $Global:YardstickPath
    }

    # Restore-RunnerPathVariable has the same problem: it reads and writes script
    # scope, so it only behaves realistically inside a real .ps1.
    $Global:PathFixture = Join-Path $TestDrive 'runner-path-scenario.ps1'
    Set-Content -Path $Global:PathFixture -Value @'
param([Parameter(Mandatory)][string]$Scenario, [string]$YardstickPath)

function Write-Log { param([string]$Message) }

$ast = [System.Management.Automation.Language.Parser]::ParseFile($YardstickPath, [ref]$null, [ref]$null)
$fn = $ast.Find({ param($n)
    $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
    $n.Name -eq 'Restore-RunnerPathVariable' }, $true)
if (-not $fn) { throw "Restore-RunnerPathVariable not found in $YardstickPath" }
. ([scriptblock]::Create($fn.Extent.Text))

# The runner sets these once at startup from preferences.
$Script:Temp       = 'C:\Yardstick\Temp'
$Script:BuildSpace = 'C:\Yardstick\Build'
$Script:Published  = 'C:\Yardstick\Published'

# hxd's actual defect: an innocuous lowercase $temp in a preDownloadScript.
$clobberingRecipe = { $temp = 'C:\Users\someone\AppData\Local\Temp\hxd-abc123' }

switch ($Scenario) {
    'ReproducesTheClobber' {
        Invoke-Command -ScriptBlock $clobberingRecipe -NoNewScope
        $Script:Temp
    }
    'RestoresAfterClobber' {
        Restore-RunnerPathVariable          # baseline
        Invoke-Command -ScriptBlock $clobberingRecipe -NoNewScope
        Restore-RunnerPathVariable
        $Script:Temp
    }
    'FirstCallChangesNothing' {
        Invoke-Command -ScriptBlock $clobberingRecipe -NoNewScope
        Restore-RunnerPathVariable          # baseline is taken from the clobbered value
        $Script:Temp
    }
    'RestoresEveryTrackedName' {
        Restore-RunnerPathVariable
        Invoke-Command -ScriptBlock {
            $temp = 'x'; $buildspace = 'y'; $PUBLISHED = 'z'
        } -NoNewScope
        Restore-RunnerPathVariable
        "$Script:Temp|$Script:BuildSpace|$Script:Published"
    }
    'StableAcrossIterations' {
        Restore-RunnerPathVariable
        for ($i = 0; $i -lt 3; $i++) {
            Restore-RunnerPathVariable
            Invoke-Command -ScriptBlock $clobberingRecipe -NoNewScope
        }
        Restore-RunnerPathVariable
        $Script:Temp
    }
    'LeavesRecipeStateAlone' {
        Restore-RunnerPathVariable
        $someRecipeVariable = 'keep me'
        Invoke-Command -ScriptBlock $clobberingRecipe -NoNewScope
        Restore-RunnerPathVariable
        $someRecipeVariable
    }
}
'@

    function Invoke-PathScenario {
        param([string]$Scenario)
        & $Global:PathFixture -Scenario $Scenario -YardstickPath $Global:YardstickPath
    }
}

Describe "Clear-RecipeVariableConstraint" {
    It "Reproduces the leak when no reset happens between recipes" {
        Invoke-Scenario 'NoReset' | Should -BeLike "*XmlDocument*"
    }

    It "Lets a later recipe reuse the variable name after a reset" {
        Invoke-Scenario 'WithReset' | Should -Be 'Hashtable'
    }

    It "Stays stable across repeated loop iterations" {
        Invoke-Scenario 'RepeatedIterations' | Should -Be 'all-passed'
    }

    It "Preserves PowerShell preference variables" {
        Invoke-Scenario 'PreferenceVariables' | Should -Be 'all-preserved'
    }

    It "Removes nothing on its first call" {
        Invoke-Scenario 'FirstCallRemovesNothing' | Should -Be 'preserved'
    }

    It "Leaves untyped recipe state alone" {
        Invoke-Scenario 'UntypedStateUntouched' | Should -Be 'keep me|first,second'
    }
}

Describe "Restore-RunnerPathVariable" {
    It "Reproduces the clobber it exists to repair" {
        Invoke-PathScenario 'ReproducesTheClobber' | Should -BeLike '*hxd-abc123'
    }

    It "Puts the runner's temp directory back after a recipe overwrites it" {
        Invoke-PathScenario 'RestoresAfterClobber' | Should -Be 'C:\Yardstick\Temp'
    }

    It "Takes a baseline on its first call and changes nothing" {
        Invoke-PathScenario 'FirstCallChangesNothing' | Should -BeLike '*hxd-abc123'
    }

    It "Restores every tracked folder variable regardless of casing" {
        Invoke-PathScenario 'RestoresEveryTrackedName' |
            Should -Be 'C:\Yardstick\Temp|C:\Yardstick\Build|C:\Yardstick\Published'
    }

    It "Stays stable across repeated loop iterations" {
        Invoke-PathScenario 'StableAcrossIterations' | Should -Be 'C:\Yardstick\Temp'
    }

    It "Leaves unrelated recipe state alone" {
        Invoke-PathScenario 'LeavesRecipeStateAlone' | Should -Be 'keep me'
    }
}

Describe "Recipe scripts run by the Yardstick host" {
    # Left-side casts leak a type constraint into the shared recipe scope. Casts in
    # detectionScript/installScript are fine -- those run on the endpoint, not here.
    It "Contain no left-side type casts" {
        $hostKeys = 'preDownloadScript', 'downloadScript', 'postDownloadScript', 'postRunScript'
        $offenders = foreach ($file in Get-ChildItem "$PSScriptRoot\..\Recipes" -Recurse -Filter *.yaml) {
            $recipe = try { ConvertFrom-Yaml (Get-Content $file.FullName -Raw) } catch { continue }
            if ($recipe -isnot [hashtable]) { continue }
            foreach ($key in $hostKeys) {
                if ($recipe.$key -match '(?m)^\s*\[[A-Za-z.]+\]\s*\$[A-Za-z_]\w*\s*=') {
                    "$($file.Name) ($key)"
                }
            }
        }
        $offenders | Should -BeNullOrEmpty
    }

    # Host-side scripts share the runner's script scope, so assigning one of its
    # folder variables redirects Yardstick's own paths. hxd did this with $temp and
    # then deleted the directory in a finally block, breaking every later recipe.
    It "Assign to no reserved runner variable" {
        $hostKeys = 'preDownloadScript', 'downloadScript', 'postDownloadScript', 'postRunScript'
        $pattern = '(?im)^\s*\$(temp|buildspace|scripts|published|recipes|icons|tools|secrets|softwaredropbox|softwarearchive|prefs|applications)\s*='
        $offenders = foreach ($file in Get-ChildItem "$PSScriptRoot\..\Recipes" -Recurse -Filter *.yaml) {
            $recipe = try { ConvertFrom-Yaml (Get-Content $file.FullName -Raw) } catch { continue }
            if ($recipe -isnot [hashtable]) { continue }
            foreach ($key in $hostKeys) {
                foreach ($match in [regex]::Matches([string]$recipe.$key, $pattern)) {
                    "$($file.Name) ($key): `$$($match.Groups[1].Value)"
                }
            }
        }
        $offenders | Should -BeNullOrEmpty
    }
}
