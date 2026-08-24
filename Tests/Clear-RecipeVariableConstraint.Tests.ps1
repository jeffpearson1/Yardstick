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
}
