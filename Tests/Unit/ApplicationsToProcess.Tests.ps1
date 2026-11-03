BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    Import-Module powershell-yaml -ErrorAction Stop
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickSupport.psm1" -Force

    # Get-ApplicationsToProcess lives in Invoke-YardstickRun.ps1, which cannot be
    # dot-sourced (its top level runs a whole update pass), so extract it via the
    # AST. It reads the runner's $Recipes and $RecipeOverrides, set below.
    $runnerPath = "$PSScriptRoot\..\..\Private\Invoke-YardstickRun.ps1"
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($runnerPath, [ref]$null, [ref]$null)
    $fn = $ast.Find({ param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-ApplicationsToProcess' }, $true)
    if (-not $fn) { throw "Get-ApplicationsToProcess not found in $runnerPath" }
    . ([scriptblock]::Create($fn.Extent.Text))

    $script:Recipes = Join-Path $TestDrive 'Recipes'
    New-Item -ItemType Directory -Path (Join-Path $script:Recipes 'Disabled') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $script:Recipes 'Interactive') -Force | Out-Null
    foreach ($id in 'alpha', 'zoom', 'Disabled\old', 'Interactive\setup') {
        Set-Content -LiteralPath (Join-Path $script:Recipes "$id.yaml") -Value "id: $(Split-Path $id -Leaf)"
    }

    function Set-Overrides([hashtable]$Entries) {
        $script:RecipeOverrides = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($key in $Entries.Keys) { $script:RecipeOverrides[$key] = $Entries[$key] }
    }
}

Describe "Get-ApplicationsToProcess -All with recipe variants" {
    It "Lists no variants when the overrides define none" {
        Set-Overrides @{ zoom = @{ scopeTags = @('Staff') } }
        @(Get-ApplicationsToProcess -All) | Should -Be @('alpha', 'setup', 'zoom')
    }

    It "Lists each variant straight after the recipe it is built from" {
        Set-Overrides @{
            'zoom.labs'  = @{ variantOf = 'zoom'; displayName = 'Zoom (Labs)' }
            'alpha.kiosk' = @{ variantOf = 'ALPHA'; displayName = 'Alpha (Kiosk)' }
        }
        @(Get-ApplicationsToProcess -All) | Should -Be @('alpha', 'alpha.kiosk', 'setup', 'zoom', 'zoom.labs')
    }

    It "Leaves a variant out with its disabled or interactive recipe" {
        Set-Overrides @{
            'old.labs'   = @{ variantOf = 'old'; displayName = 'Old (Labs)' }
            'setup.labs' = @{ variantOf = 'setup'; displayName = 'Setup (Labs)' }
        }
        @(Get-ApplicationsToProcess -All -NoInteractive) | Should -Be @('alpha', 'zoom')
    }

    It "Lists a variant whose recipe is missing, so the run reports it" {
        Set-Overrides @{ 'gone.labs' = @{ variantOf = 'gone'; displayName = 'Gone (Labs)' } }
        @(Get-ApplicationsToProcess -All) | Should -Be @('alpha', 'setup', 'zoom', 'gone.labs')
    }

    It "Lists a variant that shares a recipe's id only once, as that recipe" {
        Set-Overrides @{ alpha = @{ variantOf = 'zoom'; displayName = 'Zoom (Alpha)' } }
        @(Get-ApplicationsToProcess -All) | Should -Be @('alpha', 'setup', 'zoom')
    }
}
