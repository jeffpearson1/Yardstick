BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    Import-Module powershell-yaml -ErrorAction Stop
    Import-Module (Join-Path $script:repoRoot 'Modules\YardstickSupport.psm1') -Force
    Import-Module (Join-Path $script:repoRoot 'Modules\YardstickConfig.psm1') -Force
}

Describe "Import-RecipeOverrides" {
    BeforeEach {
        $script:path = Join-Path $TestDrive 'recipeOverrides.yaml'
        Remove-Item -LiteralPath $script:path -ErrorAction SilentlyContinue
    }

    It "Returns no overrides when the file does not exist" {
        $result = Import-RecipeOverrides -Path $script:path
        $result.Count | Should -Be 0
    }

    It "Returns no overrides for an empty file" {
        Set-Content -LiteralPath $script:path -Value ''
        (Import-RecipeOverrides -Path $script:path).Count | Should -Be 0
    }

    It "Returns no overrides for a file with only comments" {
        Set-Content -LiteralPath $script:path -Value "# Nothing here yet"
        (Import-RecipeOverrides -Path $script:path).Count | Should -Be 0
    }

    It "Reads each recipe's fields" {
        Set-Content -LiteralPath $script:path -Value @"
7zip:
  scopeTags:
    - Workstations
    - Labs
vlc:
  owner: Media team
"@
        $result = Import-RecipeOverrides -Path $script:path
        $result.Count | Should -Be 2
        $result['7zip']['scopeTags'] | Should -Be @('Workstations', 'Labs')
        $result['vlc']['owner'] | Should -Be 'Media team'
    }

    It "Matches recipe ids case-insensitively" {
        Set-Content -LiteralPath $script:path -Value "GoogleChrome:`n  scopeTags: [Workstations]"
        $result = Import-RecipeOverrides -Path $script:path
        $result.ContainsKey('googlechrome') | Should -BeTrue
    }

    It "Throws when the file is not a map of recipe ids" {
        Set-Content -LiteralPath $script:path -Value "- 7zip`n- vlc"
        { Import-RecipeOverrides -Path $script:path } | Should -Throw '*map recipe ids*'
    }

    It "Throws when a recipe's entry is not a map of fields" {
        Set-Content -LiteralPath $script:path -Value "7zip: Workstations"
        { Import-RecipeOverrides -Path $script:path } | Should -Throw "*'7zip'*"
    }

    It "Throws when an entry overrides <_>" -ForEach 'id', 'base' {
        Set-Content -LiteralPath $script:path -Value "7zip:`n  $($_): other"
        { Import-RecipeOverrides -Path $script:path } | Should -Throw "*cannot override '$_'*"
    }

    It "Throws on YAML that does not parse" {
        Set-Content -LiteralPath $script:path -Value "7zip:`n  scopeTags: [Workstations"
        { Import-RecipeOverrides -Path $script:path } | Should -Throw
    }
}

Describe "Merge-RecipeWithOverride" {
    It "Replaces overridden fields and keeps the rest" {
        $recipe = @{ id = '7zip'; displayName = '7-Zip'; scopeTags = @('Shared') }
        $result = Merge-RecipeWithOverride -Recipe $recipe -Override @{ scopeTags = @('Workstations', 'Labs') }
        $result.scopeTags | Should -Be @('Workstations', 'Labs')
        $result.displayName | Should -Be '7-Zip'
        $result.id | Should -Be '7zip'
    }

    It "Adds a field the recipe does not have" {
        $result = Merge-RecipeWithOverride -Recipe @{ id = '7zip' } -Override @{ scopeTags = @('Workstations') }
        $result.scopeTags | Should -Be @('Workstations')
    }

    It "Replaces a list rather than adding to it" {
        $recipe = @{ id = '7zip'; scopeTags = @('Shared', 'Labs') }
        $result = Merge-RecipeWithOverride -Recipe $recipe -Override @{ scopeTags = @('Workstations') }
        @($result.scopeTags) | Should -Be @('Workstations')
    }

    It "Does not change the recipe it was given" {
        $recipe = @{ id = '7zip'; scopeTags = @('Shared') }
        Merge-RecipeWithOverride -Recipe $recipe -Override @{ scopeTags = @('Workstations') } | Out-Null
        $recipe.scopeTags | Should -Be @('Shared')
    }
}

Describe "Get-YardstickRecipeOverridesPath" {
    It "Uses the RecipeOverrides entry of the project manifest" {
        $root = Join-Path $TestDrive 'WithEntry'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $root 'Yardstick.Project.psd1') -Value "@{ RecipeOverrides = 'Site\overrides.yaml' }"
        Get-YardstickRecipeOverridesPath -ProjectRoot $root | Should -Be (Join-Path $root 'Site\overrides.yaml')
    }

    It "Falls back to Local\recipeOverrides.yaml for a manifest without the entry" {
        $root = Join-Path $TestDrive 'WithoutEntry'
        New-Item -ItemType Directory -Path $root -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $root 'Yardstick.Project.psd1') -Value "@{ RecipesPath = 'Recipes' }"
        Get-YardstickRecipeOverridesPath -ProjectRoot $root | Should -Be (Join-Path $root 'Local\recipeOverrides.yaml')
    }

    It "Points at Local\recipeOverrides.yaml in the shipped manifest" {
        Get-YardstickRecipeOverridesPath -ProjectRoot $script:repoRoot | Should -Be (Join-Path $script:repoRoot 'Local\recipeOverrides.yaml')
    }
}
