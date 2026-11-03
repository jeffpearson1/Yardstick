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

    It "Reads a variant entry" {
        Set-Content -LiteralPath $script:path -Value "zoom.labs:`n  variantOf: zoom`n  displayName: Zoom (Labs)"
        (Import-RecipeOverrides -Path $script:path)['zoom.labs']['variantOf'] | Should -Be 'zoom'
    }

    It "Throws when variantOf is <Case>" -ForEach @(
        @{ Case = 'empty'; Value = "''"; Message = '*name one recipe id*' }
        @{ Case = 'a list'; Value = '[zoom, vlc]'; Message = '*name one recipe id*' }
        @{ Case = 'the entry itself'; Value = 'ZOOM.labs'; Message = '*variant of itself*' }
    ) {
        Set-Content -LiteralPath $script:path -Value "zoom.labs:`n  variantOf: $Value"
        { Import-RecipeOverrides -Path $script:path } | Should -Throw $Message
    }
}

Describe "Recipe variants" {
    BeforeAll {
        $script:recipe = @{ id = 'Zoom'; displayName = 'Zoom'; installScript = 'msiexec /i zoom.msi'; owner = 'Recipe owner' }
        $script:overrides = [hashtable]::new([System.StringComparer]::OrdinalIgnoreCase)
        $script:overrides['zoom'] = @{ scopeTags = @('Staff'); installScript = 'msiexec /i zoom.msi SSO=1' }
        $script:overrides['zoom.labs'] = @{ variantOf = 'zoom'; displayName = 'Zoom (Labs)'; installScript = 'msiexec /i zoom.msi LABS=1' }
        $script:overrides['vlc'] = @{ owner = 'Media team' }
    }

    It "Finds the recipe a variant is built from, matching ids case-insensitively" {
        Get-RecipeVariantParent -Id 'ZOOM.LABS' -Overrides $script:overrides | Should -Be 'zoom'
        Get-RecipeVariantParent -Id 'zoom' -Overrides $script:overrides | Should -BeNullOrEmpty
        Get-RecipeVariantParent -Id 'notthere' -Overrides $script:overrides | Should -BeNullOrEmpty
    }

    It "Lists only the variants" {
        $variants = @(Get-RecipeVariant -Overrides $script:overrides)
        $variants.Count | Should -Be 1
        $variants[0].Id | Should -Be 'zoom.labs'
        $variants[0].VariantOf | Should -Be 'zoom'
    }

    It "Applies only a recipe's own entry to the recipe" {
        $result = Resolve-RecipeOverride -Recipe $script:recipe -Id 'zoom' -Overrides $script:overrides
        $result.installScript | Should -Be 'msiexec /i zoom.msi SSO=1'
        $result.displayName | Should -Be 'Zoom'
        $result.id | Should -Be 'Zoom'
    }

    It "Leaves a recipe with no entry as it is" {
        $result = Resolve-RecipeOverride -Recipe $script:recipe -Id 'other' -Overrides $script:overrides
        $result.installScript | Should -Be 'msiexec /i zoom.msi'
    }

    It "Layers the variant's entry over the recipe's and takes the variant's id" {
        $result = Resolve-RecipeOverride -Recipe $script:recipe -Id 'zoom.labs' -Overrides $script:overrides
        $result.id | Should -Be 'zoom.labs'
        $result.displayName | Should -Be 'Zoom (Labs)'
        $result.installScript | Should -Be 'msiexec /i zoom.msi LABS=1'
        $result.scopeTags | Should -Be @('Staff')
        $result.owner | Should -Be 'Recipe owner'
        $result.ContainsKey('variantOf') | Should -BeFalse
    }

    It "Does not change the recipe it was given" {
        Resolve-RecipeOverride -Recipe $script:recipe -Id 'zoom.labs' -Overrides $script:overrides | Out-Null
        $script:recipe.id | Should -Be 'Zoom'
        $script:recipe.displayName | Should -Be 'Zoom'
    }

    It "Gives a variant its own dropbox folder unless its entry names one" {
        $recipe = @{ id = 'stata'; displayName = 'Stata'; manualDownload = $true; manualDownloadFolder = 'Stata' }
        $overrides = @{
            'stata.labs'  = @{ variantOf = 'stata'; displayName = 'Stata (Labs)' }
            'stata.kiosk' = @{ variantOf = 'stata'; displayName = 'Stata (Kiosk)'; manualDownloadFolder = 'StataKiosk' }
        }
        (Resolve-RecipeOverride -Recipe $recipe -Id 'stata.labs' -Overrides $overrides).ContainsKey('manualDownloadFolder') | Should -BeFalse
        (Resolve-RecipeOverride -Recipe $recipe -Id 'stata.kiosk' -Overrides $overrides).manualDownloadFolder | Should -Be 'StataKiosk'
        $recipe.manualDownloadFolder | Should -Be 'Stata'
    }

    It "Throws when a variant keeps <Case>" -ForEach @(
        @{ Case = "its recipe's displayName"; Fields = @{ variantOf = 'zoom' } }
        @{ Case = "the same displayName"; Fields = @{ variantOf = 'zoom'; displayName = 'zoom' } }
    ) {
        $overrides = @{ 'zoom.kiosk' = $Fields }
        { Resolve-RecipeOverride -Recipe $script:recipe -Id 'zoom.kiosk' -Overrides $overrides } | Should -Throw '*needs a displayName of its own*'
    }

    It "Throws for a variant of a variant, naming the recipe to use" {
        $overrides = @{
            'zoom.labs'  = @{ variantOf = 'zoom'; displayName = 'Zoom (Labs)' }
            'zoom.kiosk' = @{ variantOf = 'zoom.labs'; displayName = 'Zoom (Kiosk)' }
        }
        { Resolve-RecipeOverride -Recipe $script:recipe -Id 'zoom.kiosk' -Overrides $overrides } | Should -Throw "*Set its variantOf to 'zoom'*"
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

Describe "Get-RecipePlaceholder" {
    It "Names each field that still holds a placeholder" {
        $recipe = @{
            id              = 'sccmclient'
            installScript   = 'ccmsetup.exe SMSSITECODE=<YOUR_SITE_CODE> /mp:<YOUR_MANAGEMENT_POINT>'
            url             = '\\<YOUR_SITE_SERVER>\Client'
            displayName     = 'Configuration Manager Client'
        }
        $found = @(Get-RecipePlaceholder -Recipe $recipe)
        $found.Field | Should -Be @('installScript', 'url')
        $found[0].Placeholder | Should -Be '<YOUR_SITE_CODE>'
        $found[1].Placeholder | Should -Be '<YOUR_SITE_SERVER>'
    }

    It "Looks inside lists" {
        $found = @(Get-RecipePlaceholder -Recipe @{ urlRedirects = @('https://example.com', 'https://<YOUR_ZOOM_VANITY_URL>/') })
        $found.Field | Should -Be 'urlRedirects'
    }

    It "Returns nothing once the overrides fill the fields" {
        $recipe = @{ id = 'zoom'; url = 'https://<YOUR_ZOOM_VANITY_URL>/download' }
        $merged = Merge-RecipeWithOverride -Recipe $recipe -Override @{ url = 'https://example.zoom.us/download' }
        Get-RecipePlaceholder -Recipe $merged | Should -BeNullOrEmpty
    }

    It "Ignores look-alikes that are not placeholders" {
        $recipe = @{ installScript = 'setup.exe /log <your_log> & echo <YOUR-NAME>'; uninstallScript = '<VERSION>' }
        Get-RecipePlaceholder -Recipe $recipe | Should -BeNullOrEmpty
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
