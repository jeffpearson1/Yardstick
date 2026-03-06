BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force
}

Describe "Merge-RecipeWithBase" {
    BeforeAll {
        # Create a temporary recipes directory with a base recipe file
        $recipesPath = Join-Path $TestDrive "Recipes"
        New-Item -Path $recipesPath -ItemType Directory -Force | Out-Null

        $baseRecipeContent = @"
id: testapp
displayName: Test App
description: The base description
publisher: Test Publisher
installScript: base-install.ps1
uninstallScript: base-uninstall.ps1
detectionType: script
detectionScript: base-detection.ps1
iconFile: testapp.png
"@
        Set-Content -Path (Join-Path $recipesPath "testapp.yaml") -Value $baseRecipeContent

        # Create a base recipe that itself has a 'base' field (for chaining test)
        $chainedBaseContent = @"
base: someother
id: chainedbase
displayName: Chained Base
"@
        Set-Content -Path (Join-Path $recipesPath "chainedbase.yaml") -Value $chainedBaseContent
    }

    Context "Recipes without base field" {
        It "Returns the recipe unchanged when no base field is present" {
            $recipe = @{
                id          = "standalone"
                displayName = "Standalone App"
            }
            $result = Merge-RecipeWithBase -Recipe $recipe -RecipesPath $recipesPath
            $result['id'] | Should -Be "standalone"
            $result['displayName'] | Should -Be "Standalone App"
            $result.Keys.Count | Should -Be 2
        }
    }

    Context "Successful merge" {
        It "Inherits fields from the base recipe" {
            $child = @{
                base = "testapp"
                id   = "testapp_sdl"
                displayName = "Test App SDL"
            }
            $result = Merge-RecipeWithBase -Recipe $child -RecipesPath $recipesPath
            $result['publisher'] | Should -Be "Test Publisher"
            $result['installScript'] | Should -Be "base-install.ps1"
            $result['detectionType'] | Should -Be "script"
            $result['detectionScript'] | Should -Be "base-detection.ps1"
            $result['iconFile'] | Should -Be "testapp.png"
        }

        It "Child fields override base fields" {
            $child = @{
                base        = "testapp"
                id          = "testapp_sdl"
                displayName = "Test App SDL"
                description = "SDL override description"
            }
            $result = Merge-RecipeWithBase -Recipe $child -RecipesPath $recipesPath
            $result['id'] | Should -Be "testapp_sdl"
            $result['displayName'] | Should -Be "Test App SDL"
            $result['description'] | Should -Be "SDL override description"
        }

        It "Does not include the base key in merged result" {
            $child = @{
                base        = "testapp"
                id          = "testapp_sdl"
                displayName = "Test App SDL"
            }
            $result = Merge-RecipeWithBase -Recipe $child -RecipesPath $recipesPath
            $result.ContainsKey('base') | Should -BeFalse
        }

        It "Includes all base fields plus child overrides" {
            $child = @{
                base        = "testapp"
                id          = "testapp_sdl"
                displayName = "Test App SDL"
                scopeTags   = @("Lab")
            }
            $result = Merge-RecipeWithBase -Recipe $child -RecipesPath $recipesPath
            # Base has 8 fields; child adds scopeTags and overrides id+displayName; base key stripped
            $result.ContainsKey('scopeTags') | Should -BeTrue
            $result['scopeTags'] | Should -Be @("Lab")
            $result.ContainsKey('publisher') | Should -BeTrue
            $result.ContainsKey('uninstallScript') | Should -BeTrue
        }
    }

    Context "Error handling" {
        It "Throws when base recipe is not found" {
            $child = @{
                base = "nonexistent"
                id   = "child_app"
            }
            { Merge-RecipeWithBase -Recipe $child -RecipesPath $recipesPath } |
                Should -Throw "Base recipe 'nonexistent' not found."
        }

        It "Throws when base recipe also has a base field (chained inheritance)" {
            $child = @{
                base = "chainedbase"
                id   = "child_app"
            }
            { Merge-RecipeWithBase -Recipe $child -RecipesPath $recipesPath } |
                Should -Throw "Chained inheritance is not supported*"
        }
    }

    Context "Subdirectory support" {
        BeforeAll {
            $subDir = Join-Path $recipesPath "Interactive"
            New-Item -Path $subDir -ItemType Directory -Force | Out-Null

            $subRecipeContent = @"
id: subapp
displayName: Sub App
publisher: Sub Publisher
detectionType: msi
installScript: sub-install.ps1
uninstallScript: sub-uninstall.ps1
"@
            Set-Content -Path (Join-Path $subDir "subapp.yaml") -Value $subRecipeContent
        }

        It "Finds base recipe in subdirectory" {
            $child = @{
                base        = "subapp"
                id          = "subapp_sdl"
                displayName = "Sub App SDL"
            }
            $result = Merge-RecipeWithBase -Recipe $child -RecipesPath $recipesPath
            $result['publisher'] | Should -Be "Sub Publisher"
            $result['id'] | Should -Be "subapp_sdl"
        }
    }
}
