BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force

    # Helper: builds a minimal valid recipe hashtable
    function New-ValidRecipe {
        param(
            [string]$DetectionType = "msi"
        )
        $recipe = @{
            id               = "testapp"
            displayName      = "Test App"
            detectionType    = $DetectionType
            iconFile         = "testapp.png"
            description      = "A test application"
            publisher        = "Test Publisher"
            installScript    = "msiexec /i installer.msi /qn"
            uninstallScript  = "msiexec /x {GUID} /qn"
        }

        switch ($DetectionType) {
            "file" {
                $recipe["fileDetectionPath"] = "C:\Program Files\TestApp"
                $recipe["fileDetectionMethod"] = "version"
                $recipe["fileDetectionName"] = "testapp.exe"
            }
            "registry" {
                $recipe["registryDetectionMethod"] = "version"
                $recipe["registryDetectionKey"] = "HKLM\SOFTWARE\TestApp"
            }
            "script" {
                $recipe["detectionScript"] = "if (Test-Path 'C:\TestApp') { Write-Output 'Installed' }"
            }
        }

        return $recipe
    }
}

Describe "Test-RecipeSchema" {
    Context "Valid recipes" {
        It "Validates a minimal MSI recipe" {
            $recipe = New-ValidRecipe -DetectionType "msi"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeTrue
            $result.Errors.Count | Should -Be 0
        }

        It "Validates a file-detection recipe" {
            $recipe = New-ValidRecipe -DetectionType "file"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeTrue
        }

        It "Validates a registry-detection recipe" {
            $recipe = New-ValidRecipe -DetectionType "registry"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeTrue
        }

        It "Validates a script-detection recipe" {
            $recipe = New-ValidRecipe -DetectionType "script"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeTrue
        }

        It "Accepts powerShellInstallScript instead of installScript" {
            $recipe = New-ValidRecipe -DetectionType "msi"
            $recipe.Remove('installScript')
            $recipe['powerShellInstallScript'] = "Install-Something"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeTrue
        }
    }

    Context "Missing required fields" {
        It "Reports missing id" {
            $recipe = New-ValidRecipe
            $recipe.Remove('id')
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Missing required field 'id'"
        }

        It "Reports missing displayName" {
            $recipe = New-ValidRecipe
            $recipe.Remove('displayName')
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Missing required field 'displayName'"
        }

        It "Reports missing both install scripts" {
            $recipe = New-ValidRecipe
            $recipe.Remove('installScript')
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Missing install script: provide either 'installScript' or 'powerShellInstallScript'"
        }

        It "Reports missing both uninstall scripts" {
            $recipe = New-ValidRecipe
            $recipe.Remove('uninstallScript')
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Missing uninstall script: provide either 'uninstallScript' or 'powerShellUninstallScript'"
        }
    }

    Context "Detection type conditional fields" {
        It "Reports missing fileDetectionPath for file detection" {
            $recipe = New-ValidRecipe -DetectionType "file"
            $recipe.Remove('fileDetectionPath')
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Missing required field 'fileDetectionPath' for detectionType 'file'"
        }

        It "Reports missing registryDetectionKey for registry detection" {
            $recipe = New-ValidRecipe -DetectionType "registry"
            $recipe.Remove('registryDetectionKey')
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Missing required field 'registryDetectionKey' for detectionType 'registry'"
        }

        It "Reports missing detectionScript for script detection" {
            $recipe = New-ValidRecipe -DetectionType "script"
            $recipe.Remove('detectionScript')
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Should -Contain "Missing required field 'detectionScript' for detectionType 'script'"
        }
    }

    Context "Enumeration validation" {
        It "Reports invalid detectionType" {
            $recipe = New-ValidRecipe
            $recipe['detectionType'] = "invalid"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Where-Object { $_ -match "Invalid value.*detectionType" } | Should -Not -BeNullOrEmpty
        }

        It "Reports invalid installExperience" {
            $recipe = New-ValidRecipe
            $recipe['installExperience'] = "admin"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.IsValid | Should -BeFalse
            $result.Errors | Where-Object { $_ -match "Invalid value.*installExperience" } | Should -Not -BeNullOrEmpty
        }

        It "Accepts valid installExperience" {
            $recipe = New-ValidRecipe
            $recipe['installExperience'] = "system"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.Errors | Where-Object { $_ -match "installExperience" } | Should -BeNullOrEmpty
        }
    }

    Context "Unknown field warnings" {
        It "Warns on unknown field" {
            $recipe = New-ValidRecipe
            $recipe['unknownField'] = "something"
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.Warnings | Where-Object { $_ -match "Unknown field 'unknownField'" } | Should -Not -BeNullOrEmpty
        }

        It "Warns on incorrect casing" {
            $recipe = New-ValidRecipe
            $recipe['urlredirects'] = $true
            $result = Test-RecipeSchema -Recipe $recipe -RecipeId "testapp"
            $result.Warnings | Where-Object { $_ -match "incorrect casing.*urlRedirects" } | Should -Not -BeNullOrEmpty
        }
    }
}
