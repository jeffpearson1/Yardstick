BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force
}

Describe "Test-Prerequisites" {
    BeforeEach {
        # Recreated per test: TestDrive persists across It blocks, so the dummy
        # curl.exe one test drops would otherwise satisfy the next one.
        $script:toolsPath = Join-Path $TestDrive "Tools"
        if (Test-Path $script:toolsPath) {
            Remove-Item $script:toolsPath -Recurse -Force
        }
        New-Item -Path $script:toolsPath -ItemType Directory | Out-Null
    }

    It "Returns a result object with expected properties" {
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.PSObject.Properties.Name | Should -Contain "IsValid"
        $result.PSObject.Properties.Name | Should -Contain "Errors"
        $result.PSObject.Properties.Name | Should -Contain "Warnings"
    }

    It "Warns when curl.exe is not found" {
        # Windows ships curl.exe in System32, so the PATH lookup has to be mocked
        # away for this to test the bundled-copy check at all.
        Mock -ModuleName YardstickSupport Get-Command { $null } -ParameterFilter { $Name -eq 'curl.exe' }
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.Warnings | Where-Object { $_ -match "curl.exe" } | Should -Not -BeNullOrEmpty
    }

    It "Does not warn about curl.exe when it exists" {
        Mock -ModuleName YardstickSupport Get-Command { $null } -ParameterFilter { $Name -eq 'curl.exe' }
        # Create a dummy curl.exe
        New-Item -Path (Join-Path $script:toolsPath "curl.exe") -ItemType File -Force | Out-Null
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.Warnings | Where-Object { $_ -match "curl.exe" } | Should -BeNullOrEmpty
    }

    It "Does not warn when curl is only on PATH" {
        # Mocked rather than leaning on the host's own curl: the lookup resolves
        # through PATHEXT, and a runner whose shell truncates that (Git Bash sets
        # it to .CPL) would fail this test for reasons that have nothing to do
        # with the branch being exercised. No curl.exe under ToolsPath, so a
        # warning here means the PATH hit was ignored.
        Mock -ModuleName YardstickSupport Get-Command {
            [pscustomobject]@{ Name = 'curl.exe'; Source = 'C:\Windows\System32\curl.exe' }
        } -ParameterFilter { $Name -eq 'curl.exe' }
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.Warnings | Where-Object { $_ -match "curl.exe" } | Should -BeNullOrEmpty
    }

    It "Detects missing required modules as errors" {
        # Verify that the function checks for required modules.
        # If powershell-yaml IS installed, this validates no false positive.
        # If powershell-yaml is NOT installed, this validates the error is raised.
        $yamlInstalled = [bool](Get-Module -ListAvailable -Name 'powershell-yaml')
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        if (-not $yamlInstalled) {
            $result.IsValid | Should -BeFalse
            $result.Errors | Where-Object { $_ -match "powershell-yaml" } | Should -Not -BeNullOrEmpty
        } else {
            $result.Errors | Where-Object { $_ -match "powershell-yaml" } | Should -BeNullOrEmpty
        }
    }

    It "Detects missing optional modules as warnings not errors" {
        # Selenium/TUN.CredentialManager are likely not installed on CI/test machines
        $seleniumInstalled = $null -ne (Get-Module -ListAvailable -Name 'Selenium')
        if (-not $seleniumInstalled) {
            $result = Test-Prerequisites -ToolsPath $script:toolsPath
            $result.Warnings | Where-Object { $_ -match "Selenium" } | Should -Not -BeNullOrEmpty
            # Even with missing optional modules, IsValid depends on required modules only
        }
    }

    It "Validates HttpClient .NET type is available" {
        # System.Net.Http.HttpClient should be available in modern .NET
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.Errors | Where-Object { $_ -match "HttpClient" } | Should -BeNullOrEmpty
    }

    It "Warns when only an older Selenium is installed" {
        Mock -ModuleName YardstickSupport Get-YardstickOptionalModule {
            [pscustomobject]@{ Name = 'Selenium'; Purpose = 'recipes that drive a browser'; UsedByRecipes = $true; MinimumVersion = '4.0.0-preview3'; Prerelease = $true; InstalledVersion = '3.0.1'; IsCurrent = $false }
        }
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.Warnings | Where-Object { $_ -match "Selenium' 3\.0\.1 is installed.*4\.0\.0-preview3" } | Should -Not -BeNullOrEmpty
    }

    It "Does not warn about modules recipes don't use, such as Pester" {
        Mock -ModuleName YardstickSupport Get-YardstickOptionalModule {
            [pscustomobject]@{ Name = 'Pester'; Purpose = 'running the unit tests'; UsedByRecipes = $false; MinimumVersion = '5.0.0'; InstalledVersion = '3.4.0'; IsCurrent = $false }
            [pscustomobject]@{ Name = 'TUN.CredentialManager'; Purpose = 'recipes that sign in to a vendor portal'; UsedByRecipes = $true; MinimumVersion = $null; InstalledVersion = $null; IsCurrent = $false }
        }
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.Warnings | Where-Object { $_ -match "Pester" } | Should -BeNullOrEmpty
        $result.Warnings | Where-Object { $_ -match "TUN\.CredentialManager' is not installed.*vendor portal" } | Should -Not -BeNullOrEmpty
    }
}

Describe "Get-YardstickOptionalModule" {
    It "Treats Selenium as current = <Expected> for installed versions <Installed>" -ForEach @(
        @{ Installed = @(@{ V = '3.0.1'; P = '' }); Expected = $false }
        @{ Installed = @(@{ V = '4.0.0'; P = 'preview2' }); Expected = $false }
        @{ Installed = @(@{ V = '3.0.1'; P = '' }, @{ V = '4.0.0'; P = 'preview3' }); Expected = $true }
        @{ Installed = @(@{ V = '4.0.0'; P = '' }); Expected = $true }
        @{ Installed = @(); Expected = $false }
    ) {
        $Global:YardstickTestSeleniumVersions = $Installed
        Mock -ModuleName YardstickSupport Get-Module { }
        Mock -ModuleName YardstickSupport Get-Module {
            foreach ($i in $Global:YardstickTestSeleniumVersions) {
                [pscustomobject]@{ Name = 'Selenium'; Version = [version]$i.V; PrivateData = @{ PSData = @{ Prerelease = $i.P } } }
            }
        } -ParameterFilter { $Name -eq 'Selenium' }

        $selenium = Get-YardstickOptionalModule | Where-Object Name -eq 'Selenium'
        $selenium.IsCurrent | Should -Be $Expected
        $selenium.Prerelease | Should -BeTrue
        Remove-Variable -Name YardstickTestSeleniumVersions -Scope Global
    }

    It "Treats the Pester that ships with Windows as too old and skips its publisher check" {
        Mock -ModuleName YardstickSupport Get-Module { }
        Mock -ModuleName YardstickSupport Get-Module {
            [pscustomobject]@{ Name = 'Pester'; Version = [version]'3.4.0'; PrivateData = $null }
        } -ParameterFilter { $Name -eq 'Pester' }

        $pester = Get-YardstickOptionalModule | Where-Object Name -eq 'Pester'
        $pester.InstalledVersion | Should -Be '3.4.0'
        $pester.IsCurrent | Should -BeFalse
        $pester.SkipPublisherCheck | Should -BeTrue
        $pester.UsedByRecipes | Should -BeFalse
    }
}
