BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\Modules\YardstickSupport.psm1"
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
        Mock -ModuleName YardstickSupport Get-Command { $null } -ParameterFilter { $Name -eq 'curl' }
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.Warnings | Where-Object { $_ -match "curl.exe" } | Should -Not -BeNullOrEmpty
    }

    It "Does not warn about curl.exe when it exists" {
        Mock -ModuleName YardstickSupport Get-Command { $null } -ParameterFilter { $Name -eq 'curl' }
        # Create a dummy curl.exe
        New-Item -Path (Join-Path $script:toolsPath "curl.exe") -ItemType File -Force | Out-Null
        $result = Test-Prerequisites -ToolsPath $script:toolsPath
        $result.Warnings | Where-Object { $_ -match "curl.exe" } | Should -BeNullOrEmpty
    }

    It "Does not warn when curl is only on PATH" {
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
}
