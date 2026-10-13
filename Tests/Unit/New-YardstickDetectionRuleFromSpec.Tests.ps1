BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickSupport.psm1" -Force
}

Describe "New-YardstickDetectionRuleFromSpec" {
    Context "registry" {
        It "builds a version comparison from an explicit value" {
            $rule = New-YardstickDetectionRuleFromSpec -DefaultVersion '9.9.9' -Spec @{
                detectionType              = 'registry'
                registryDetectionMethod    = 'version'
                registryDetectionKey       = 'HKEY_LOCAL_MACHINE\SOFTWARE\TestApp'
                registryDetectionValueName = 'DisplayVersion'
                registryDetectionOperator  = 'greaterThanOrEqual'
                registryDetectionValue     = '1.2.3'
            }
            $rule.'@odata.type' | Should -Be '#microsoft.graph.win32LobAppRegistryDetection'
            $rule.detectionType | Should -Be 'version'
            $rule.operator | Should -Be 'greaterThanOrEqual'
            $rule.keyPath | Should -Be 'HKEY_LOCAL_MACHINE\SOFTWARE\TestApp'
            $rule.valueName | Should -Be 'DisplayVersion'
            $rule.detectionValue | Should -Be '1.2.3'
        }

        It "falls back to the default version when no value is given" {
            $rule = New-YardstickDetectionRuleFromSpec -DefaultVersion '9.9.9' -Spec @{
                detectionType              = 'registry'
                registryDetectionMethod    = 'version'
                registryDetectionKey       = 'HKEY_LOCAL_MACHINE\SOFTWARE\TestApp'
                registryDetectionValueName = 'DisplayVersion'
                registryDetectionOperator  = 'greaterThanOrEqual'
            }
            $rule.detectionValue | Should -Be '9.9.9'
        }

        It "builds a string comparison and honours Is32BitApp" {
            $rule = New-YardstickDetectionRuleFromSpec -Is32BitApp $true -Spec @{
                detectionType              = 'registry'
                registryDetectionMethod    = 'string'
                registryDetectionKey       = 'HKEY_LOCAL_MACHINE\SOFTWARE\TestApp'
                registryDetectionValueName = 'DisplayName'
                registryDetectionOperator  = 'equal'
                registryDetectionValue     = 'Test App ESR (x64 en-US)'
            }
            $rule.detectionType | Should -Be 'string'
            $rule.operator | Should -Be 'equal'
            $rule.detectionValue | Should -Be 'Test App ESR (x64 en-US)'
            $rule.check32BitOn64System | Should -BeTrue
        }

        It "builds a key existence rule" {
            $rule = New-YardstickDetectionRuleFromSpec -Spec @{
                detectionType           = 'registry'
                registryDetectionMethod = 'exists'
                registryDetectionKey    = 'HKEY_LOCAL_MACHINE\SOFTWARE\TestApp'
            }
            $rule.detectionType | Should -Be 'exists'
            $rule.operator | Should -Be 'notConfigured'
        }

        It "throws on an unsupported method" {
            { New-YardstickDetectionRuleFromSpec -Spec @{
                detectionType           = 'registry'
                registryDetectionMethod = 'notExists'
                registryDetectionKey    = 'HKEY_LOCAL_MACHINE\SOFTWARE\TestApp'
            } } | Should -Throw "*Unsupported registryDetectionMethod 'notExists'*"
        }
    }

    Context "file" {
        It "builds a version rule that falls back to the default version" {
            $rule = New-YardstickDetectionRuleFromSpec -DefaultVersion '4.5.6' -Spec @{
                detectionType         = 'file'
                fileDetectionMethod   = 'version'
                fileDetectionPath     = 'C:\Program Files\TestApp'
                fileDetectionName     = 'testapp.exe'
                fileDetectionOperator = 'greaterThanOrEqual'
            }
            $rule.'@odata.type' | Should -Be '#microsoft.graph.win32LobAppFileSystemDetection'
            $rule.detectionType | Should -Be 'version'
            $rule.detectionValue | Should -Be '4.5.6'
            $rule.path | Should -Be 'C:\Program Files\TestApp'
            $rule.fileOrFolderName | Should -Be 'testapp.exe'
        }

        It "prefers an explicit fileDetectionVersion" {
            $rule = New-YardstickDetectionRuleFromSpec -DefaultVersion '4.5.6' -Spec @{
                detectionType         = 'file'
                fileDetectionMethod   = 'version'
                fileDetectionPath     = 'C:\Program Files\TestApp'
                fileDetectionName     = 'testapp.exe'
                fileDetectionOperator = 'greaterThanOrEqual'
                fileDetectionVersion  = '4.5.6.100'
            }
            $rule.detectionValue | Should -Be '4.5.6.100'
        }

        It "builds an existence rule" {
            $rule = New-YardstickDetectionRuleFromSpec -Spec @{
                detectionType       = 'file'
                fileDetectionMethod = 'exists'
                fileDetectionPath   = 'C:\Program Files\TestApp'
                fileDetectionName   = 'testapp.exe'
            }
            $rule.detectionType | Should -Be 'exists'
            $rule.operator | Should -Be 'notConfigured'
        }
    }

    Context "msi" {
        It "uses the product code and version" {
            $rule = New-YardstickDetectionRuleFromSpec -ProductCode '{CODE}' -DefaultVersion '1.0' -Spec @{ detectionType = 'msi' }
            $rule.'@odata.type' | Should -Be '#microsoft.graph.win32LobAppProductCodeDetection'
            $rule.productCode | Should -Be '{CODE}'
            $rule.productVersion | Should -Be '1.0'
        }
    }

    It "refuses script detection" {
        { New-YardstickDetectionRuleFromSpec -Spec @{ detectionType = 'script' } } | Should -Throw "*Unsupported detectionType 'script'*"
    }
}
