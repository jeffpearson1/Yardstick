BeforeAll {
    Import-Module "$PSScriptRoot\..\Modules\YardstickGraph.psm1" -Force
    Import-Module "$PSScriptRoot\..\Modules\YardstickIntune.psm1" -Force
}

Describe 'Yardstick Intune payload builders' {
    It 'maps architecture and Windows release requirements' {
        $rule = New-YardstickWin32AppRequirementRule -Architecture AllWithARM64 -MinimumSupportedWindowsRelease W11_24H2
        $rule.allowedArchitectures | Should -Be 'x64,x86,arm64'
        $rule.applicableArchitectures | Should -Be 'none'
        $rule.minimumSupportedWindowsRelease | Should -Be 'Windows11_24H2'
    }

    It 'builds MSI detection rules in Graph shape' {
        $rule = New-YardstickWin32AppDetectionRuleMsi -ProductCode '{CODE}' -ProductVersion '1.2.3' -ProductVersionOperator greaterThanOrEqual
        $rule.'@odata.type' | Should -Be '#microsoft.graph.win32LobAppProductCodeDetection'
        $rule.productCode | Should -Be '{CODE}'
        $rule.productVersionOperator | Should -Be 'greaterThanOrEqual'
    }

    It 'preserves the MIME type for JPEG icons' {
        $iconPath = Join-Path $TestDrive 'icon.jpg'
        [io.file]::WriteAllBytes($iconPath, [byte[]](1, 2, 3))
        $icon = New-YardstickWin32AppIcon -FilePath $iconPath
        $icon.type | Should -Be 'image/jpeg'
        $icon.value | Should -Be ([convert]::ToBase64String([byte[]](1, 2, 3)))
    }

    It 'builds an EXE app body when optional requirements are absent' {
        InModuleScope YardstickIntune {
            $metadata = [xml]'<ApplicationInfo><FileName>payload.bin</FileName><SetupFile>setup.exe</SetupFile></ApplicationInfo>'
            $requirement = New-YardstickWin32AppRequirementRule -Architecture x64 -MinimumSupportedWindowsRelease W10_22H2
            $body = New-YardstickWin32AppBody -Metadata $metadata -DisplayName 'App' -Description 'Desc' -Publisher 'Pub' -InstallCommandLine 'install.cmd' -UninstallCommandLine 'uninstall.cmd' -DetectionRule @{ rule = 'one' } -RequirementRule $requirement
            $body.'@odata.type' | Should -Be '#microsoft.graph.win32LobApp'
            $body.allowedArchitectures | Should -Be 'x64'
            $body.minimumFreeDiskSpaceInMB | Should -BeNullOrEmpty
            $body.returnCodes.Count | Should -Be 5
            $body.roleScopeTagIds | Should -Be @('0')
        }
    }
}

Describe 'Yardstick Intune assignments' {
    BeforeEach {
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickIntune {
            if ($Resource -eq 'deviceAppManagement/mobileApps/app') {
                return [pscustomobject]@{ id = 'app'; displayName = 'App' }
            }
            if ($Resource -eq 'deviceAppManagement/mobileApps/app/assignments') {
                return @([pscustomobject]@{
                    id = 'assignment-1'
                    intent = 'required'
                    target = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget'; groupId = 'group-1' }
                })
            }
        }
    }

    It 'projects a single assignment even when settings are absent' {
        $assignment = @(Get-YardstickWin32AppAssignment -ID app)
        $assignment.Count | Should -Be 1
        $assignment[0].GroupID | Should -Be 'group-1'
        $assignment[0].GroupMode | Should -Be 'Exclude'
        $assignment[0].Notifications | Should -BeNullOrEmpty
    }

    It 'removes every assignment from an app' {
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickIntune {
            if ($Resource -eq 'deviceAppManagement/mobileApps/app') {
                return [pscustomobject]@{ id = 'app'; displayName = 'App' }
            }
            if ($Resource -eq 'deviceAppManagement/mobileApps/app/assignments') {
                return @([pscustomobject]@{ id = 'one' }, [pscustomobject]@{ id = 'two' })
            }
        }
        Remove-YardstickWin32AppAssignment -ID app
        Should -Invoke Invoke-YardstickGraphRequest -ModuleName YardstickIntune -Times 2 -Exactly -ParameterFilter {
            $Method -eq 'Delete' -and $Resource -match '/assignments/(one|two)$'
        }
    }

    It 'returns the existing assignment when Intune reports a duplicate' {
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickIntune {
            if ($Method -eq 'Post') {
                throw 'The MobileApp Assignment already exists, for AppId: app, AssignmentId: group-9_0_0.'
            }
            return @([pscustomobject]@{
                id     = 'existing-1'
                intent = 'available'
                target = [pscustomobject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-9' }
            })
        }
        $result = Add-YardstickWin32AppAssignmentGroup -Include -ID app -GroupID group-9 -Intent available
        $result.id | Should -Be 'existing-1'
    }

    It 'rethrows a duplicate error when no matching assignment can be found' {
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickIntune {
            if ($Method -eq 'Post') {
                throw 'The MobileApp Assignment already exists, for AppId: app, AssignmentId: group-9_0_0.'
            }
            return @()
        }
        { Add-YardstickWin32AppAssignmentGroup -Include -ID app -GroupID group-9 -Intent available } | Should -Throw '*already exists*'
    }
}

Describe 'Yardstick Intune relationships' {
    It 'preserves supersedence while replacing dependencies' {
        Mock Get-YardstickWin32AppSupersedence -ModuleName YardstickIntune {
            [ordered]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; targetId = 'old' }
        }
        Mock Set-YardstickWin32AppRelationships -ModuleName YardstickIntune {}
        $dependency = [ordered]@{ '@odata.type' = '#microsoft.graph.mobileAppDependency'; targetId = 'dependency' }
        Add-YardstickWin32AppDependency -ID app -Dependency $dependency
        Should -Invoke Set-YardstickWin32AppRelationships -ModuleName YardstickIntune -Times 1 -Exactly -ParameterFilter {
            $ID -eq 'app' -and $Relationships.Count -eq 2 -and $Relationships.targetId -contains 'old' -and $Relationships.targetId -contains 'dependency'
        }
    }
}
