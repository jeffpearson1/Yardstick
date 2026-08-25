BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force
    # The IntuneWin32App cmdlets are mocked below, but Pester can only mock
    # commands it is able to resolve from the module's scope.
    Import-Module IntuneWin32App -ErrorAction Stop
}

Describe "Test-IsVersionDetection" {
    It "returns true for msi" {
        Test-IsVersionDetection -DetectionType 'msi' | Should -BeTrue
    }
    It "returns true for file+version" {
        Test-IsVersionDetection -DetectionType 'file' -FileDetectionMethod 'version' | Should -BeTrue
    }
    It "returns false for file+exists" {
        Test-IsVersionDetection -DetectionType 'file' -FileDetectionMethod 'exists' | Should -BeFalse
    }
    It "returns true for registry+version" {
        Test-IsVersionDetection -DetectionType 'registry' -RegistryDetectionMethod 'version' | Should -BeTrue
    }
    It "returns false for registry+exists" {
        Test-IsVersionDetection -DetectionType 'registry' -RegistryDetectionMethod 'exists' | Should -BeFalse
    }
    It "returns false for script" {
        Test-IsVersionDetection -DetectionType 'script' | Should -BeFalse
    }
}

Describe "Get-DetectAnchorName" {
    It "prefixes the reserved marker" {
        Get-DetectAnchorName -DisplayName 'Google Chrome' | Should -Be 'Ω DETECT - Google Chrome'
    }
}

Describe "New-SupersedenceObject" {
    BeforeEach {
        Mock -ModuleName YardstickSupport New-IntuneWin32AppSupersedence {
            [ordered]@{
                '@odata.type'      = '#microsoft.graph.mobileAppSupersedence'
                'supersedenceType' = $SupersedenceType.ToLower()
                'targetId'         = $ID
            }
        }
    }

    It "builds one ordered dictionary per explicit target" {
        $built = New-SupersedenceObject -TargetIds @('a', 'b') -Type 'Update'
        $built.Count | Should -Be 2
        $built[0].targetId | Should -Be 'a'
        $built[0].supersedenceType | Should -Be 'update'
    }

    It "drops targets that cannot be resolved instead of emitting nulls" {
        Mock -ModuleName YardstickSupport New-IntuneWin32AppSupersedence {
            if ($ID -eq 'missing') { return $null }
            [ordered]@{ 'targetId' = $ID; 'supersedenceType' = $SupersedenceType.ToLower() }
        }
        $built = New-SupersedenceObject -TargetIds @('good', 'missing') -Type 'Update'
        $built.Count | Should -Be 1
        $built[0].targetId | Should -Be 'good'
    }

    It "preserves the supersedence type when rebuilding from existing relationships" {
        $relationships = @(
            [PSCustomObject]@{ targetId = 'x'; supersedenceType = 'replace' },
            [PSCustomObject]@{ targetId = 'y'; supersedenceType = 'update' }
        )
        $built = New-SupersedenceObject -Relationships $relationships
        $built.Count | Should -Be 2
        $built[0].supersedenceType | Should -Be 'replace'
        $built[1].supersedenceType | Should -Be 'update'
    }

    It "returns an empty array rather than null when given nothing" {
        $built = New-SupersedenceObject -TargetIds @()
        $built.Count | Should -Be 0
    }
}

Describe "Set-YardstickSupersedence" {
    BeforeEach {
        # Default: nothing is attached beforehand, and the post-attach read-back
        # reports exactly what was submitted.
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship { @() }
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence {}
        Mock -ModuleName YardstickSupport New-IntuneWin32AppSupersedence {
            [ordered]@{ 'targetId' = $ID; 'supersedenceType' = $SupersedenceType.ToLower() }
        }
        Mock -ModuleName YardstickSupport Add-IntuneWin32AppSupersedence {
            $script:AttachedForNewApp = @($Supersedence | ForEach-Object {
                [PSCustomObject]@{ sourceId = $ID; targetId = $_.targetId; supersedenceType = $_.supersedenceType }
            })
        }
        # Read-back after the attach returns whatever the mocked Add recorded.
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            if ($Direction -eq 'Forward' -and $Id -eq 'new' -and $script:AttachedForNewApp) {
                return $script:AttachedForNewApp
            }
            return @()
        }
        $script:AttachedForNewApp = $null
    }

    It "attaches one supersedence entry per target" {
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        $targets = @(
            [PSCustomObject]@{ id = 'old1'; DisplayName = 'App (N-1)'; displayVersion = '2.0' },
            [PSCustomObject]@{ id = 'old2'; DisplayName = 'App (N-2)'; displayVersion = '1.0' }
        )
        $count = Set-YardstickSupersedence -NewApp $new -SupersededApps $targets -Type 'Update'
        $count | Should -Be 2
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppSupersedence -Times 1 -Exactly
    }

    It "reports what Intune actually stored, not what was submitted" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship { @() }
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        $old = [PSCustomObject]@{ id = 'old'; DisplayName = 'App (N-1)'; displayVersion = '2.0' }
        Set-YardstickSupersedence -NewApp $new -SupersededApps @($old) | Should -Be 0
    }

    It "returns 0 and skips Add when there are no targets" {
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        $count = Set-YardstickSupersedence -NewApp $new -SupersededApps @() -Type 'Update'
        $count | Should -Be 0
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppSupersedence -Times 0 -Exactly
    }

    It "clears stale links on the new app when there are no targets" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            @([PSCustomObject]@{ sourceId = 'new'; targetId = 'gone' })
        }
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        Set-YardstickSupersedence -NewApp $new -SupersededApps @() | Out-Null
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter { $ID -eq 'new' }
    }

    It "excludes the new app from targets to prevent self-supersedence" {
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        $targets = @(
            [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' },
            [PSCustomObject]@{ id = 'old'; DisplayName = 'App (N-1)'; displayVersion = '2.0' }
        )
        $count = Set-YardstickSupersedence -NewApp $new -SupersededApps $targets -Type 'Update'
        $count | Should -Be 1
    }

    It "de-duplicates repeated targets" {
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        $old = [PSCustomObject]@{ id = 'old'; DisplayName = 'App (N-1)'; displayVersion = '2.0' }
        $count = Set-YardstickSupersedence -NewApp $new -SupersededApps @($old, $old) -Type 'Update'
        $count | Should -Be 1
    }

    It "keeps the graph inside Intune's 10-node limit" {
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '20.0' }
        $targets = @(1..12 | ForEach-Object {
            [PSCustomObject]@{ id = "old$_"; DisplayName = "App (N-$_)"; displayVersion = "$_.0" }
        })
        $count = Set-YardstickSupersedence -NewApp $new -SupersededApps $targets -Type 'Replace'
        $count | Should -Be 9
    }

    It "never trims away an update-only anchor when capping" {
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '20.0' }
        $targets = @(1..12 | ForEach-Object {
            [PSCustomObject]@{ id = "old$_"; DisplayName = "App (N-$_)"; displayVersion = "$_.0" }
        })
        # The anchor is deliberately the oldest version - a naive newest-first trim drops it.
        $targets += [PSCustomObject]@{ id = 'anchor'; DisplayName = 'Ω DETECT - App'; displayVersion = '0.1' }
        $count = Set-YardstickSupersedence -NewApp $new -SupersededApps $targets -Type 'Replace' -UpdateOnlyIds @('anchor')
        $count | Should -Be 9
        Should -Invoke -ModuleName YardstickSupport New-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter { $ID -eq 'anchor' -and $SupersedenceType -eq 'Update' }
    }

    It "forces Update for ids listed in UpdateOnlyIds even when Type is Replace" {
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        $targets = @(
            [PSCustomObject]@{ id = 'old';    DisplayName = 'App (N-1)';    displayVersion = '2.0' },
            [PSCustomObject]@{ id = 'anchor'; DisplayName = 'Ω DETECT - App'; displayVersion = '1.0' }
        )
        Set-YardstickSupersedence -NewApp $new -SupersededApps $targets -Type 'Replace' -UpdateOnlyIds @('anchor') | Out-Null
        Should -Invoke -ModuleName YardstickSupport New-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter { $ID -eq 'anchor' -and $SupersedenceType -eq 'Update' }
        Should -Invoke -ModuleName YardstickSupport New-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter { $ID -eq 'old' -and $SupersedenceType -eq 'Replace' }
    }

    It "strips stale forward links off targets before re-parenting them" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            if ($Id -eq 'old') { return @([PSCustomObject]@{ sourceId = 'old'; targetId = 'older' }) }
            if ($Id -eq 'new' -and $script:AttachedForNewApp) { return $script:AttachedForNewApp }
            return @()
        }
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        $old = [PSCustomObject]@{ id = 'old'; DisplayName = 'App (N-1)'; displayVersion = '2.0' }
        Set-YardstickSupersedence -NewApp $new -SupersededApps @($old) | Out-Null
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter { $ID -eq 'old' }
    }
}

Describe "Set-AssignmentAutoUpdate" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest {
            if ($Method -ne 'Patch') {
                return @(
                    [PSCustomObject]@{
                        id       = 'assign-required'
                        intent   = 'required'
                        target   = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget' }
                        settings = [PSCustomObject]@{ notifications = 'showAll' }
                    },
                    [PSCustomObject]@{
                        id       = 'assign-available'
                        intent   = 'available'
                        target   = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget' }
                        settings = [PSCustomObject]@{ notifications = 'hideAll'; deliveryOptimizationPriority = 'foreground' }
                    },
                    [PSCustomObject]@{
                        id       = 'assign-excluded'
                        intent   = 'available'
                        target   = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget' }
                        settings = $null
                    }
                )
            }
            return $null
        }
    }

    It "patches only available-intent group assignments" {
        $count = Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true -IntentFilter 'available'
        $count | Should -Be 1
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and $Resource -like '*assignments/assign-available'
        }
    }

    It "uses the documented autoUpdateSupersededAppsState property" {
        Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true | Out-Null
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and
            $Body.settings.autoUpdateSettings.'@odata.type' -eq '#microsoft.graph.win32LobAppAutoUpdateSettings' -and
            $Body.settings.autoUpdateSettings.autoUpdateSupersededAppsState -eq 'enabled'
        }
    }

    It "preserves the existing assignment settings on the patch" {
        Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true | Out-Null
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and
            $Body.settings.notifications -eq 'hideAll' -and
            $Body.settings.deliveryOptimizationPriority -eq 'foreground'
        }
    }

    It "writes notConfigured when disabling" {
        Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $false | Out-Null
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and $Body.settings.autoUpdateSettings.autoUpdateSupersededAppsState -eq 'notConfigured'
        }
    }

    It "is a no-op when the state already matches" {
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest {
            if ($Method -ne 'Patch') {
                return @(
                    [PSCustomObject]@{
                        id       = 'assign-available'
                        intent   = 'available'
                        target   = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget' }
                        settings = [PSCustomObject]@{
                            autoUpdateSettings = [PSCustomObject]@{ autoUpdateSupersededAppsState = 'enabled' }
                        }
                    }
                )
            }
            return $null
        }
        $count = Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true
        $count | Should -Be 1
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 0 -Exactly -ParameterFilter { $Method -eq 'Patch' }
    }

    It "returns 0 when there are no matching assignments" {
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest { @() }
        Set-AssignmentAutoUpdate -AppId 'app1' | Should -Be 0
    }
}

Describe "Set-AssignmentAutoUpdate skip groups" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest {
            if ($Method -ne 'Patch') {
                return @(
                    [PSCustomObject]@{
                        id       = 'assign-normal'
                        intent   = 'available'
                        target   = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-normal' }
                        settings = [PSCustomObject]@{ notifications = 'showAll' }
                    },
                    [PSCustomObject]@{
                        id       = 'assign-skipped'
                        intent   = 'available'
                        target   = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'group-skipped' }
                        settings = [PSCustomObject]@{
                            notifications      = 'showAll'
                            autoUpdateSettings = [PSCustomObject]@{ autoUpdateSupersededAppsState = 'enabled' }
                        }
                    },
                    [PSCustomObject]@{
                        id       = 'assign-all-devices'
                        intent   = 'available'
                        target   = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.allDevicesAssignmentTarget' }
                        settings = [PSCustomObject]@{ notifications = 'showAll' }
                    }
                )
            }
            return $null
        }
    }

    It "enables auto-update on groups that are not in the skip list" {
        Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true -SkipGroupIds @('group-skipped') | Out-Null
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and $Resource -like '*assignments/assign-normal' -and
            $Body.settings.autoUpdateSettings.autoUpdateSupersededAppsState -eq 'enabled'
        }
    }

    It "forces a skipped group back to notConfigured" {
        Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true -SkipGroupIds @('group-skipped') | Out-Null
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and $Resource -like '*assignments/assign-skipped' -and
            $Body.settings.autoUpdateSettings.autoUpdateSupersededAppsState -eq 'notConfigured'
        }
    }

    It "matches group ids case-insensitively" {
        Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true -SkipGroupIds @('GROUP-SKIPPED') | Out-Null
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and $Resource -like '*assignments/assign-skipped'
        }
    }

    It "does not count skipped assignments in the returned total" {
        # assign-normal and assign-all-devices are enabled; assign-skipped is not.
        Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true -SkipGroupIds @('group-skipped') | Should -Be 2
    }

    It "cannot skip targets that carry no group id" {
        Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true -SkipGroupIds @('group-skipped') | Out-Null
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and $Resource -like '*assignments/assign-all-devices' -and
            $Body.settings.autoUpdateSettings.autoUpdateSupersededAppsState -eq 'enabled'
        }
    }

    It "leaves every assignment alone when the skip list is empty" {
        $count = Set-AssignmentAutoUpdate -AppId 'app1' -Enabled $true
        $count | Should -Be 3
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 0 -Exactly -ParameterFilter {
            $Method -eq 'Patch' -and $Body.settings.autoUpdateSettings.autoUpdateSupersededAppsState -eq 'notConfigured'
        }
    }
}

Describe "Remove-YardstickApp" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Remove-IntuneWin32App {}
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence {}
        Mock -ModuleName YardstickSupport Remove-SupersedenceReference {}
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppDependency {}
        Mock -ModuleName YardstickSupport Remove-DependencyReference {}
        # Nothing but supersedence on these apps: no assignments, no dependencies.
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest { @() }
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency { @() }
        # Deletion succeeded: the app is no longer resolvable.
        Mock -ModuleName YardstickSupport Get-IntuneWin32App { $null }
    }

    AfterEach {
        Remove-Variable -Name MockSupersedenceCleared -Scope Global -ErrorAction SilentlyContinue
    }

    It "detaches superseding parents before deleting" {
        # The mock models the detach taking effect: the read-back has to come up
        # empty or Remove-YardstickApp refuses the delete.
        $Global:MockSupersedenceCleared = $false
        Mock -ModuleName YardstickSupport Remove-SupersedenceReference { $Global:MockSupersedenceCleared = $true }
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            if ($Global:MockSupersedenceCleared) { return @() }
            @([PSCustomObject]@{ sourceId = 'parent'; targetId = 'doomed' })
        }
        Remove-YardstickApp -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' })
        Should -Invoke -ModuleName YardstickSupport Remove-SupersedenceReference -Times 1 -Exactly -ParameterFilter {
            $ParentId -eq 'parent' -and $TargetId -eq 'doomed'
        }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32App -Times 1 -Exactly -ParameterFilter { $Id -eq 'doomed' }
    }

    It "clears its own forward links before deleting" {
        $Global:MockSupersedenceCleared = $false
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence { $Global:MockSupersedenceCleared = $true }
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            if ($Global:MockSupersedenceCleared) { return @() }
            @([PSCustomObject]@{ sourceId = 'doomed'; targetId = 'older' })
        }
        Remove-YardstickApp -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' })
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter { $ID -eq 'doomed' }
        Should -Invoke -ModuleName YardstickSupport Remove-SupersedenceReference -Times 0 -Exactly
    }

    It "makes no relationship calls when the app is unrelated" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship { @() }
        Remove-YardstickApp -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' })
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence -Times 0 -Exactly
        Should -Invoke -ModuleName YardstickSupport Remove-SupersedenceReference -Times 0 -Exactly
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32App -Times 1 -Exactly
    }

    It "throws when Intune silently refuses the delete" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship { @() }
        Mock -ModuleName YardstickSupport Get-IntuneWin32App {
            [PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }
        }
        { Remove-YardstickApp -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) } |
            Should -Throw -ExpectedMessage "*still reports app*"
    }

    It "does not attempt the delete when a link survives the unwind" {
        # A supersedence link that will not go away: the read-back still sees it,
        # so the delete could only fail. Naming the link beats a bare failure.
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            @([PSCustomObject]@{ sourceId = 'doomed'; targetId = 'older' })
        }
        { Remove-YardstickApp -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) } |
            Should -Throw -ExpectedMessage "*supersedes older*"
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32App -Times 0 -Exactly
    }

    It "refuses the delete when the links could not be read back at all" {
        # An expired token makes the read throw. Treating that as "nothing left"
        # would delete an app whose links were never confirmed gone.
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship { @() }
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest { throw "Graph authentication header is missing." }
        { Remove-YardstickApp -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) } |
            Should -Throw -ExpectedMessage "*could not be read back*"
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32App -Times 0 -Exactly
    }
}

Describe "Clear-YardstickAppLink" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence {}
        Mock -ModuleName YardstickSupport Remove-SupersedenceReference {}
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppDependency {}
        Mock -ModuleName YardstickSupport Remove-DependencyReference {}
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship { @() }
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency { @() }
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest { @() }
    }

    AfterEach {
        Remove-Variable -Name MockDependentDetached, MockDependenciesCleared, MockAssignmentsRead `
            -Scope Global -ErrorAction SilentlyContinue
    }

    It "deletes every assignment by its own id" {
        # By-id is the point: Remove-IntuneWin32AppAssignmentGroup matches on the
        # target, so two assignments sharing a group would go as a pair.
        $Global:MockAssignmentsRead = $false
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest {
            if ($Method -eq 'Delete') { return @() }
            if ($Global:MockAssignmentsRead) { return @() }
            $Global:MockAssignmentsRead = $true
            @(
                [PSCustomObject]@{ id = 'a1'; intent = 'required';  target = @{ groupId = 'g1' } },
                [PSCustomObject]@{ id = 'a2'; intent = 'available'; target = @{ groupId = 'g1' } }
            )
        }
        $surviving = @(Clear-YardstickAppLink -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) -RetryDelaySeconds 0)
        $surviving | Should -BeNullOrEmpty
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Delete' -and $Resource -eq 'deviceAppManagement/mobileApps/doomed/assignments/a1'
        }
        Should -Invoke -ModuleName YardstickSupport Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $Method -eq 'Delete' -and $Resource -eq 'deviceAppManagement/mobileApps/doomed/assignments/a2'
        }
    }

    It "detaches the apps that depend on this one" {
        # The mock models the removal actually taking effect, so the read-back at
        # the end of Clear-YardstickAppLink sees a clean app.
        $Global:MockDependentDetached = $false
        Mock -ModuleName YardstickSupport Remove-DependencyReference { $Global:MockDependentDetached = $true }
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            if ($ID -ne 'doomed' -or $Global:MockDependentDetached) { return @() }
            @([PSCustomObject]@{ targetId = 'dependent'; targetType = 'parent'; dependencyType = 'detect' })
        }
        $surviving = @(Clear-YardstickAppLink -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) -RetryDelaySeconds 0)
        $surviving | Should -BeNullOrEmpty
        Should -Invoke -ModuleName YardstickSupport Remove-DependencyReference -Times 1 -Exactly -ParameterFilter {
            $ParentId -eq 'dependent' -and $TargetId -eq 'doomed'
        }
    }

    It "clears the app's own dependencies" {
        $Global:MockDependenciesCleared = $false
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppDependency { $Global:MockDependenciesCleared = $true }
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            if ($ID -ne 'doomed' -or $Global:MockDependenciesCleared) { return @() }
            @([PSCustomObject]@{ targetId = 'vcredist'; targetType = 'child'; dependencyType = 'autoInstall' })
        }
        $surviving = @(Clear-YardstickAppLink -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) -RetryDelaySeconds 0)
        $surviving | Should -BeNullOrEmpty
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppDependency -Times 1 -Exactly -ParameterFilter { $ID -eq 'doomed' }
    }

    It "reports a dependency that refuses to go away" {
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            if ($ID -ne 'doomed') { return @() }
            @([PSCustomObject]@{ targetId = 'vcredist'; targetType = 'child'; dependencyType = 'detect' })
        }
        $surviving = @(Clear-YardstickAppLink -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) -RetryDelaySeconds 0)
        $surviving | Should -Contain 'dependency on vcredist'
    }

    It "reports every kind of link that survives" {
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest {
            if ($Method -eq 'Delete') { return @() }
            @([PSCustomObject]@{ id = 'a1'; intent = 'required'; target = @{ groupId = 'g1' } })
        }
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            if ($ID -ne 'doomed') { return @() }
            @(
                [PSCustomObject]@{ targetId = 'dependent'; targetType = 'parent'; dependencyType = 'detect' },
                [PSCustomObject]@{ targetId = 'vcredist';  targetType = 'child';  dependencyType = 'detect' }
            )
        }
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            @(
                [PSCustomObject]@{ sourceId = 'newest'; targetId = 'doomed' },
                [PSCustomObject]@{ sourceId = 'doomed'; targetId = 'older' }
            )
        }
        $surviving = @(Clear-YardstickAppLink -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) -RetryDelaySeconds 0)
        $surviving | Should -Contain 'assignment a1'
        $surviving | Should -Contain 'dependency from dependent'
        $surviving | Should -Contain 'dependency on vcredist'
        $surviving | Should -Contain 'superseded by newest'
        $surviving | Should -Contain 'supersedes older'
    }

    It "treats a read it could not complete as a surviving link" {
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency { throw "Authentication token was not found" }
        $surviving = @(Clear-YardstickAppLink -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' }) -RetryDelaySeconds 0)
        $surviving | Should -Contain 'dependencies (could not be read back)'
    }

    It "detaches the parent named by targetType on a real Graph relationship" {
        # Graph reports the link from the child's side as sourceId=<child>,
        # targetId=<parent>, targetType='parent'. The parent id is the one that is
        # not the app being deleted.
        $Global:MockSupersedenceCleared = $false
        Mock -ModuleName YardstickSupport Remove-SupersedenceReference { $Global:MockSupersedenceCleared = $true }
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            if ($Global:MockSupersedenceCleared) { return @() }
            @([PSCustomObject]@{ sourceId = 'doomed'; targetId = 'newest'; targetType = 'parent'; supersedenceType = 'update' })
        }
        $surviving = @(Clear-YardstickAppLink -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-2)' }) -RetryDelaySeconds 0)
        $surviving | Should -BeNullOrEmpty
        Should -Invoke -ModuleName YardstickSupport Remove-SupersedenceReference -Times 1 -Exactly -ParameterFilter {
            $ParentId -eq 'newest' -and $TargetId -eq 'doomed'
        }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence -Times 0 -Exactly
        Remove-Variable -Name MockSupersedenceCleared -Scope Global -ErrorAction SilentlyContinue
    }
}

Describe "Remove-DependencyReference" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppDependency {}
        Mock -ModuleName YardstickSupport Add-IntuneWin32AppDependency {}
        Mock -ModuleName YardstickSupport New-IntuneWin32AppDependency {
            [ordered]@{ 'targetId' = $ID; 'dependencyType' = $DependencyType }
        }
    }

    It "clears dependencies entirely when the removed target was the only one" {
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            @([PSCustomObject]@{ targetId = 'doomed'; targetType = 'child'; dependencyType = 'detect' })
        }
        Remove-DependencyReference -ParentId 'parent' -TargetId 'doomed'
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppDependency -Times 1 -Exactly -ParameterFilter { $ID -eq 'parent' }
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppDependency -Times 0 -Exactly
    }

    It "rebuilds the remaining targets when other dependencies survive" {
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            @(
                [PSCustomObject]@{ targetId = 'doomed'; targetType = 'child'; dependencyType = 'detect' },
                [PSCustomObject]@{ targetId = 'keeper'; targetType = 'child'; dependencyType = 'autoInstall' }
            )
        }
        Remove-DependencyReference -ParentId 'parent' -TargetId 'doomed'
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppDependency -Times 1 -Exactly -ParameterFilter {
            $Dependency.Count -eq 1 -and $Dependency[0].targetId -eq 'keeper' -and $Dependency[0].dependencyType -eq 'AutoInstall'
        }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppDependency -Times 0 -Exactly
    }

    It "leaves the parent's own parent entries alone" {
        # A 'parent' entry names an app that depends on the parent - not the
        # parent's dependency, and not ours to rewrite.
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            @(
                [PSCustomObject]@{ targetId = 'grandparent'; targetType = 'parent'; dependencyType = 'detect' },
                [PSCustomObject]@{ targetId = 'doomed';      targetType = 'child';  dependencyType = 'detect' }
            )
        }
        Remove-DependencyReference -ParentId 'parent' -TargetId 'doomed'
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppDependency -Times 1 -Exactly -ParameterFilter { $ID -eq 'parent' }
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppDependency -Times 0 -Exactly
    }

    It "does nothing when the parent does not reference the target" {
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            @([PSCustomObject]@{ targetId = 'someone-else'; targetType = 'child'; dependencyType = 'detect' })
        }
        Remove-DependencyReference -ParentId 'parent' -TargetId 'doomed'
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppDependency -Times 0 -Exactly
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppDependency -Times 0 -Exactly
    }
}

Describe "Remove-SupersedenceReference" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence {}
        Mock -ModuleName YardstickSupport Add-IntuneWin32AppSupersedence {}
        Mock -ModuleName YardstickSupport New-IntuneWin32AppSupersedence {
            [ordered]@{ 'targetId' = $ID; 'supersedenceType' = $SupersedenceType.ToLower() }
        }
    }

    It "clears supersedence entirely when the removed target was the only one" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            @([PSCustomObject]@{ sourceId = 'parent'; targetId = 'doomed'; supersedenceType = 'update' })
        }
        Remove-SupersedenceReference -ParentId 'parent' -TargetId 'doomed'
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter { $ID -eq 'parent' }
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppSupersedence -Times 0 -Exactly
    }

    It "rebuilds the remaining targets when other links survive" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            @(
                [PSCustomObject]@{ sourceId = 'parent'; targetId = 'doomed'; supersedenceType = 'update' },
                [PSCustomObject]@{ sourceId = 'parent'; targetId = 'keeper'; supersedenceType = 'replace' }
            )
        }
        Remove-SupersedenceReference -ParentId 'parent' -TargetId 'doomed'
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter {
            $Supersedence.Count -eq 1 -and $Supersedence[0].targetId -eq 'keeper'
        }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence -Times 0 -Exactly
    }

    It "does nothing when the parent does not reference the target" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            @([PSCustomObject]@{ sourceId = 'parent'; targetId = 'someone-else'; supersedenceType = 'update' })
        }
        Remove-SupersedenceReference -ParentId 'parent' -TargetId 'doomed'
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppSupersedence -Times 0 -Exactly
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence -Times 0 -Exactly
    }
}

Describe "Get-YardstickSupersedenceRelationship" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest {
            @(
                [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; sourceId = 'me';     targetId = 'child' },
                [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; sourceId = 'parent'; targetId = 'me' },
                [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.mobileAppDependency';   sourceId = 'me';     targetId = 'dep' }
            )
        }
    }

    It "ignores dependency relationships" {
        (Get-YardstickSupersedenceRelationship -Id 'me').Count | Should -Be 2
    }

    It "Forward returns only links where this app is the parent" {
        $forward = Get-YardstickSupersedenceRelationship -Id 'me' -Direction Forward
        $forward.Count | Should -Be 1
        $forward[0].targetId | Should -Be 'child'
    }

    It "Reverse returns only links where this app is the child" {
        $reverse = Get-YardstickSupersedenceRelationship -Id 'me' -Direction Reverse
        $reverse.Count | Should -Be 1
        $reverse[0].sourceId | Should -Be 'parent'
    }

    It "reads direction off targetType, not sourceId" {
        # What Graph actually returns: the collection is reported from the queried
        # app's perspective, so sourceId is 'me' on BOTH links and only targetType
        # says which way each one points. Keying off sourceId called them both
        # forward, so a superseded app was never detached from its parent and its
        # deletion failed with "supersedes <parent id>".
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest {
            @(
                [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; sourceId = 'me'; targetId = 'newest'; targetType = 'parent' },
                [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; sourceId = 'me'; targetId = 'older';  targetType = 'child' }
            )
        }
        $forward = @(Get-YardstickSupersedenceRelationship -Id 'me' -Direction Forward)
        $forward.Count | Should -Be 1
        $forward[0].targetId | Should -Be 'older'

        $reverse = @(Get-YardstickSupersedenceRelationship -Id 'me' -Direction Reverse)
        $reverse.Count | Should -Be 1
        $reverse[0].targetId | Should -Be 'newest'
    }

    It "returns an empty array when Graph fails" {
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest { throw "boom" }
        (Get-YardstickSupersedenceRelationship -Id 'me').Count | Should -Be 0
    }

    It "reports a true count of <expected> to a caller that wraps the call in @()" -ForEach @(
        @{ expected = 0 }
        @{ expected = 1 }
        @{ expected = 2 }
        @{ expected = 3 }
    ) {
        # Regression: the returns used to be comma-wrapped, which handed the
        # caller a single item that was itself the array. Every count - 0, 2, 3 -
        # then measured as 1, which is what produced "Expected 2 supersedence
        # target(s) ... but Intune reports 1" and made the stale-link strip in
        # Set-YardstickSupersedence fire against apps that had no links at all.
        $n = $expected
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest {
            if ($n -eq 0) { return @() }
            return @(1..$n | ForEach-Object {
                [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.mobileAppSupersedence'; sourceId = 'me'; targetId = "t$_" }
            })
        }
        # Exactly how Set-YardstickSupersedence consumes it.
        @(Get-YardstickSupersedenceRelationship -Id 'me' -Direction Forward).Count | Should -Be $expected
    }
}

Describe "Get-SameAppAllVersions" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Invoke-WithRetry {
            @(
                [PSCustomObject]@{ id = '1'; DisplayName = 'App';           displayVersion = '3.0'; createdDateTime = (Get-Date) },
                [PSCustomObject]@{ id = '2'; DisplayName = 'App (N-1)';     displayVersion = '2.0'; createdDateTime = (Get-Date) },
                [PSCustomObject]@{ id = '3'; DisplayName = 'Ω DETECT - App';  displayVersion = '1.0'; createdDateTime = (Get-Date) },
                [PSCustomObject]@{ id = '4'; DisplayName = 'App Companion'; displayVersion = '9.0'; createdDateTime = (Get-Date) }
            )
        }
    }

    It "includes the anchor and excludes look-alike apps" {
        $all = Get-SameAppAllVersions 'App'
        $all.Count | Should -Be 3
        $all.id | Should -Not -Contain '4'
        $all.DisplayName | Should -Contain 'Ω DETECT - App'
    }

    It "sorts newest first" {
        (Get-SameAppAllVersions 'App')[0].displayVersion | Should -Be '3.0'
    }

    It "always returns an array even for a single match" {
        Mock -ModuleName YardstickSupport Invoke-WithRetry {
            @([PSCustomObject]@{ id = '1'; DisplayName = 'App'; displayVersion = '3.0'; createdDateTime = (Get-Date) })
        }
        $all = Get-SameAppAllVersions 'App'
        $all -is [array] | Should -BeTrue
        $all[0].displayVersion | Should -Be '3.0'
    }
}

Describe "Move-AssignmentsAndDependencies intent split" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppAssignment {
            @(
                [PSCustomObject]@{ id = 'a-req';   GroupID = 'g-req';   Intent = 'required';  FilterType = 'none' },
                [PSCustomObject]@{ id = 'a-avail'; GroupID = 'g-avail'; Intent = 'available'; FilterType = 'none' }
            )
        }
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency { @() }
        # A successful add returns the created assignment; that object is the
        # signal the module uses to skip the read-back.
        Mock -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup {
            [PSCustomObject]@{ id = 'new-assignment' }
        }
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppAssignmentGroup {}
    }

    It "IntentFilter='required' moves only required and removes from source" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -IntentFilter 'required'
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $Intent -eq 'required' }
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 0 -Exactly -ParameterFilter { $Intent -eq 'available' }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $GroupID -eq 'g-req' }
    }

    It "IntentFilter='required' with -SkipDependencies moves required and leaves dependencies alone" {
        # The Ω DETECT - anchor's required pass. Unlike an (N-x) version the anchor is
        # never deleted, so no dependent link needs rewriting - and its frozen child
        # set must not be merged onto the current app, where Add-IntuneWin32AppDependency
        # would resurrect it on every run.
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppDependency {
            @([PSCustomObject]@{ id = 'anchor_child1'; sourceId = 'anchor'; targetId = 'child1'; targetType = 'child'; dependencyType = 'detect' })
        }
        Mock -ModuleName YardstickSupport New-IntuneWin32AppDependency { [ordered]@{ targetId = $ID; dependencyType = $DependencyType } }
        Mock -ModuleName YardstickSupport Add-IntuneWin32AppDependency {}
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppDependency {}
        $from = [PSCustomObject]@{ id = 'anchor'; DisplayName = 'Ω DETECT - App' }
        $to   = [PSCustomObject]@{ id = 'to';     DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -IntentFilter 'required' -SkipDependencies
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $Intent -eq 'required' }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $ID -eq 'anchor' -and $GroupID -eq 'g-req' }
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppDependency -Times 0 -Exactly
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppDependency -Times 0 -Exactly
    }

    It "migrates by id, so a source whose rename has not propagated still works" {
        # $Anchor falls back to the raw pre-rename candidate when Intune's lookup is
        # still stale (Yardstick.ps1 step 1). The rename is a metadata patch and
        # assignments are keyed by id, so the lagging DisplayName must not matter.
        $from = [PSCustomObject]@{ id = 'anchor'; DisplayName = 'App' }  # not yet 'Ω DETECT - App'
        $to   = [PSCustomObject]@{ id = 'to';     DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -IntentFilter 'required' -SkipDependencies
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $ID -eq 'to' }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $ID -eq 'anchor' }
    }

    It "IntentFilter='available' with CopyOnly copies available and leaves source untouched" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -IntentFilter 'available' -CopyOnly -SkipDependencies
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $Intent -eq 'available' }
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 0 -Exactly -ParameterFilter { $Intent -eq 'required' }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppAssignmentGroup -Times 0 -Exactly
    }

    It "no filter processes every intent" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 2 -Exactly
    }

    It "keeps the source assignment when the target did not actually receive it" {
        # Add-IntuneWin32AppAssignmentGroup downgrades Graph failures to warnings
        # and returns nothing, so the target legitimately ends up without the
        # assignment. The read-back is the only thing that catches it.
        Mock -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup {}
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest { @() }
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppAssignment {
            if ($Id -eq 'to') { return @() }
            @([PSCustomObject]@{ id = 'a-req'; GroupID = 'g-req'; Intent = 'required'; FilterType = 'none' })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -IntentFilter 'required' -RetryDelaySeconds 0
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppAssignmentGroup -Times 0 -Exactly
        # An unverified add is retried rather than accepted on the first pass.
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 3 -Exactly
    }

    It "-CopyOnly:`$false falls back to a move" {
        # Yardstick passes -CopyOnly:$willBeSuperseded, so the false case must
        # behave exactly like the old move-everything behavior.
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -IntentFilter 'available' -CopyOnly:$false -SkipDependencies
        Should -Invoke -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $Intent -eq 'available' }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppAssignmentGroup -Times 1 -Exactly -ParameterFilter { $GroupID -eq 'g-avail' }
    }
}

Describe "Move-AssignmentsAndDependencies child dependencies" {
    BeforeEach {
        # Add-IntuneWin32AppDependency replaces an app's whole dependency set, so
        # the mock models that: the last write wins and becomes what a later Get
        # returns for the target app.
        $Global:MockToDependencies = @()
        $Global:MockAddDependencyCalls = @()

        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport { @() }
        Mock Get-IntuneWin32AppDependency -ModuleName YardstickSupport {
            switch ($ID) {
                'from' {
                    @(
                        [PSCustomObject]@{ id = 'from_child1'; sourceId = 'from'; targetId = 'child1'; targetType = 'child'; dependencyType = 'detect' },
                        [PSCustomObject]@{ id = 'from_child2'; sourceId = 'from'; targetId = 'child2'; targetType = 'child'; dependencyType = 'autoInstall' }
                    )
                }
                'to' { $Global:MockToDependencies }
                default { @() }
            }
        }
        Mock New-IntuneWin32AppDependency -ModuleName YardstickSupport {
            [ordered]@{ targetId = $ID; dependencyType = $DependencyType }
        }
        Mock Add-IntuneWin32AppDependency -ModuleName YardstickSupport {
            $Global:MockAddDependencyCalls += , @($Dependency)
            $Global:MockToDependencies = @($Dependency | ForEach-Object {
                [PSCustomObject]@{ targetId = $_.targetId; targetType = 'child'; dependencyType = $_.dependencyType }
            })
        }
        Mock Remove-IntuneWin32AppDependency -ModuleName YardstickSupport {}
    }

    AfterEach {
        Remove-Variable -Name MockToDependencies, MockAddDependencyCalls -Scope Global -ErrorAction SilentlyContinue
    }

    It "submits every child dependency in a single Add so none are clobbered" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to
        Should -Invoke Add-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 1
        $Global:MockAddDependencyCalls[0].targetId | Should -Be @('child1', 'child2')
    }

    It "never passes a relationship id where an app id is expected" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to
        # Graph relationship ids look like '<sourceId>_<targetId>' and Graph
        # rejects them with 'Invalid app id' when used as an app id.
        Should -Invoke Get-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 0 -ParameterFilter { $ID -like '*_*' }
        Should -Invoke New-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 0 -ParameterFilter { $ID -like '*_*' }
        Should -Invoke Add-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 0 -ParameterFilter { $ID -like '*_*' }
    }

    It "only ever writes dependencies to the target app" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to
        Should -Invoke Add-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 0 -ParameterFilter { $ID -ne 'to' }
        Should -Invoke Remove-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 0
    }

    It "keeps dependencies the target already has" {
        $Global:MockToDependencies = @(
            [PSCustomObject]@{ targetId = 'existing'; targetType = 'child'; dependencyType = 'detect' }
        )
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to
        $Global:MockAddDependencyCalls[0].targetId | Should -Be @('existing', 'child1', 'child2')
    }

    It "normalizes the dependency type to the values New-IntuneWin32AppDependency accepts" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to
        $Global:MockAddDependencyCalls[0].dependencyType | Should -Be @('Detect', 'AutoInstall')
    }

    It "never adds the source or target app as a dependency of the target" {
        Mock Get-IntuneWin32AppDependency -ModuleName YardstickSupport {
            switch ($ID) {
                'from' {
                    @(
                        [PSCustomObject]@{ id = 'from_to';     sourceId = 'from'; targetId = 'to';     targetType = 'child'; dependencyType = 'detect' },
                        [PSCustomObject]@{ id = 'from_child1'; sourceId = 'from'; targetId = 'child1'; targetType = 'child'; dependencyType = 'detect' }
                    )
                }
                'to' { $Global:MockToDependencies }
                default { @() }
            }
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to
        $Global:MockAddDependencyCalls[0].targetId | Should -Be @('child1')
    }

    It "does not carry a dependency onto the target when that app is being deleted this run" {
        # child2 is pruned later in the same run. Copying the dependency here
        # would create a fresh link that then blocks child2's own deletion.
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -ExcludeDependencyTargetIds @('child2')
        $Global:MockAddDependencyCalls[0].targetId | Should -Be @('child1')
    }

    It "drops a dependency the target already holds on an app being deleted this run" {
        # Add-IntuneWin32AppDependency replaces the whole set, so re-submitting an
        # existing entry is what keeps it alive - and it points at a doomed app.
        $Global:MockToDependencies = @(
            [PSCustomObject]@{ targetId = 'doomed'; targetType = 'child'; dependencyType = 'detect' }
        )
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -ExcludeDependencyTargetIds @('doomed')
        $Global:MockAddDependencyCalls[0].targetId | Should -Be @('child1', 'child2')
    }

    It "matches excluded dependency targets regardless of GUID casing" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -ExcludeDependencyTargetIds @('CHILD2')
        $Global:MockAddDependencyCalls[0].targetId | Should -Be @('child1')
    }

    It "ignores parent relationships when building the target's dependency set" {
        Mock Get-IntuneWin32AppDependency -ModuleName YardstickSupport {
            switch ($ID) {
                'from' {
                    @(
                        [PSCustomObject]@{ id = 'from_child1'; sourceId = 'from'; targetId = 'child1'; targetType = 'child';  dependencyType = 'detect' },
                        [PSCustomObject]@{ id = 'from_parent'; sourceId = 'from'; targetId = 'parent'; targetType = 'parent'; dependencyType = 'detect' }
                    )
                }
                'to' { $Global:MockToDependencies }
                default { @() }
            }
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -AllowDependentLinkUpdates $false
        $Global:MockAddDependencyCalls[0].targetId | Should -Be @('child1')
    }

    It "does not touch dependencies when SkipDependencies is set" {
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        Should -Invoke Add-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 0
        Should -Invoke Remove-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 0
    }

    It "retries when the target app does not report the dependencies back" {
        # Write never lands - Add-IntuneWin32AppDependency warns instead of
        # throwing when Graph rejects the update, so the read-back is the only
        # signal that anything went wrong.
        Mock Add-IntuneWin32AppDependency -ModuleName YardstickSupport {
            $Global:MockAddDependencyCalls += , @($Dependency)
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to
        Should -Invoke Add-IntuneWin32AppDependency -ModuleName YardstickSupport -Exactly -Times 3
    }
}

Describe "Get-YardstickAppAssignment" {
    It "drops the all-null assignment the cmdlet invents for an app with none" {
        # Get-IntuneWin32AppAssignment projects the empty OData envelope into one
        # assignment object with every property null, because a bare PSCustomObject
        # reports .Count 1 under PowerShell 7.
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = $null; AppName = $null; FilterID = $null; FilterType = $null
                GroupID = $null; GroupName = $null; Intent = $null; GroupMode = $null
                Notifications = $null; RestartSettings = $null; InstallTimeSettings = $null
            })
        }
        @(Get-YardstickAppAssignment -Id 'empty-app').Count | Should -Be 0
    }

    It "passes real assignments through untouched" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{ Type = '#microsoft.graph.groupAssignmentTarget'; GroupID = 'g1'; Intent = 'available' })
        }
        $result = @(Get-YardstickAppAssignment -Id 'real-app')
        $result.Count | Should -Be 1
        $result[0].GroupID | Should -Be 'g1'
    }
}

Describe "Move-AssignmentsAndDependencies assignment targets" {
    BeforeEach {
        Mock Get-IntuneWin32AppDependency -ModuleName YardstickSupport { @() }
        # A successful add returns the created assignment object.
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport { [PSCustomObject]@{ id = 'new-assignment' } }
        Mock Add-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport { [PSCustomObject]@{ id = 'new-assignment' } }
        Mock Add-IntuneWin32AppAssignmentAllUsers -ModuleName YardstickSupport { [PSCustomObject]@{ id = 'new-assignment' } }
        Mock Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {}
        Mock Remove-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport {}
        Mock Remove-IntuneWin32AppAssignmentAllUsers -ModuleName YardstickSupport {}
    }

    It "migrates an All Devices assignment instead of skipping it" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.allDevicesAssignmentTarget'
                GroupID = $null; GroupMode = $null; Intent = 'required'
                FilterType = 'none'; FilterID = $null; Notifications = 'hideAll'
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $ID -eq 'to' -and $Intent -eq 'required' -and $Notification -eq 'hideAll' }
        Should -Invoke Remove-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $ID -eq 'from' }
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 0
    }

    It "migrates an All Users assignment instead of skipping it" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.allLicensedUsersAssignmentTarget'
                GroupID = $null; GroupMode = $null; Intent = 'available'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentAllUsers -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $ID -eq 'to' -and $Intent -eq 'available' }
        Should -Invoke Remove-IntuneWin32AppAssignmentAllUsers -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $ID -eq 'from' }
    }

    It "leaves a filtered All Devices assignment alone rather than widening it to every device" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.allDevicesAssignmentTarget'
                GroupID = $null; GroupMode = $null; Intent = 'required'
                FilterType = 'include'; FilterID = 'filter-1'; Notifications = 'hideAll'
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport -Exactly -Times 0
        Should -Invoke Remove-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport -Exactly -Times 0
    }

    It "protects the source from deletion when it holds an assignment that cannot be migrated" {
        # The filter is only readable by id and only settable by name, so this
        # assignment cannot be recreated on $to. Pruning $from would destroy
        # targeting nobody can put back, so retention has to leave it alone.
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.allDevicesAssignmentTarget'
                GroupID = $null; GroupMode = $null; Intent = 'required'
                FilterType = 'include'; FilterID = 'filter-1'; Notifications = 'hideAll'
            })
        }
        $protected = [System.Collections.Generic.HashSet[string]]::new()
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -ProtectedSourceIds $protected
        $protected.Contains('from') | Should -BeTrue
    }

    It "does not protect the source when the assignment simply belongs to the other intent pass" {
        # The available pass will migrate this one; protecting here would stop
        # retention pruning an app that is perfectly safe to delete.
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.groupAssignmentTarget'
                GroupID = 'g-avail'; GroupMode = 'Include'; Intent = 'available'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
            })
        }
        $protected = [System.Collections.Generic.HashSet[string]]::new()
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -IntentFilter 'required' -ProtectedSourceIds $protected
        $protected.Count | Should -Be 0
    }

    It "does not protect the source when the cmdlet invents an assignment for an app that has none" {
        # The phantom used to fall through to the "unsupported target type" branch
        # and protect the source forever, so apps with no assignments at all could
        # never be pruned and piled up as (N-2)/(N-3) versions.
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = $null; GroupID = $null; GroupMode = $null; Intent = $null
                FilterType = $null; FilterID = $null; Notifications = $null
            })
        }
        $protected = [System.Collections.Generic.HashSet[string]]::new()
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-2)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -ProtectedSourceIds $protected
        $protected.Count | Should -Be 0
    }

    It "keeps a group exclusion an exclusion" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.exclusionGroupAssignmentTarget'
                GroupID = 'g-excl'; GroupMode = 'Exclude'; Intent = 'required'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        # Pester's mock enforces the real cmdlet's parameter sets, so a call that
        # binds -Exclude also proves no Include-only parameter (Notification,
        # install times, filters) was passed alongside it.
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $Exclude -eq $true -and $GroupID -eq 'g-excl' -and $Intent -eq 'required' }
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 0 -ParameterFilter { $Include -eq $true }
    }

    It "keeps a group inclusion an inclusion" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.groupAssignmentTarget'
                GroupID = 'g-incl'; GroupMode = 'Include'; Intent = 'required'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $Include -eq $true -and $GroupID -eq 'g-incl' }
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 0 -ParameterFilter { $Exclude -eq $true }
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $ID -eq 'from' -and $GroupID -eq 'g-incl' }
    }

    It "carries a group assignment's filter across" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.groupAssignmentTarget'
                GroupID = 'g-incl'; GroupMode = 'Include'; Intent = 'required'
                FilterType = 'exclude'; FilterID = 'filter-1'; Notifications = 'showAll'
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $FilterID -eq 'filter-1' -and $FilterMode -eq 'exclude' }
    }

    It "does not leak one assignment's install time settings onto the next" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @(
                [PSCustomObject]@{
                    Type = '#microsoft.graph.groupAssignmentTarget'
                    GroupID = 'g-timed'; GroupMode = 'Include'; Intent = 'required'
                    FilterType = 'include'; FilterID = 'filter-1'; Notifications = 'showAll'
                    InstallTimeSettings = [PSCustomObject]@{ useLocalTime = $true; startDateTime = [datetime]'2026-01-01 09:00'; deadlineDateTime = $null }
                },
                [PSCustomObject]@{
                    Type = '#microsoft.graph.groupAssignmentTarget'
                    GroupID = 'g-untimed'; GroupMode = 'Include'; Intent = 'required'
                    FilterType = 'include'; FilterID = 'filter-1'; Notifications = 'showAll'
                    InstallTimeSettings = $null
                }
            )
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter {
            $GroupID -eq 'g-timed' -and $AvailableTime.Hour -eq 9 -and $UseLocalTime -eq $true
        }
        # The second assignment has no install time settings of its own. The old
        # code never reset $startDateTime/$useLocalTime between iterations, so it
        # inherited the first assignment's schedule.
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter {
            $GroupID -eq 'g-untimed' -and
            (($null -eq $AvailableTime) -or ($AvailableTime -eq [datetime]::MinValue)) -and
            ($UseLocalTime -ne $true)
        }
    }

    It "still skips targets it has no way to migrate" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.configurationManagerCollectionAssignmentTarget'
                GroupID = $null; GroupMode = $null; Intent = 'required'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 0
        Should -Invoke Add-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport -Exactly -Times 0
        Should -Invoke Add-IntuneWin32AppAssignmentAllUsers -ModuleName YardstickSupport -Exactly -Times 0
    }

    It "leaves the source untouched under CopyOnly" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.allDevicesAssignmentTarget'
                GroupID = $null; GroupMode = $null; Intent = 'available'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -CopyOnly -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport -Exactly -Times 1
        Should -Invoke Remove-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport -Exactly -Times 0
    }
}

Describe "Move-AssignmentsAndDependencies install time settings" {
    BeforeEach {
        Mock Get-IntuneWin32AppDependency -ModuleName YardstickSupport { @() }
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport { [PSCustomObject]@{ id = 'new-assignment' } }
        Mock Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {}
    }

    It "rebases onto the offset date while keeping the source time of day" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.groupAssignmentTarget'
                GroupID = 'g-timed'; GroupMode = 'Include'; Intent = 'required'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
                InstallTimeSettings = [PSCustomObject]@{
                    useLocalTime = $true
                    startDateTime = [datetime]'2020-03-05 09:30'
                    deadlineDateTime = [datetime]'2020-03-05 17:45'
                }
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -AvailableDateOffset 2 -DeadlineDateOffset 4 -SkipDependencies
        $expectedAvailable = (Get-Date).Date.AddDays(2).AddHours(9).AddMinutes(30)
        $expectedDeadline  = (Get-Date).Date.AddDays(4).AddHours(17).AddMinutes(45)
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter {
            ($AvailableTime -eq $expectedAvailable) -and ($DeadlineTime -eq $expectedDeadline) -and ($UseLocalTime -eq $true)
        }
    }

    It "builds the rebased date without going through a culture-formatted string" {
        # A dd/MM/yyyy culture reads back "08/04/2026" as 8 April, and throws
        # outright once the day of the month passes the 12th.
        $originalCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('en-GB')
            Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
                @([PSCustomObject]@{
                    Type = '#microsoft.graph.groupAssignmentTarget'
                    GroupID = 'g-timed'; GroupMode = 'Include'; Intent = 'required'
                    FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
                    InstallTimeSettings = [PSCustomObject]@{
                        useLocalTime = $false
                        startDateTime = [datetime]'2020-03-05 09:30'
                        deadlineDateTime = $null
                    }
                })
            }
            $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
            $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
            Move-AssignmentsAndDependencies -From $from -To $to -AvailableDateOffset 0 -SkipDependencies
            $expectedAvailable = (Get-Date).Date.AddHours(9).AddMinutes(30)
            Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter {
                $AvailableTime -eq $expectedAvailable
            }
        }
        finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $originalCulture
        }
    }

    It "nudges a rebased deadline that has already passed into the future" {
        # A deadline-only assignment rebased onto today: an early-morning
        # deadline is already past for an afternoon run, and the cmdlet rejects
        # that combination with an uncatchable `break`.
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.groupAssignmentTarget'
                GroupID = 'g-timed'; GroupMode = 'Include'; Intent = 'required'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
                InstallTimeSettings = [PSCustomObject]@{
                    useLocalTime = $false
                    startDateTime = $null
                    deadlineDateTime = [datetime]'2020-03-05 00:01'
                }
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -DeadlineDateOffset 0 -SkipDependencies
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter {
            $DeadlineTime -gt (Get-Date)
        }
    }

    It "leaves a future deadline alone" {
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.groupAssignmentTarget'
                GroupID = 'g-timed'; GroupMode = 'Include'; Intent = 'required'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
                InstallTimeSettings = [PSCustomObject]@{
                    useLocalTime = $false
                    startDateTime = $null
                    deadlineDateTime = [datetime]'2020-03-05 23:59'
                }
            })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -DeadlineDateOffset 3 -SkipDependencies
        $expectedDeadline = (Get-Date).Date.AddDays(3).AddHours(23).AddMinutes(59)
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter {
            $DeadlineTime -eq $expectedDeadline
        }
    }
}

Describe "Test-YardstickAssignmentPresent" {
    BeforeEach {
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport {
            @(
                [PSCustomObject]@{ id = 'x1'; intent = 'required';  target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'g-1' } },
                [PSCustomObject]@{ id = 'x2'; intent = 'available'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.allDevicesAssignmentTarget' } }
            )
        }
    }

    It "matches a group assignment by group id and intent" {
        Test-YardstickAssignmentPresent -AppId 'to' -CountKey 'g-1' -Intent 'required' | Should -BeTrue
    }

    It "does not match when the intent differs" {
        Test-YardstickAssignmentPresent -AppId 'to' -CountKey 'g-1' -Intent 'available' | Should -BeFalse
    }

    It "falls back to the target type for virtual targets that carry no group id" {
        Test-YardstickAssignmentPresent -AppId 'to' -CountKey '#microsoft.graph.allDevicesAssignmentTarget' -Intent 'available' | Should -BeTrue
    }

    It "does not mistake an exclusion for the include it was asked about" {
        # Both carry the same group id, so CountKey alone cannot tell them apart -
        # and treating one as the other would delete the wrong source assignment.
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport {
            @([PSCustomObject]@{ id = 'x1'; intent = 'required'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget'; groupId = 'g-1' } })
        }
        Test-YardstickAssignmentPresent -AppId 'to' -CountKey 'g-1' -Intent 'required' -TargetType '#microsoft.graph.groupAssignmentTarget' | Should -BeFalse
    }

    It "still matches when neither side reports a target type" {
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport {
            @([PSCustomObject]@{ id = 'x1'; intent = 'required'; target = [PSCustomObject]@{ groupId = 'g-1' } })
        }
        Test-YardstickAssignmentPresent -AppId 'to' -CountKey 'g-1' -Intent 'required' | Should -BeTrue
    }

    It "reports absent rather than throwing when Graph fails" {
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport { throw "boom" }
        Test-YardstickAssignmentPresent -AppId 'to' -CountKey 'g-1' -Intent 'required' | Should -BeFalse
    }

    It "finds an assignment on an app that holds exactly one" {
        # Get-IntuneWin32AppAssignment returns $null in this case under Windows
        # PowerShell 5.1, which is why this reads through Graph directly.
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport {
            @([PSCustomObject]@{ id = 'x1'; intent = 'required'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'g-1' } })
        }
        Test-YardstickAssignmentPresent -AppId 'to' -CountKey 'g-1' -Intent 'required' | Should -BeTrue
    }
}

Describe "Move-AssignmentsAndDependencies assignment error handling" {
    BeforeEach {
        Mock Get-IntuneWin32AppDependency -ModuleName YardstickSupport { @() }
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.groupAssignmentTarget'
                GroupID = 'g-1'; GroupMode = 'Include'; Intent = 'required'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
            })
        }
        Mock Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {}
        # Target reports the assignment as present unless a test overrides this.
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport {
            @([PSCustomObject]@{ id = 'x1'; intent = 'required'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'g-1' } })
        }
    }

    It "records the Intune warning in the Yardstick log instead of only the console" {
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            Write-Warning "An error occurred while creating a Win32 app assignment. Error message: BadRequest: something went wrong"
        }
        Mock Write-Log -ModuleName YardstickSupport {}
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0
        Should -Invoke Write-Log -ModuleName YardstickSupport -ParameterFilter {
            $Content -match 'WARNING from Intune' -and $Content -match 'BadRequest'
        }
    }

    It "treats an 'already exists' conflict as success when the assignment really is there" {
        # The exact production case: Intune rejects the duplicate, but the
        # assignment we wanted is already on the target, so this is a no-op.
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            Write-Warning "An error occurred while creating a Win32 app assignment. Error message: BadRequest: The MobileApp Assignment already exists"
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 1 -Exactly
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 1 -Exactly
    }

    It "never strips the source under -CopyOnly, even when Intune says the target already has it" {
        # Run N+1 against the Ω DETECT - anchor: the anchor still holds the available
        # assignment it kept last run, a kept (N-x) version already copied the same
        # group onto the new app, and Intune rejects the duplicate. That conflict is
        # a success - but a copy must still never delete what it copied, because
        # stripping an available assignment destroys the on-device auto-update
        # component. This is the assertion that the anchor survives repeat runs.
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @([PSCustomObject]@{
                Type = '#microsoft.graph.groupAssignmentTarget'
                GroupID = 'g-1'; GroupMode = 'Include'; Intent = 'available'
                FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
            })
        }
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport {
            @([PSCustomObject]@{ id = 'x1'; intent = 'available'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'g-1' } })
        }
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            Write-Warning "An error occurred while creating a Win32 app assignment. Error message: BadRequest: The MobileApp Assignment already exists"
        }
        $from = [PSCustomObject]@{ id = 'anchor'; DisplayName = 'Ω DETECT - App' }
        $to   = [PSCustomObject]@{ id = 'to';     DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -IntentFilter 'available' -CopyOnly -SkipDependencies -RetryDelaySeconds 0
        # A conflict is not retried into a loop, and the copy is not undone.
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 1 -Exactly
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 0 -Exactly
    }

    It "does not burn retries on a conflict a retry cannot resolve" {
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            Write-Warning "An error occurred while creating a Win32 app assignment. Error message: BadRequest: The MobileApp Assignment already exists"
        }
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport { @() }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 1 -Exactly
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 0 -Exactly
    }

    It "retries a transient failure and keeps the source when it never lands" {
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            Write-Warning "An error occurred while creating a Win32 app assignment. Error message: 503 Service Unavailable"
        }
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport { @() }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0
        Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 3 -Exactly
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 0 -Exactly
    }

    It "protects the source from deletion when the assignment never lands on the target" {
        # The source copy is now the only one, so retention must not prune it.
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            Write-Warning "An error occurred while creating a Win32 app assignment. Error message: 503 Service Unavailable"
        }
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport { @() }
        $protected = [System.Collections.Generic.HashSet[string]]::new()
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0 -ProtectedSourceIds $protected
        $protected.Contains('from') | Should -BeTrue
    }

    It "recovers when a retry succeeds after a transient failure" {
        # Mocks run in the module's scope, so the counter has to be global - the
        # same convention the break-containment test uses.
        $Global:MockAddAttempts = 0
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            $Global:MockAddAttempts++
            if ($Global:MockAddAttempts -eq 1) {
                Write-Warning "An error occurred while creating a Win32 app assignment. Error message: 429 Too Many Requests"
                return
            }
            [PSCustomObject]@{ id = 'new-assignment' }
        }
        # Read-back never finds it, so the first attempt's warning is a real
        # failure; the second attempt succeeds by returning the created object.
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport { @() }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        try {
            Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0
            Should -Invoke Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 2 -Exactly
            Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 1 -Exactly
        }
        finally {
            Remove-Variable -Name MockAddAttempts -Scope Global -ErrorAction SilentlyContinue
        }
    }

    It "trusts a returned assignment object without reading it back" {
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport { [PSCustomObject]@{ id = 'new-assignment' } }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0
        Should -Invoke Invoke-YardstickGraphRequest -ModuleName YardstickSupport -Times 0 -Exactly
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 1 -Exactly
    }

    It "never logs a successful add when Intune rejected it" {
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            Write-Warning "An error occurred while creating a Win32 app assignment. Error message: BadRequest: nope"
        }
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport { @() }
        Mock Write-Log -ModuleName YardstickSupport {}
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0
        Should -Invoke Write-Log -ModuleName YardstickSupport -Times 0 -Exactly -ParameterFilter { $Content -match '^Added group' }
        Should -Invoke Write-Log -ModuleName YardstickSupport -ParameterFilter { $Content -match 'leaving the source assignment' }
    }

    It "leaves the source alone when the add landed on a different target type" {
        # Target holds an exclusion for the same group; the include we tried to
        # add is not there, so the source copy must survive.
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {
            Write-Warning "An error occurred while creating a Win32 app assignment. Error message: BadRequest: nope"
        }
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport {
            @([PSCustomObject]@{ id = 'x1'; intent = 'required'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.exclusionGroupAssignmentTarget'; groupId = 'g-1' } })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies -RetryDelaySeconds 0
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Times 0 -Exactly
    }
}

Describe "Move-AssignmentsAndDependencies break containment" {
    It "keeps migrating later assignments when a cmdlet bails out with break" {
        Mock Get-IntuneWin32AppDependency -ModuleName YardstickSupport { @() }
        Mock Get-IntuneWin32AppAssignment -ModuleName YardstickSupport {
            @(
                [PSCustomObject]@{
                    Type = '#microsoft.graph.groupAssignmentTarget'
                    GroupID = 'g-bails'; GroupMode = 'Include'; Intent = 'required'
                    FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
                },
                [PSCustomObject]@{
                    Type = '#microsoft.graph.groupAssignmentTarget'
                    GroupID = 'g-after'; GroupMode = 'Include'; Intent = 'required'
                    FilterType = 'none'; FilterID = $null; Notifications = 'showAll'
                }
            )
        }
        Mock Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {}
        # Both groups report as present on the target. g-bails must STILL not be
        # removed: its cmdlet bailed out with `break` and never returned, and a
        # bail-out is gated out of the read-back precisely so a pre-existing
        # assignment cannot be mistaken for one this run just created.
        Mock Invoke-YardstickGraphRequest -ModuleName YardstickSupport {
            @(
                [PSCustomObject]@{ id = 'x1'; intent = 'required'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'g-bails' } },
                [PSCustomObject]@{ id = 'x2'; intent = 'required'; target = [PSCustomObject]@{ '@odata.type' = '#microsoft.graph.groupAssignmentTarget'; groupId = 'g-after' } }
            )
        }

        $Global:MockBreakAdds = @()
        try {
            # Pester's mocks absorb `break`, so shadow the cmdlet with a real
            # function to reproduce what the IntuneWin32App cmdlets actually do
            # on a past deadline or an expired token. It has to be global rather
            # than InModuleScope - a function defined inside an InModuleScope
            # block dies with that block - and a function outranks a cmdlet in
            # PowerShell's command resolution, so the module picks this up.
            function global:Add-IntuneWin32AppAssignmentGroup {
                [CmdletBinding()]
                param([switch]$Include, [switch]$Exclude, $ID, $GroupID, $Intent,
                      $Notification, $AvailableTime, $DeadlineTime, [bool]$UseLocalTime,
                      $FilterMode, $FilterID)
                begin {
                    $Global:MockBreakAdds += $GroupID
                    if ($GroupID -eq 'g-bails') { break }
                }
            }
            $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
            $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
            Move-AssignmentsAndDependencies -From $from -To $to -SkipDependencies
        }
        finally {
            Remove-Item function:global:Add-IntuneWin32AppAssignmentGroup -ErrorAction SilentlyContinue
        }

        # The assignment that bailed is retried, then the loop carries on to the
        # next one instead of silently abandoning it.
        @($Global:MockBreakAdds | Where-Object { $_ -eq 'g-bails' }).Count | Should -Be 3
        $Global:MockBreakAdds | Should -Contain 'g-after'
        # A bail-out is not success, so the source assignment must survive.
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 0 -ParameterFilter { $GroupID -eq 'g-bails' }
        Should -Invoke Remove-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport -Exactly -Times 1 -ParameterFilter { $GroupID -eq 'g-after' }

        Remove-Variable -Name MockBreakAdds -Scope Global -ErrorAction SilentlyContinue
    }
}

