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
        Get-DetectAnchorName -DisplayName 'Google Chrome' | Should -Be '{DETECT} Google Chrome'
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
        $targets += [PSCustomObject]@{ id = 'anchor'; DisplayName = '{DETECT} App'; displayVersion = '0.1' }
        $count = Set-YardstickSupersedence -NewApp $new -SupersededApps $targets -Type 'Replace' -UpdateOnlyIds @('anchor')
        $count | Should -Be 9
        Should -Invoke -ModuleName YardstickSupport New-IntuneWin32AppSupersedence -Times 1 -Exactly -ParameterFilter { $ID -eq 'anchor' -and $SupersedenceType -eq 'Update' }
    }

    It "forces Update for ids listed in UpdateOnlyIds even when Type is Replace" {
        $new = [PSCustomObject]@{ id = 'new'; DisplayName = 'App'; displayVersion = '3.0' }
        $targets = @(
            [PSCustomObject]@{ id = 'old';    DisplayName = 'App (N-1)';    displayVersion = '2.0' },
            [PSCustomObject]@{ id = 'anchor'; DisplayName = '{DETECT} App'; displayVersion = '1.0' }
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

Describe "Remove-YardstickApp" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Remove-IntuneWin32App {}
        Mock -ModuleName YardstickSupport Remove-IntuneWin32AppSupersedence {}
        Mock -ModuleName YardstickSupport Remove-SupersedenceReference {}
        # Deletion succeeded: the app is no longer resolvable.
        Mock -ModuleName YardstickSupport Get-IntuneWin32App { $null }
    }

    It "detaches superseding parents before deleting" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
            @([PSCustomObject]@{ sourceId = 'parent'; targetId = 'doomed' })
        }
        Remove-YardstickApp -App ([PSCustomObject]@{ id = 'doomed'; DisplayName = 'App (N-3)' })
        Should -Invoke -ModuleName YardstickSupport Remove-SupersedenceReference -Times 1 -Exactly -ParameterFilter {
            $ParentId -eq 'parent' -and $TargetId -eq 'doomed'
        }
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32App -Times 1 -Exactly -ParameterFilter { $Id -eq 'doomed' }
    }

    It "clears its own forward links before deleting" {
        Mock -ModuleName YardstickSupport Get-YardstickSupersedenceRelationship {
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

    It "returns an empty array when Graph fails" {
        Mock -ModuleName YardstickSupport Invoke-YardstickGraphRequest { throw "boom" }
        (Get-YardstickSupersedenceRelationship -Id 'me').Count | Should -Be 0
    }
}

Describe "Get-SameAppAllVersions" {
    BeforeEach {
        Mock -ModuleName YardstickSupport Invoke-WithRetry {
            @(
                [PSCustomObject]@{ id = '1'; DisplayName = 'App';           displayVersion = '3.0'; createdDateTime = (Get-Date) },
                [PSCustomObject]@{ id = '2'; DisplayName = 'App (N-1)';     displayVersion = '2.0'; createdDateTime = (Get-Date) },
                [PSCustomObject]@{ id = '3'; DisplayName = '{DETECT} App';  displayVersion = '1.0'; createdDateTime = (Get-Date) },
                [PSCustomObject]@{ id = '4'; DisplayName = 'App Companion'; displayVersion = '9.0'; createdDateTime = (Get-Date) }
            )
        }
    }

    It "includes the anchor and excludes look-alike apps" {
        $all = Get-SameAppAllVersions 'App'
        $all.Count | Should -Be 3
        $all.id | Should -Not -Contain '4'
        $all.DisplayName | Should -Contain '{DETECT} App'
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
        Mock -ModuleName YardstickSupport Add-IntuneWin32AppAssignmentGroup {}
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
        # Add-IntuneWin32AppAssignmentGroup downgrades Graph failures to warnings,
        # so the target legitimately ends up without the assignment.
        Mock -ModuleName YardstickSupport Get-IntuneWin32AppAssignment {
            if ($Id -eq 'to') { return @() }
            @([PSCustomObject]@{ id = 'a-req'; GroupID = 'g-req'; Intent = 'required'; FilterType = 'none' })
        }
        $from = [PSCustomObject]@{ id = 'from'; DisplayName = 'App (N-1)' }
        $to   = [PSCustomObject]@{ id = 'to';   DisplayName = 'App' }
        Move-AssignmentsAndDependencies -From $from -To $to -IntentFilter 'required'
        Should -Invoke -ModuleName YardstickSupport Remove-IntuneWin32AppAssignmentGroup -Times 0 -Exactly
    }
}
