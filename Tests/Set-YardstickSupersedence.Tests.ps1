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

Describe "Move-AssignmentsAndDependencies assignment targets" {
    BeforeEach {
        Mock Get-IntuneWin32AppDependency -ModuleName YardstickSupport { @() }
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {}
        Mock Add-IntuneWin32AppAssignmentAllDevices -ModuleName YardstickSupport {}
        Mock Add-IntuneWin32AppAssignmentAllUsers -ModuleName YardstickSupport {}
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
        Mock Add-IntuneWin32AppAssignmentGroup -ModuleName YardstickSupport {}
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

