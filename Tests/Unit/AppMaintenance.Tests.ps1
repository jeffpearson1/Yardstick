BeforeAll {
    # Invoke-YardstickAppMaintenance and Invoke-YardstickSkip live in
    # Invoke-YardstickRun.ps1, which cannot be dot-sourced (its top level runs a
    # whole update pass), so extract them via the AST. They read the runner's
    # $Script:* recipe settings and per-recipe state, which each test sets.
    $runnerPath = "$PSScriptRoot\..\..\Private\Invoke-YardstickRun.ps1"
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($runnerPath, [ref]$null, [ref]$null)
    # The functions sort with [VersionPro]. Define the class in this scope too, so
    # the extracted functions resolve it however Pester scopes this file.
    $versionPro = (Get-Content "$PSScriptRoot\..\..\Modules\VersionPro.psm1" -Raw) -replace '(?m)^\s*Export-ModuleMember.*$', ''
    . ([scriptblock]::Create($versionPro))
    foreach ($name in 'Invoke-YardstickAppMaintenance', 'Invoke-YardstickSkip', 'Invoke-YardstickInterruptedUploadHeal') {
        $fn = $ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $fn) { throw "$name not found in $runnerPath" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    # Stand-ins for the module commands they call, for Mock to replace.
    function Write-Log { param([string]$Content, [string]$Level) }
    function Get-YardstickWin32App { [CmdletBinding()] param([string]$Id, [string]$DisplayName, [switch]$AllowMissing) }
    function Get-SameAppAllVersions { [CmdletBinding()] param([string]$DisplayName, [int]$MinGenerations, [string]$StopWhenVersion, [switch]$NoAbsenceProof) }
    function Get-DetectAnchorName { param([string]$DisplayName) "Ω DETECT - $DisplayName" }
    function Get-DetectAnchor { param([string]$DisplayName) }
    function Test-IsVersionDetection { param($DetectionType, $FileDetectionMethod, $RegistryDetectionMethod) $true }
    function Set-DetectAnchor { param($App, $DisplayName) }
    function Set-YardstickWin32App { param($Id, $DisplayName) }
    function Move-AssignmentsAndDependencies { param($From, $To, $AvailableDateOffset, $DeadlineDateOffset, $IntentFilter, $AllowDependentLinkUpdates, $DependentLinkOptions, $DependentUpdateStatus, $ProtectedSourceIds, $ExcludeDependencyTargetIds, [switch]$SkipDependencies, [switch]$CopyOnly) }
    function Get-YardstickAppAssignment { param($Id) }
    function Add-YardstickWin32AppAssignmentGroup { param([switch]$Include, $ID, $GroupID, $Intent, $Notification) }
    function Remove-YardstickApp { param($App) }
    function Set-YardstickSupersedence { param($NewApp, $SupersededApps, $Type, $UpdateOnlyIds) @($SupersededApps).Count }
    function Set-AssignmentAutoUpdate { param($AppId, $Enabled, $IntentFilter, $SkipGroupIds) 0 }
    function Add-SuccessfulApplication { param($ApplicationId, $DisplayName, $Version, $Action, $Dependents, $AutoUpdateStatus) }
    function Add-FailedApplication { param($ApplicationId, $DisplayName, $Version, $ErrorMessage, $FailureStage) }
    function Set-YardstickWatchdogStage { param($AppId, $Stage) }
    $supportPath = "$PSScriptRoot\..\..\Modules\YardstickSupport.psm1"
    $supportAst = [System.Management.Automation.Language.Parser]::ParseFile($supportPath, [ref]$null, [ref]$null)
    $fn = $supportAst.Find({ param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-YardstickInterruptedUpload' }, $true)
    . ([scriptblock]::Create($fn.Extent.Text))
    function Test-YardstickAppPublished {
        param($App)
        if ($null -eq $App) { return $false }
        $state = [string]$App.publishingState
        return (-not $state) -or ($state -eq 'published')
    }

    function New-App {
        param([string]$Id, [string]$Name, [string]$Version, [string]$State = 'published')
        [pscustomobject]@{ id = $Id; displayName = $Name; displayVersion = $Version; publishingState = $State; createdDateTime = (Get-Date) }
    }
}

Describe 'Invoke-YardstickAppMaintenance after an upload' {
    BeforeEach {
        $Script:DisplayName = 'calibre'
        $Script:NumVersionsToKeep = 2
        $Script:UseDetectAnchor = $true
        $Script:Supersedence = $true
        $Script:UninstallPreviousVersion = $false
        $Script:AutoUpdateOnAssignment = $false
        $Script:DefaultDeploymentGroups = @()
        $AppId_Processing = 'calibre'
        $NoDelete = $false
        $ProtectedDependencyAppIds = [System.Collections.Generic.HashSet[string]]::new()
        $DependentLinkOptions = @{}
        $DependentUpdateStatus = [ordered]@{}

        # calibre on 2026-11-23, before the upload.
        $script:family = @(
            (New-App '9d828d0f' 'calibre' '9.18.0'),
            (New-App 'd28e7fe0' 'calibre (N-1)' '9.17.0'),
            (New-App '34d9517f' 'Ω DETECT - calibre' '9.14.0')
        )
        $script:newApp = New-App '5366c25d' 'calibre' '9.19.0'
        $script:byId = @{}
        foreach ($a in $script:family) { $script:byId[$a.id] = $a }

        Mock Get-YardstickWin32App { $script:byId[$Id] }
        Mock Get-SameAppAllVersions { throw 'no name lookup after an upload' }
        Mock Start-Sleep {}
        Mock Get-YardstickAppAssignment { @() }
        Mock Set-YardstickWin32App {}
        Mock Remove-YardstickApp {}
        Mock Set-YardstickSupersedence { @($SupersededApps).Count }
        Mock Write-Log {}
    }

    It 'finishes calibre from the family it already has, with no name lookup' {
        $result = Invoke-YardstickAppMaintenance -NewAppId '5366c25d' -NewApp $script:newApp -ExistingApps $script:family
        Should -Invoke Get-SameAppAllVersions -Times 0 -Exactly
        Should -Invoke Start-Sleep -Times 0 -Exactly
        $result.CurrentApp.id | Should -Be '5366c25d'
        Should -Invoke Set-YardstickWin32App -Times 1 -Exactly -ParameterFilter { $Id -eq '9d828d0f' -and $DisplayName -eq 'calibre (N-1)' }
        Should -Invoke Remove-YardstickApp -Times 1 -Exactly -ParameterFilter { $App.id -eq 'd28e7fe0' }
        Should -Invoke Set-YardstickSupersedence -Times 1 -Exactly -ParameterFilter {
            $NewApp.id -eq '5366c25d' -and
            (@($SupersededApps.id | Sort-Object) -join ',') -eq '34d9517f,9d828d0f' -and
            (@($UpdateOnlyIds) -join ',') -eq '34d9517f'
        }
        @($result.StepErrors).Count | Should -Be 0
    }

    It 'reads the new app by id when the caller does not pass it' {
        $script:byId['5366c25d'] = $script:newApp
        $result = Invoke-YardstickAppMaintenance -NewAppId '5366c25d' -ExistingApps $script:family
        $result.CurrentApp.id | Should -Be '5366c25d'
    }

    It 'throws with the -Repair command when the new app cannot be read back' {
        { Invoke-YardstickAppMaintenance -NewAppId 'missing' -ExistingApps $script:family } | Should -Throw '*-AppId calibre -Repair*'
    }

    It 'works for a first upload with no existing versions' {
        $result = Invoke-YardstickAppMaintenance -NewAppId '5366c25d' -NewApp $script:newApp -ExistingApps @()
        $result.CurrentApp.id | Should -Be '5366c25d'
        Should -Invoke Get-SameAppAllVersions -Times 0 -Exactly
    }

    It 'drops a member deleted since the lookup and still supersedes the rest' {
        $script:byId.Remove('d28e7fe0')
        $null = Invoke-YardstickAppMaintenance -NewAppId '5366c25d' -NewApp $script:newApp -ExistingApps $script:family
        Should -Invoke Remove-YardstickApp -Times 0 -Exactly
        Should -Invoke Set-YardstickSupersedence -Times 1 -Exactly -ParameterFilter { (@($SupersededApps.id | Sort-Object) -join ',') -eq '34d9517f,9d828d0f' }
    }

    It 'keeps the copy it already has when a member cannot be re-read' {
        Mock Get-YardstickWin32App { if ($Id -eq '9d828d0f') { throw '504' }; $script:byId[$Id] }
        $null = Invoke-YardstickAppMaintenance -NewAppId '5366c25d' -NewApp $script:newApp -ExistingApps $script:family
        Should -Invoke Set-YardstickWin32App -Times 1 -Exactly -ParameterFilter { $Id -eq '9d828d0f' }
    }

    It 'pins a new anchor without looking it up by name' {
        $family = @((New-App '9d828d0f' 'calibre' '9.18.0'), (New-App 'd28e7fe0' 'calibre (N-1)' '9.17.0'))
        Mock Set-DetectAnchor {}
        Mock Get-DetectAnchor { throw 'no name lookup' }
        $result = Invoke-YardstickAppMaintenance -NewAppId '5366c25d' -NewApp $script:newApp -ExistingApps $family
        Should -Invoke Set-DetectAnchor -Times 1 -Exactly -ParameterFilter { $App.id -eq 'd28e7fe0' }
        Should -Invoke Get-DetectAnchor -Times 0 -Exactly
        $result.AnchorStatus | Should -BeLike 'pinned*'
        # The pinned anchor must not be renamed back to (N-x).
        Should -Invoke Set-YardstickWin32App -Times 0 -Exactly -ParameterFilter { $Id -eq 'd28e7fe0' }
    }

    It 'carries on and reports the step when an old version''s assignments cannot be read' {
        Mock Get-YardstickAppAssignment { throw 'Graph 504' }
        $result = Invoke-YardstickAppMaintenance -NewAppId '5366c25d' -NewApp $script:newApp -ExistingApps $script:family
        $result.StepErrors | Should -Contain 'available assignments'
        Should -Invoke Set-YardstickSupersedence -Times 1 -Exactly
    }

    It 'does not delete either of two older apps at the same version' {
        $Script:UseDetectAnchor = $false
        $family = @(
            (New-App '9d828d0f' 'calibre' '9.18.0'),
            (New-App 'aaaa' 'calibre (N-1)' '9.17.0'),
            (New-App 'bbbb' 'calibre' '9.17.0')
        )
        foreach ($a in $family) { $script:byId[$a.id] = $a }
        $null = Invoke-YardstickAppMaintenance -NewAppId '5366c25d' -NewApp $script:newApp -ExistingApps $family
        Should -Invoke Remove-YardstickApp -Times 0 -Exactly -ParameterFilter { $App.id -in 'aaaa', 'bbbb' }
    }
}

Describe 'Invoke-YardstickAppMaintenance sweep' {
    BeforeEach {
        $Script:DisplayName = 'App'
        $Script:NumVersionsToKeep = 2
        $Script:UseDetectAnchor = $false
        $Script:Supersedence = $true
        $Script:UninstallPreviousVersion = $false
        $Script:AutoUpdateOnAssignment = $false
        $Script:DefaultDeploymentGroups = @()
        $NoDelete = $false
        $ProtectedDependencyAppIds = [System.Collections.Generic.HashSet[string]]::new()
        Mock Get-SameAppAllVersions { throw 'should use ExistingApps' }
        Mock Get-YardstickAppAssignment { @() }
        Mock Set-YardstickWin32App {}
        Mock Remove-YardstickApp {}
        Mock Set-YardstickSupersedence { @($SupersededApps).Count }
        Mock Write-Log {}
    }

    It 'uses the family it is given instead of reading Intune again' {
        $family = @((New-App '2' 'App' '2.0'), (New-App '1' 'App (N-1)' '1.0'))
        $result = Invoke-YardstickAppMaintenance -ExistingApps $family
        $result.CurrentApp.id | Should -Be '2'
        Should -Invoke Get-SameAppAllVersions -Times 0 -Exactly
    }

    It 'never makes an unpublished app current' {
        $family = @((New-App 'ghost' 'App' '3.0' 'notPublished'), (New-App '2' 'App' '2.0'))
        $result = Invoke-YardstickAppMaintenance -ExistingApps $family
        $result.CurrentApp.id | Should -Be '2'
        Should -Invoke Remove-YardstickApp -Times 0 -Exactly -ParameterFilter { $App.id -eq 'ghost' }
        Should -Invoke Set-YardstickSupersedence -Times 1 -Exactly -ParameterFilter { @($SupersededApps.id) -notcontains 'ghost' }
    }
}

Describe 'Invoke-YardstickSkip' {
    BeforeEach {
        $Script:DisplayName = 'App'
        $AppId_Processing = 'app'
        $CurrentDisplayName = 'App'
        $DependentUpdateStatus = [ordered]@{}
        Mock Write-Log {}
        Mock Invoke-YardstickAppMaintenance { [pscustomobject]@{ ChangedSomething = $false } }
    }

    It 'only logs the reason without -Repair' {
        $Repair = $false
        Invoke-YardstickSkip -Reason 'Nothing new.'
        Should -Invoke Invoke-YardstickAppMaintenance -Times 0 -Exactly
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Content -eq 'Nothing new. Skipping update.' }
    }

    It 'passes the family it was given to the -Repair sweep' {
        $Repair = $true
        $family = @((New-App '1' 'App' '1.0'))
        Invoke-YardstickSkip -Reason 'Nothing new.' -ExistingApps $family
        Should -Invoke Invoke-YardstickAppMaintenance -Times 1 -Exactly -ParameterFilter { @($ExistingApps).Count -eq 1 }
    }

    It 'lets the -Repair sweep read Intune when given no family' {
        $Repair = $true
        Invoke-YardstickSkip -Reason 'Nothing new.'
        Should -Invoke Invoke-YardstickAppMaintenance -Times 1 -Exactly -ParameterFilter { $null -eq $ExistingApps }
    }
}

Describe 'Invoke-YardstickInterruptedUploadHeal' {
    BeforeEach {
        $Script:DisplayName = 'calibre'
        $Script:Version = '9.19.0'
        $Script:NumVersionsToKeep = 2
        $Script:AutoHealInterruptedUploads = $true
        $AppId_Processing = 'calibre'
        $CurrentDisplayName = 'calibre'
        $DependentUpdateStatus = [ordered]@{}

        $script:newest = New-App '5366c25d' 'calibre' '9.19.0'
        $script:newest | Add-Member -NotePropertyName isAssigned -NotePropertyValue $false
        $script:older = New-App '9d828d0f' 'calibre' '9.18.0'
        $script:older | Add-Member -NotePropertyName isAssigned -NotePropertyValue $true
        $script:bare = @($script:newest, $script:older)

        Mock Write-Log {}
        Mock Get-SameAppAllVersions { @($script:newest, $script:older) }
        Mock Invoke-YardstickAppMaintenance {
            [pscustomobject]@{ CurrentApp = $script:newest; SupersededCount = 2; PrunedCount = 1; RenamedCount = 1; AnchorStatus = $null; StepErrors = @() }
        }
        Mock Add-SuccessfulApplication {}
        Mock Add-FailedApplication {}
    }

    It 'finishes calibre on a normal run and reports it as an update' {
        Invoke-YardstickInterruptedUploadHeal -Reason 'calibre 9.19.0 is already in Intune.' -ExistingApps $script:bare | Should -BeTrue
        Should -Invoke Get-SameAppAllVersions -Times 1 -Exactly
        Should -Invoke Invoke-YardstickAppMaintenance -Times 1 -Exactly -ParameterFilter { @($ExistingApps).Count -eq 2 }
        Should -Invoke Add-SuccessfulApplication -Times 1 -Exactly -ParameterFilter {
            $Action -eq 'Updated' -and $Version -eq '9.19.0' -and $AutoUpdateStatus -like 'finished an interrupted upload*'
        }
        # The log report keys a sweep off this wording.
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Content -like '*an earlier update did not finish. Running maintenance sweep.' }
    }

    It 'reports Maintained when the newer app already had assignments but is Yardstick''s' {
        $script:newest.isAssigned = $true
        $script:newest | Add-Member -NotePropertyName notes -NotePropertyValue 'Created by Yardstick, upload 2026-11-24T14:00:00Z 1a2b3c4d'
        Invoke-YardstickInterruptedUploadHeal -Reason 'r' -ExistingApps $script:bare | Should -BeTrue
        Should -Invoke Add-SuccessfulApplication -Times 1 -Exactly -ParameterFilter { $Action -eq 'Maintained' }
    }

    It 'leaves an assigned app that Yardstick did not create to a person' {
        $script:newest.isAssigned = $true
        Invoke-YardstickInterruptedUploadHeal -Reason 'r' -ExistingApps $script:bare | Should -BeFalse
        Should -Invoke Invoke-YardstickAppMaintenance -Times 0 -Exactly
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Level -eq 'Warning' -and $Content -like '*-Repair*' }
    }

    It 'does nothing when the newer app is not the version found upstream' {
        $Script:Version = '9.18.0'
        Invoke-YardstickInterruptedUploadHeal -Reason 'r' -ExistingApps $script:bare | Should -BeFalse
        Should -Invoke Invoke-YardstickAppMaintenance -Times 0 -Exactly
    }

    It 'does nothing when autoHealInterruptedUploads is off' {
        $Script:AutoHealInterruptedUploads = $false
        Invoke-YardstickInterruptedUploadHeal -Reason 'r' -ExistingApps $script:bare | Should -BeFalse
        Should -Invoke Invoke-YardstickAppMaintenance -Times 0 -Exactly
    }

    It 'does nothing for a healthy family' {
        Invoke-YardstickInterruptedUploadHeal -Reason 'r' -ExistingApps @($script:newest) | Should -BeFalse
        Should -Invoke Write-Log -Times 0 -Exactly
    }

    It 'reports two apps at the same version for a person to resolve, without a sweep' {
        $twin = New-App 'f087d401' 'calibre' '9.19.0'
        Invoke-YardstickInterruptedUploadHeal -Reason 'r' -ExistingApps @($script:newest, $twin) | Should -BeTrue
        Should -Invoke Invoke-YardstickAppMaintenance -Times 0 -Exactly
        Should -Invoke Add-FailedApplication -Times 1 -Exactly -ParameterFilter { $FailureStage -eq 'Duplicate Versions' -and $ErrorMessage -like '*f087d401*' }
    }

    It 'records a Post-Upload Maintenance failure with the -Repair command when finishing fails' {
        Mock Get-SameAppAllVersions { throw 'Could not retrieve the versions of calibre from Intune: 504' }
        Invoke-YardstickInterruptedUploadHeal -Reason 'r' -ExistingApps $script:bare | Should -BeTrue
        Should -Invoke Add-FailedApplication -Times 1 -Exactly -ParameterFilter {
            $FailureStage -eq 'Post-Upload Maintenance' -and $ErrorMessage -like 'Intune was changed:*-AppId calibre -Repair*'
        }
    }

    It 'reports what is still incomplete' {
        Mock Invoke-YardstickAppMaintenance {
            [pscustomobject]@{ CurrentApp = $script:newest; SupersededCount = 0; PrunedCount = 0; RenamedCount = 1; AnchorStatus = $null; StepErrors = @('supersedence') }
        }
        Invoke-YardstickInterruptedUploadHeal -Reason 'r' -ExistingApps $script:bare | Should -BeTrue
        Should -Invoke Add-SuccessfulApplication -Times 1 -Exactly -ParameterFilter { $AutoUpdateStatus -like '*incomplete: supersedence*' }
    }
}
