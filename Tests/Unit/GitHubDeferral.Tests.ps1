BeforeAll {
    # The deferral functions live in Invoke-YardstickRun.ps1, which cannot be
    # dot-sourced (its top level runs a whole update pass), so extract them via the
    # AST. They read the runner's $GitHubDeferral, $AppId_Processing and
    # $CurrentDisplayName, which each test sets in its own scope.
    $runnerPath = "$PSScriptRoot\..\..\Private\Invoke-YardstickRun.ps1"
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($runnerPath, [ref]$null, [ref]$null)
    foreach ($name in 'Add-YardstickGitHubDeferral', 'Get-YardstickDeferredRecipe') {
        $fn = $ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $fn) { throw "$name not found in $runnerPath" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    # Stand-ins for the module and runner commands they call, for Mock to replace.
    function Write-Log { param([string]$Content, [string]$Level) }
    function Get-YardstickGitHubThrottle { param($ErrorRecord) }
    function Get-YardstickGitHubHold { }
    function Add-FailedApplication { param($ApplicationId, $DisplayName, $Version, $ErrorMessage, $FailureStage) }
    function Set-YardstickLogContext { param($App) }
    function Clear-YardstickLogContext { }
    function Set-YardstickWatchdogStage { param($AppId, $Stage) }
    function Update-YardstickWatchdogProgress { }

    function New-TestDeferral {
        @{
            MaxDeferrals   = 3
            MaxWaitMinutes = 75
            WaitedSeconds  = 0
            Counts         = @{}
            Recipes        = [System.Collections.Generic.List[object]]::new()
        }
    }
}

Describe "Add-YardstickGitHubDeferral" {
    BeforeEach {
        $GitHubDeferral = New-TestDeferral
        $AppId_Processing = 'kindle'
        $CurrentDisplayName = 'Kindle'
        $Script:Version = $null
        Mock Write-Log {}
    }

    It "Leaves a recipe GitHub did not throttle to fail as usual" {
        Mock Get-YardstickGitHubThrottle { $null }
        Add-YardstickGitHubDeferral -Stage 'Pre-Download Script' -ErrorRecord 'boom' | Should -BeFalse
        $GitHubDeferral.Recipes.Count | Should -Be 0
        $GitHubDeferral.Counts.Count | Should -Be 0
    }

    It "Sets a throttled recipe aside until GitHub's wait has passed" {
        $retryAt = (Get-Date).AddMinutes(3)
        $Global:GitHubDeferralTestRetryAt = $retryAt
        Mock Get-YardstickGitHubThrottle {
            [pscustomobject]@{ Host = 'raw.githubusercontent.com'; RetryAt = $Global:GitHubDeferralTestRetryAt }
        }

        Add-YardstickGitHubDeferral -Stage 'Pre-Download Script' -ErrorRecord 'GitHub throttled it' | Should -BeTrue

        $GitHubDeferral.Counts['kindle'] | Should -Be 1
        $GitHubDeferral.Recipes.Count | Should -Be 1
        $recipe = $GitHubDeferral.Recipes[0]
        $recipe.Id | Should -Be 'kindle'
        $recipe.DisplayName | Should -Be 'Kindle'
        $recipe.Stage | Should -Be 'Pre-Download Script'
        $recipe.RetryAt | Should -Be $retryAt
        $recipe.Message | Should -Be 'GitHub throttled it'
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter {
            $Level -eq 'Warning' -and $Content -match "^Deferring kindle until $($retryAt.ToString('HH:mm:ss')):"
        }
        Remove-Variable GitHubDeferralTestRetryAt -Scope Global
    }

    It "Stops setting a recipe aside after MaxDeferrals" {
        Mock Get-YardstickGitHubThrottle { [pscustomobject]@{ Host = 'api.github.com'; RetryAt = (Get-Date).AddMinutes(1) } }
        $GitHubDeferral.Counts['kindle'] = 3
        Add-YardstickGitHubDeferral -Stage 'Download Script' | Should -BeFalse
        $GitHubDeferral.Recipes.Count | Should -Be 0
    }

    It "Does not set a recipe aside for longer than a run waits" {
        Mock Get-YardstickGitHubThrottle { [pscustomobject]@{ Host = 'api.github.com'; RetryAt = (Get-Date).AddMinutes(76) } }
        Add-YardstickGitHubDeferral -Stage 'Download Script' | Should -BeFalse
        $GitHubDeferral.Recipes.Count | Should -Be 0
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Level -eq 'Warning' -and $Content -match 'longer than the 75 minutes' }
    }
}

Describe "Get-YardstickDeferredRecipe" {
    BeforeEach {
        $GitHubDeferral = New-TestDeferral
        Mock Write-Log {}
        Mock Get-YardstickGitHubHold { $null }
        Mock Add-FailedApplication {}
        Mock Set-YardstickLogContext {}
        Mock Clear-YardstickLogContext {}
        Mock Set-YardstickWatchdogStage {}
        Mock Update-YardstickWatchdogProgress {}
        # Real time has to pass for the wait to end; a Start-Sleep in here would
        # call this mock again.
        Mock Start-Sleep { [Threading.Thread]::Sleep(200) }
    }

    It "Returns nothing when no recipe is waiting" {
        @(Get-YardstickDeferredRecipe).Count | Should -Be 0
        Should -Invoke Start-Sleep -Times 0 -Exactly
    }

    It "Returns at once a recipe whose wait has already passed" {
        $GitHubDeferral.Recipes.Add([pscustomobject]@{ Id = 'kindle'; RetryAt = (Get-Date).AddSeconds(-5); Stage = 'Pre-Download Script' })
        Get-YardstickDeferredRecipe | Should -Be 'kindle'
        Should -Invoke Start-Sleep -Times 0 -Exactly
        $GitHubDeferral.Recipes.Count | Should -Be 0
    }

    It "Waits for the earliest recipe, keeping the watchdog alive, and leaves later ones waiting" {
        $GitHubDeferral.Recipes.Add([pscustomobject]@{ Id = 'snagit'; RetryAt = (Get-Date).AddMinutes(10); Stage = 'Pre-Download Script' })
        $GitHubDeferral.Recipes.Add([pscustomobject]@{ Id = 'kindle'; RetryAt = (Get-Date).AddSeconds(1); Stage = 'Pre-Download Script' })

        @(Get-YardstickDeferredRecipe) | Should -Be @('kindle')

        Should -Invoke Set-YardstickWatchdogStage -Times 1 -Exactly -ParameterFilter { $Stage -match '^Waiting for GitHub until ' }
        Should -Invoke Start-Sleep -ParameterFilter { $Seconds -le 30 }
        Should -Invoke Update-YardstickWatchdogProgress
        $GitHubDeferral.WaitedSeconds | Should -BeLessOrEqual 2
        @($GitHubDeferral.Recipes.Id) | Should -Be @('snagit')
        Should -Invoke Add-FailedApplication -Times 0 -Exactly
    }

    It "Fails every waiting recipe when the next wait would run past the budget" {
        $GitHubDeferral.WaitedSeconds = 75 * 60 - 60
        $GitHubDeferral.Recipes.Add([pscustomobject]@{ Id = 'kindle'; DisplayName = 'Kindle'; Version = $null; RetryAt = (Get-Date).AddMinutes(5); Stage = 'Pre-Download Script'; Message = 'Skipped it' })
        $GitHubDeferral.Recipes.Add([pscustomobject]@{ Id = 'snagit'; DisplayName = 'Snagit'; Version = $null; RetryAt = (Get-Date).AddMinutes(6); Stage = 'Version Validation'; Message = '' })

        @(Get-YardstickDeferredRecipe).Count | Should -Be 0

        Should -Invoke Start-Sleep -Times 0 -Exactly
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Content -eq 'Starting update for kindle...' }
        Should -Invoke Add-FailedApplication -Times 1 -Exactly -ParameterFilter {
            $ApplicationId -eq 'kindle' -and $FailureStage -eq 'Pre-Download Script' -and $ErrorMessage -match 'past the 75 minutes a run waits for GitHub\. Skipped it$'
        }
        Should -Invoke Add-FailedApplication -Times 1 -Exactly -ParameterFilter { $ApplicationId -eq 'snagit' -and $FailureStage -eq 'Version Validation' }
        $GitHubDeferral.Recipes.Count | Should -Be 0
    }

    It "Picks up a longer wait GitHub asked for after the recipe was set aside" {
        $hold = (Get-Date).AddMinutes(20)
        $Global:GitHubDeferralTestHold = $hold
        Mock Get-YardstickGitHubHold { $Global:GitHubDeferralTestHold }
        $GitHubDeferral.WaitedSeconds = 75 * 60 - 600
        $GitHubDeferral.Recipes.Add([pscustomobject]@{ Id = 'kindle'; DisplayName = 'Kindle'; Version = $null; RetryAt = (Get-Date).AddSeconds(-1); Stage = 'Pre-Download Script'; Message = '' })

        @(Get-YardstickDeferredRecipe).Count | Should -Be 0
        Should -Invoke Add-FailedApplication -Times 1 -Exactly -ParameterFilter { $ErrorMessage -match "wait until $($hold.ToString('HH:mm'))" }
        Remove-Variable GitHubDeferralTestHold -Scope Global
    }
}
