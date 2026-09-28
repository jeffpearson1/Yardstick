BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force
}

Describe "Get-YardstickTimeoutPreference" {
    It "returns the default when the key is absent" {
        Get-YardstickTimeoutPreference -Value $null -Default 120 | Should -Be 120
    }

    It "returns the default when the key is present but blank" {
        # ConvertFrom-Yaml gives "" for `webRequestTimeoutSeconds:` with no value.
        Get-YardstickTimeoutPreference -Value "" -Default 120 | Should -Be 120
        Get-YardstickTimeoutPreference -Value "   " -Default 120 | Should -Be 120
    }

    It "returns a configured value" {
        Get-YardstickTimeoutPreference -Value 45 -Default 120 | Should -Be 45
    }

    It "accepts a value YAML parsed as a string" {
        Get-YardstickTimeoutPreference -Value "45" -Default 120 | Should -Be 45
    }

    It "falls back to the default on an explicit 0 without -AllowZero" {
        Get-YardstickTimeoutPreference -Value 0 -Default 120 | Should -Be 120
    }

    It "honours an explicit 0 with -AllowZero" {
        Get-YardstickTimeoutPreference -Value 0 -Default 120 -AllowZero | Should -Be 0
    }

    It "falls back to the default on junk rather than throwing" {
        Get-YardstickTimeoutPreference -Value "soon" -Default 120 | Should -Be 120
    }

    It "falls back to the default on a negative value" {
        Get-YardstickTimeoutPreference -Value -5 -Default 120 | Should -Be 120
    }
}

Describe "Set-YardstickWebRequestTimeout" {
    BeforeEach {
        $Global:PSDefaultParameterValues = @{}
    }

    AfterAll {
        $Global:PSDefaultParameterValues = @{}
    }

    It "sets the default for both web cmdlets" {
        Set-YardstickWebRequestTimeout -TimeoutSeconds 120
        $Global:PSDefaultParameterValues['Invoke-WebRequest:TimeoutSec'] | Should -Be 120
        $Global:PSDefaultParameterValues['Invoke-RestMethod:TimeoutSec'] | Should -Be 120
    }

    It "removes the defaults when given 0" {
        Set-YardstickWebRequestTimeout -TimeoutSeconds 120
        Set-YardstickWebRequestTimeout -TimeoutSeconds 0
        $Global:PSDefaultParameterValues.ContainsKey('Invoke-WebRequest:TimeoutSec') | Should -BeFalse
        $Global:PSDefaultParameterValues.ContainsKey('Invoke-RestMethod:TimeoutSec') | Should -BeFalse
    }

    It "re-asserts after a recipe wipes the table" {
        # This is the whole reason the runner calls it once per iteration: a recipe
        # that does `$PSDefaultParameterValues = @{}` would otherwise leave every
        # later recipe able to hang indefinitely.
        Set-YardstickWebRequestTimeout -TimeoutSeconds 120
        $Global:PSDefaultParameterValues = @{}
        Set-YardstickWebRequestTimeout -TimeoutSeconds 120
        $Global:PSDefaultParameterValues['Invoke-WebRequest:TimeoutSec'] | Should -Be 120
    }

    It "leaves unrelated defaults alone" {
        $Global:PSDefaultParameterValues['Get-ChildItem:Force'] = $true
        Set-YardstickWebRequestTimeout -TimeoutSeconds 120
        $Global:PSDefaultParameterValues['Get-ChildItem:Force'] | Should -BeTrue
    }

    It "builds the table when nothing has created it yet" {
        $Global:PSDefaultParameterValues = $null
        Set-YardstickWebRequestTimeout -TimeoutSeconds 90
        $Global:PSDefaultParameterValues['Invoke-RestMethod:TimeoutSec'] | Should -Be 90
    }
}

Describe "Test-YardstickWatchdogExpired" {
    BeforeAll {
        $Global:Now = [datetime]'2026-09-25T09:00:00'
    }

    It "is not expired inside the budget" {
        $state = @{ StageStarted = $Global:Now.AddMinutes(-59) }
        Test-YardstickWatchdogExpired -State $state -Now $Global:Now -TimeoutMinutes 60 | Should -BeFalse
    }

    It "is expired exactly at the budget" {
        $state = @{ StageStarted = $Global:Now.AddMinutes(-60) }
        Test-YardstickWatchdogExpired -State $state -Now $Global:Now -TimeoutMinutes 60 | Should -BeTrue
    }

    It "is expired past the budget" {
        $state = @{ StageStarted = $Global:Now.AddMinutes(-125) }
        Test-YardstickWatchdogExpired -State $state -Now $Global:Now -TimeoutMinutes 60 | Should -BeTrue
    }

    It "never fires when the budget is 0" {
        $state = @{ StageStarted = $Global:Now.AddDays(-1) }
        Test-YardstickWatchdogExpired -State $state -Now $Global:Now -TimeoutMinutes 0 | Should -BeFalse
    }

    It "never fires on a negative budget" {
        $state = @{ StageStarted = $Global:Now.AddDays(-1) }
        Test-YardstickWatchdogExpired -State $state -Now $Global:Now -TimeoutMinutes -1 | Should -BeFalse
    }

    It "does not fire when the state has no stage start" {
        Test-YardstickWatchdogExpired -State @{} -Now $Global:Now -TimeoutMinutes 60 | Should -BeFalse
    }

    It "does not fire on an unusable stage start" {
        $state = @{ StageStarted = 'whenever' }
        Test-YardstickWatchdogExpired -State $state -Now $Global:Now -TimeoutMinutes 60 | Should -BeFalse
    }
}

Describe "Set-YardstickWatchdogStage" {
    AfterEach {
        Stop-YardstickWatchdog
    }

    It "does nothing when no watchdog is running" {
        Stop-YardstickWatchdog
        { Set-YardstickWatchdogStage -AppId 'firefox' -Stage 'Download' } | Should -Not -Throw
    }

    It "resets the clock, which is what makes the budget per-stage rather than per-recipe" {
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        $state = Get-YardstickWatchdogState

        # Backdate past the budget, then report progress. A per-recipe budget would
        # still be expired here; a per-stage one is not.
        $state['StageStarted'] = (Get-Date).AddMinutes(-90)
        Test-YardstickWatchdogExpired -State $state -Now (Get-Date) -TimeoutMinutes 60 | Should -BeTrue

        Set-YardstickWatchdogStage -AppId 'firefox' -Stage 'Download'
        Test-YardstickWatchdogExpired -State $state -Now (Get-Date) -TimeoutMinutes 60 | Should -BeFalse
    }

    It "records the app and stage for the diagnosis" {
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        Set-YardstickWatchdogStage -AppId 'firefox' -Stage 'Pre-Download Script'
        $state = Get-YardstickWatchdogState
        $state['AppId'] | Should -Be 'firefox'
        $state['Stage'] | Should -Be 'Pre-Download Script'
    }

    It "keeps the current app when only the stage is given" {
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        Set-YardstickWatchdogStage -AppId 'firefox' -Stage 'Download'
        Set-YardstickWatchdogStage -Stage 'Packaging'
        $state = Get-YardstickWatchdogState
        $state['AppId'] | Should -Be 'firefox'
        $state['Stage'] | Should -Be 'Packaging'
    }
}

Describe "Start-YardstickWatchdog" {
    AfterEach {
        Stop-YardstickWatchdog
    }

    It "starts no thread when the budget is 0" {
        Start-YardstickWatchdog -TimeoutMinutes 0
        Get-YardstickWatchdogState | Should -BeNullOrEmpty
        Get-Job -Name 'YardstickWatchdog' -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
    }

    It "runs on a real thread sharing this process, with state both threads can see" {
        # The whole design rests on this. A Register-ObjectEvent timer would never
        # observe a change while the main thread is blocked, because its handler
        # runs on the main runspace's event queue. A thread job runs on a separate
        # thread in the same AppDomain, so the hashtable is genuinely shared rather
        # than serialised across a process boundary.
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        $job = Get-Job -Name 'YardstickWatchdog'
        $job | Should -Not -BeNullOrEmpty
        $job.PSJobTypeName | Should -Be 'ThreadJob'

        $state = Get-YardstickWatchdogState
        $state.IsSynchronized | Should -BeTrue -Because 'two threads write to it'
    }

    It "replaces a watchdog that is already running rather than stacking a second one" {
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        @(Get-Job -Name 'YardstickWatchdog').Count | Should -Be 1
    }
}

Describe "Stop-YardstickWatchdog" {
    It "is safe to call when nothing is running" {
        { Stop-YardstickWatchdog } | Should -Not -Throw
        { Stop-YardstickWatchdog } | Should -Not -Throw
    }

    It "clears the state and removes the job" {
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        Stop-YardstickWatchdog
        Get-YardstickWatchdogState | Should -BeNullOrEmpty
        Get-Job -Name 'YardstickWatchdog' -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
    }
}

Describe "Update-YardstickWatchdogProgress" {
    AfterEach {
        Stop-YardstickWatchdog
    }

    It "does nothing when no watchdog is running" {
        Stop-YardstickWatchdog
        { Update-YardstickWatchdogProgress } | Should -Not -Throw
    }

    It "keeps a slow-but-moving stage alive past its budget" {
        # The false positive worth caring about: a 1.6 GB download on a slow link
        # is not a hang, and killing it would break runs that work today.
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        Set-YardstickWatchdogStage -AppId 'clion' -Stage 'Download'
        $state = Get-YardstickWatchdogState

        $state['StageStarted'] = (Get-Date).AddMinutes(-90)
        Test-YardstickWatchdogExpired -State $state -Now (Get-Date) -TimeoutMinutes 60 | Should -BeTrue

        Update-YardstickWatchdogProgress
        Test-YardstickWatchdogExpired -State $state -Now (Get-Date) -TimeoutMinutes 60 | Should -BeFalse
    }

    It "does not change which stage is being reported" {
        # It must not relabel the stage, or the diagnosis would name the wrong one.
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        Set-YardstickWatchdogStage -AppId 'clion' -Stage 'Upload'
        Update-YardstickWatchdogProgress
        $state = Get-YardstickWatchdogState
        $state['AppId'] | Should -Be 'clion'
        $state['Stage'] | Should -Be 'Upload'
    }

    It "re-arms the half-budget warning" {
        # WarnedFor is keyed on the stage start, so a reset must clear the way for
        # a later stall in the same stage to warn again.
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        Set-YardstickWatchdogStage -AppId 'clion' -Stage 'Upload'
        $state = Get-YardstickWatchdogState
        $state['WarnedFor'] = $state['StageStarted']

        Update-YardstickWatchdogProgress
        $state['WarnedFor'] | Should -Not -Be $state['StageStarted']
    }
}

Describe "Invoke-YardstickBitsDownload" {
    BeforeAll {
        # A stand-in for a BITS job: JobState and BytesTransferred are read on every
        # poll, so a script property lets each test drive the transfer's behaviour.
        function New-FakeBitsJob {
            param([scriptblock]$StateSequence, [scriptblock]$ByteSequence)
            $job = [pscustomobject]@{ BytesTotal = 1000; ErrorDescription = 'the far end gave up' }
            $job | Add-Member -MemberType ScriptProperty -Name JobState -Value $StateSequence
            $job | Add-Member -MemberType ScriptProperty -Name BytesTransferred -Value $ByteSequence
            return $job
        }
    }

    It "completes a transfer that finishes" {
        Mock -ModuleName YardstickSupport Start-BitsTransfer {
            New-FakeBitsJob -StateSequence { 'Transferred' } -ByteSequence { 1000 }
        }
        Mock -ModuleName YardstickSupport Complete-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Remove-BitsTransfer {} -RemoveParameterType 'BitsJob'

        { Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' -PollSeconds 1 } |
            Should -Not -Throw
        Should -Invoke Complete-BitsTransfer -ModuleName YardstickSupport -Times 1
        Should -Invoke Remove-BitsTransfer -ModuleName YardstickSupport -Times 0
    }

    It "does not abandon a slow transfer that is still making progress" {
        # The bound is a stall timeout, not a total one, so a big file on a slow
        # link must survive even a 0-minute window as long as bytes keep moving.
        $Global:FakeBytes = 0
        Mock -ModuleName YardstickSupport Start-BitsTransfer {
            New-FakeBitsJob `
                -StateSequence { if ($Global:FakeBytes -ge 1000) { 'Transferred' } else { 'Transferring' } } `
                -ByteSequence { $Global:FakeBytes += 100; $Global:FakeBytes }
        }
        Mock -ModuleName YardstickSupport Complete-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Remove-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Start-Sleep {}

        { Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' -StallTimeoutMinutes 1 -PollSeconds 1 } |
            Should -Not -Throw
        Should -Invoke Complete-BitsTransfer -ModuleName YardstickSupport -Times 1
    }

    It "abandons and cleans up a transfer that moves no data" {
        Mock -ModuleName YardstickSupport Start-BitsTransfer {
            New-FakeBitsJob -StateSequence { 'Transferring' } -ByteSequence { 0 }
        }
        Mock -ModuleName YardstickSupport Complete-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Remove-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Start-Sleep {}
        # Advance the clock two minutes per poll so a one-minute stall window is
        # crossed without the test actually waiting.
        $Global:FakeClock = [datetime]'2026-09-25T09:00:00'
        Mock -ModuleName YardstickSupport Get-Date {
            $Global:FakeClock = $Global:FakeClock.AddMinutes(2)
            $Global:FakeClock
        }

        { Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' -StallTimeoutMinutes 1 -PollSeconds 1 } |
            Should -Throw -ExpectedMessage '*moved no data*'
        # A BITS job left behind keeps retrying under its own service for days.
        Should -Invoke Remove-BitsTransfer -ModuleName YardstickSupport -Times 1
    }

    It "never abandons a transfer when the stall bound is disabled" {
        $Global:FakeBytes2 = 0
        Mock -ModuleName YardstickSupport Start-BitsTransfer {
            New-FakeBitsJob `
                -StateSequence { if ($Global:FakeBytes2 -ge 3) { 'Transferred' } else { 'Transferring' } } `
                -ByteSequence { $Global:FakeBytes2++; 0 }
        }
        Mock -ModuleName YardstickSupport Complete-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Remove-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Start-Sleep {}

        # Zero bytes throughout, but StallTimeoutMinutes 0 means "no bound".
        { Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' -StallTimeoutMinutes 0 -PollSeconds 1 } |
            Should -Not -Throw
        Should -Invoke Complete-BitsTransfer -ModuleName YardstickSupport -Times 1
    }

    It "reports progress to the watchdog so a slow download is not killed" {
        # A 2 GB download can outlast the watchdog's stage budget on a slow link.
        # The poll loop is the only place that knows bytes are still moving.
        $Global:FakeBytes3 = 0
        Mock -ModuleName YardstickSupport Start-BitsTransfer {
            New-FakeBitsJob `
                -StateSequence { if ($Global:FakeBytes3 -ge 300) { 'Transferred' } else { 'Transferring' } } `
                -ByteSequence { $Global:FakeBytes3 += 100; $Global:FakeBytes3 }
        }
        Mock -ModuleName YardstickSupport Complete-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Remove-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Start-Sleep {}
        Mock -ModuleName YardstickSupport Update-YardstickWatchdogProgress {}

        Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' -PollSeconds 1
        Should -Invoke Update-YardstickWatchdogProgress -ModuleName YardstickSupport -Times 1 -Because 'each advance in bytes is progress'
    }

    It "surfaces the BITS error description when the job fails" {
        Mock -ModuleName YardstickSupport Start-BitsTransfer {
            New-FakeBitsJob -StateSequence { 'Error' } -ByteSequence { 0 }
        }
        Mock -ModuleName YardstickSupport Complete-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Remove-BitsTransfer {} -RemoveParameterType 'BitsJob'

        { Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' -PollSeconds 1 } |
            Should -Throw -ExpectedMessage '*the far end gave up*'
        Should -Invoke Remove-BitsTransfer -ModuleName YardstickSupport -Times 1
    }

    It "does not swallow a Start-BitsTransfer failure" {
        Mock -ModuleName YardstickSupport Start-BitsTransfer { throw 'HTTP 404' }
        { Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' } |
            Should -Throw -ExpectedMessage '*404*'
    }
}

Describe "Module web calls are bounded" {
    # Each module has its own session state, so the $Global:PSDefaultParameterValues
    # the runner sets for recipes does NOT reach module code. That is easy to forget
    # when adding a call, so assert it statically rather than trusting review.
    It "every Invoke-WebRequest/Invoke-RestMethod in Modules\*.psm1 passes -TimeoutSec" {
        $offenders = [System.Collections.Generic.List[string]]::new()

        foreach ($file in Get-ChildItem "$PSScriptRoot\..\..\Modules" -Filter '*.psm1' -Recurse) {
            $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            $calls = $ast.FindAll({
                param($node)
                $node -is [System.Management.Automation.Language.CommandAst] -and
                $node.GetCommandName() -in 'Invoke-WebRequest', 'Invoke-RestMethod'
            }, $true)

            foreach ($call in $calls) {
                $bounded = $false
                foreach ($element in $call.CommandElements) {
                    # -TimeoutSec directly, or a splat whose hashtable literal
                    # carries a TimeoutSec key.
                    if ($element -is [System.Management.Automation.Language.CommandParameterAst] -and
                        $element.ParameterName -like 'TimeoutSec*') { $bounded = $true; break }
                    if ($element -is [System.Management.Automation.Language.VariableExpressionAst] -and $element.Splatted) {
                        $name = $element.VariablePath.UserPath
                        $assigned = $ast.FindAll({
                            param($node)
                            $node -is [System.Management.Automation.Language.HashtableAst] -and
                            $node.KeyValuePairs.Item1.Extent.Text -contains 'TimeoutSec'
                        }, $true)
                        if ($assigned -and $name) { $bounded = $true; break }
                    }
                }
                if (-not $bounded) {
                    $offenders.Add("$($file.Name):$($call.Extent.StartLineNumber)")
                }
            }
        }

        $offenders -join ', ' | Should -BeNullOrEmpty -Because 'module web calls cannot inherit the runner global default and would hang the run'
    }
}


Describe "Main processing loop control flow" {
    # Yardstick.ps1 sets $ErrorActionPreference = 'Stop' at the top, which makes
    # Write-Error a TERMINATING error. Every per-stage failure handler in the main
    # loop was written as `Add-FailedApplication; Write-Error; continue`, so the
    # Write-Error threw, the continue never ran, and control landed in the outer
    # catch - which recorded the same failure a second time under "General
    # Processing" and logged "Unexpected error processing <app>" instead of the
    # real cause. The loop is top-level script code and cannot be executed here,
    # so assert its shape.
    BeforeAll {
        $script:YardstickPath = "$PSScriptRoot\..\..\Yardstick.ps1"
        $script:YardstickAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $script:YardstickPath, [ref]$null, [ref]$null)

        # The main loop is the foreach over $Applications.
        $script:MainLoop = $script:YardstickAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.ForEachStatementAst] -and
            $node.Condition.Extent.Text -match '\$Applications\b'
        }, $true)
    }

    It "finds the main processing loop" {
        $script:MainLoop | Should -Not -BeNullOrEmpty
    }

    It "sets ErrorActionPreference to Stop, which is what makes Write-Error terminating" {
        # If this ever stops being true the rest of this Describe is moot rather
        # than wrong, but the next reader deserves to know which assumption moved.
        $assignment = $script:YardstickAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq '$ErrorActionPreference'
        }, $true)

        $assignment | Should -Not -BeNullOrEmpty
        $assignment.Right.Extent.Text | Should -Match "'Stop'|`"Stop`""
    }

    It "uses no Write-Error inside the main loop" {
        $offenders = $script:MainLoop.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -eq 'Write-Error'
        }, $true)

        ($offenders | ForEach-Object { "line $($_.Extent.StartLineNumber)" }) -join ', ' |
            Should -BeNullOrEmpty -Because 'Write-Error throws under ErrorActionPreference Stop, making the `continue` after it unreachable'
    }

    It "runs the per-recipe cleanup in a finally, so a `continue` cannot skip it" {
        # `continue` from inside a try still runs its finally. Without one, every
        # handler that skips to the next recipe would also skip the post-run
        # script - which is a recipe's cleanup hook, and revokes an OAuth token
        # for at least one recipe - and the backup drain.
        $tryStatement = $script:MainLoop.Body.Find({
            param($node)
            $node -is [System.Management.Automation.Language.TryStatementAst]
        }, $false)

        $tryStatement | Should -Not -BeNullOrEmpty
        $tryStatement.Finally | Should -Not -BeNullOrEmpty -Because 'the outer try needs a finally for the cleanup to survive a continue'

        $finallyText = $tryStatement.Finally.Extent.Text
        $finallyText | Should -Match 'PostRunScript'
        $finallyText | Should -Match 'Wait-YardstickBackup'
    }

    It "clears PostRunScript each iteration so an early skip cannot run the previous recipe's hook" {
        # The configuration and recipe-validation handlers skip out before
        # Set-ScriptVariables runs, and the finally fires on those paths too.
        $reset = $script:MainLoop.Body.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left.Extent.Text -eq '$Script:PostRunScript' -and
            $node.Right.Extent.Text -eq '$null'
        }, $true)

        $reset | Should -Not -BeNullOrEmpty
    }
}
