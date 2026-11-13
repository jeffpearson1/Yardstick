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

    It "also bounds a response body that stops arriving" {
        # -TimeoutSec only covers the headers. The TreeSize host on 2026-11-04 sent
        # headers and half the installer, then went quiet with the socket open.
        Set-YardstickWebRequestTimeout -TimeoutSeconds 120
        $Global:PSDefaultParameterValues['Invoke-WebRequest:OperationTimeoutSeconds'] | Should -Be 120
        $Global:PSDefaultParameterValues['Invoke-RestMethod:OperationTimeoutSeconds'] | Should -Be 120
    }

    It "removes the defaults when given 0" {
        Set-YardstickWebRequestTimeout -TimeoutSeconds 120
        Set-YardstickWebRequestTimeout -TimeoutSeconds 0
        $Global:PSDefaultParameterValues.Keys | Where-Object { $_ -like 'Invoke-*' } | Should -BeNullOrEmpty
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

    It "abandons a stuck stage by interrupting the main thread, and logs it once" {
        # The interrupt lands on this thread - Pester's - exactly as it would land on
        # the runner's, and surfaces as an error the stage's catch can handle.
        Start-YardstickWatchdog -TimeoutMinutes 1 -PollSeconds 1 -NoExit -AbandonGraceMinutes 60
        $state = Get-YardstickWatchdogState
        $state['StageStarted'] = (Get-Date).AddMinutes(-5)

        $caught = $null
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            Start-Sleep -Seconds 20
        } catch [System.Threading.ThreadInterruptedException] {
            $caught = $_
        }
        $timer.Stop()
        Start-Sleep -Seconds 2   # more polls on the same stuck stage

        $caught | Should -Not -BeNullOrEmpty -Because 'the blocked call must let go rather than wait out its 20 seconds'
        $timer.Elapsed.TotalSeconds | Should -BeLessThan 15
        $state['ExpiredFor'] | Should -Be $state['StageStarted']
        $logged = @(Get-Content -LiteralPath (Join-Path $Global:LogLocation $Global:LogFile) |
            Where-Object { $_ -match 'WATCHDOG: no progress' })
        $logged.Count | Should -Be 1 -Because 'one stuck stage is abandoned once, not on every poll'
        $logged[-1] | Should -Match 'Abandoning'
        (Get-Job -Name 'YardstickWatchdog').State | Should -Be 'Running'
    }

    It "with -NoExit, only logs when an abandoned stage does not let go within the grace period" {
        # The stage here catches the interrupt and carries on, as a recipe with its
        # own catch-all might. Without -NoExit this would end the Pester process
        # itself, which is the point: at a prompt the console must survive.
        Start-YardstickWatchdog -TimeoutMinutes 1 -PollSeconds 1 -NoExit -AbandonGraceMinutes 0
        $state = Get-YardstickWatchdogState
        $state['StageStarted'] = (Get-Date).AddMinutes(-5)

        $interrupts = 0
        $deadline = (Get-Date).AddSeconds(20)
        while (-not $state['GaveUpFor'] -and (Get-Date) -lt $deadline) {
            # Whole seconds, so the thread is almost always waiting when the
            # watchdog looks; it does not interrupt a thread that is running.
            try { Start-Sleep -Seconds 1 } catch [System.Threading.ThreadInterruptedException] { $interrupts++ }
        }
        try { Start-Sleep -Seconds 2 } catch [System.Threading.ThreadInterruptedException] { $interrupts++ }

        $state['GaveUpFor'] | Should -Be $state['StageStarted']
        $interrupts | Should -Be 1 -Because 'a stage is interrupted once; after that the only recourse is ending the run'
        $logged = @(Get-Content -LiteralPath (Join-Path $Global:LogLocation $Global:LogFile) |
            Where-Object { $_ -match 'did not let go' })
        $logged.Count | Should -Be 1
        $logged[-1] | Should -Match 'Not ending the run'
        (Get-Job -Name 'YardstickWatchdog').State | Should -Be 'Running'
    }
}

Describe "Watchdog interrupt hand-off" {
    # Thread.Interrupt on a thread that is not waiting stays pending until its next
    # wait, wherever that is. The watchdog only interrupts a waiting thread, and the
    # main-thread side clears an interrupt that slipped through, so one meant for an
    # abandoned stage cannot reach whatever runs after it.
    AfterEach {
        Stop-YardstickWatchdog
        # Never let a failed test leave an interrupt pending for the rest of the suite.
        try { [System.Threading.Thread]::Sleep(0) } catch [System.Threading.ThreadInterruptedException] { $null = $_ }
    }

    It "does not interrupt a thread that is busy rather than waiting" {
        # Start-Process -Wait and native calls read as Running, like this loop. An
        # interrupt sent there would go off later, somewhere unrelated.
        $box = [hashtable]::Synchronized(@{})
        $busy = Start-ThreadJob -ArgumentList $box -ScriptBlock {
            param($box)
            $box.Thread = [System.Threading.Thread]::CurrentThread
            $end = (Get-Date).AddSeconds(5)
            while ((Get-Date) -lt $end) { }
            # Stay Running until the test has stopped the watchdog. Without this
            # hand-off, its one-second poll can coincide with the sleep below and
            # legitimately observe WaitSleepJoin, making the test timing-dependent.
            $box.BusyDone = $true
            while (-not $box.CheckPending) { }
            try { Start-Sleep -Milliseconds 100; 'nothing pending' } catch { 'an interrupt was pending' }
        }
        while (-not $box.Thread) { Start-Sleep -Milliseconds 20 }

        Start-YardstickWatchdog -TimeoutMinutes 1 -PollSeconds 1 -NoExit -AbandonGraceMinutes 60 -MainThread $box.Thread
        $state = Get-YardstickWatchdogState
        $state['StageStarted'] = (Get-Date).AddMinutes(-5)

        $deadline = (Get-Date).AddSeconds(10)
        try {
            while (-not $box.BusyDone -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 20 }
            while (-not $state['ExpiredFor'] -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 20 }
            Stop-YardstickWatchdog
        } finally {
            # Always release the job, even when watchdog teardown itself fails.
            $box.CheckPending = $true
        }
        $result = Receive-Job -Job $busy -Wait -AutoRemoveJob

        $state['ExpiredFor'] | Should -Be $state['StageStarted'] -Because 'the stage was abandoned'
        $state['InterruptedFor'] | Should -BeNullOrEmpty
        $result | Should -Be 'nothing pending'
    }

    It "Set-YardstickWatchdogStage clears an interrupt the abandoned stage left pending" {
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        $state = Get-YardstickWatchdogState
        $state['InterruptedFor'] = $state['StageStarted']

        $escaped = $null
        try {
            [System.Threading.Thread]::CurrentThread.Interrupt()
            Set-YardstickWatchdogStage -AppId 'next' -Stage 'Starting up'
            Start-Sleep -Milliseconds 50
        } catch {
            $escaped = $_
        }
        $escaped | Should -BeNullOrEmpty -Because 'the next stage must not inherit the old stage''s interrupt'
    }

    It "Stop-YardstickWatchdog clears an interrupt the last stage left pending" {
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        $state = Get-YardstickWatchdogState
        $state['InterruptedFor'] = $state['StageStarted']

        $escaped = $null
        try {
            [System.Threading.Thread]::CurrentThread.Interrupt()
            Stop-YardstickWatchdog
            Start-Sleep -Milliseconds 50
        } catch {
            $escaped = $_
        }
        $escaped | Should -BeNullOrEmpty -Because 'the email report runs after this'
    }

    It "Set-YardstickWatchdogStage forgets an abandonment the recipe recovered from" {
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        $state = Get-YardstickWatchdogState
        $state['Abandoned'] = @{ AppId = 'treesize'; Stage = 'Pre-Download Script'; Minutes = 60.2; TimeoutMinutes = 60 }
        Set-YardstickWatchdogStage -Stage 'Download'
        $state['Abandoned'] | Should -BeNullOrEmpty
    }
}

Describe "Add-FailedApplication after the watchdog abandons a recipe" {
    BeforeEach {
        Initialize-ApplicationTracker
        Start-YardstickWatchdog -TimeoutMinutes 60 -PollSeconds 3600
        (Get-YardstickWatchdogState)['Abandoned'] = @{ AppId = 'treesize'; Stage = 'Pre-Download Script'; Minutes = 60.2; TimeoutMinutes = 60 }
    }

    AfterEach {
        Stop-YardstickWatchdog
    }

    It "replaces the interrupt's message with what happened" {
        Add-FailedApplication -ApplicationId 'treesize' -ErrorMessage 'Thread was interrupted from a waiting state.' -FailureStage 'Pre-Download Script'
        $failure = InModuleScope YardstickSupport { $Script:FailedApplications[0] }
        $failure.ErrorMessage | Should -Match "Abandoned after 60.2 minutes without progress in stage 'Pre-Download Script'"
        $failure.ErrorMessage | Should -Match 'moved on to the next recipe'
    }

    It "does it once, so a later failure of the same app keeps its own message" {
        Add-FailedApplication -ApplicationId 'treesize' -ErrorMessage 'Thread was interrupted from a waiting state.' -FailureStage 'Pre-Download Script'
        (Get-YardstickWatchdogState)['Abandoned'] | Should -BeNullOrEmpty
        Add-FailedApplication -ApplicationId 'treesize' -ErrorMessage 'token revoke failed' -FailureStage 'Post-Run Script'
        Get-Content -LiteralPath (Join-Path $Global:LogLocation $Global:LogFile) -Tail 1 | Should -Match 'Post-Run Script failed: token revoke failed'
    }

    It "leaves other apps' failures alone" {
        Add-FailedApplication -ApplicationId 'firefox' -ErrorMessage 'HTTP 404' -FailureStage 'Download'
        $failure = InModuleScope YardstickSupport { $Script:FailedApplications[0] }
        $failure.ErrorMessage | Should -Be 'HTTP 404'
        (Get-YardstickWatchdogState)['Abandoned'] | Should -Not -BeNullOrEmpty
    }
}

Describe "Test-YardstickDedicatedProcess" {
    # Decides whether the watchdog may end the process: only when pwsh was started
    # to run a script or command and will exit afterwards anyway.
    It "is true for <Case>" -ForEach @(
        @{ Case = 'the scheduled task';       Arguments = @('pwsh.exe', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'C:\Yardstick\Yardstick.ps1', '-All', '-NoInteractive') }
        @{ Case = 'pwsh -Command';            Arguments = @('pwsh.exe', '-Command', 'Invoke-Yardstick -All') }
        @{ Case = 'abbreviated -f';           Arguments = @('pwsh.exe', '-f', 'Yardstick.ps1') }
        @{ Case = '-EncodedCommand as -ec';   Arguments = @('pwsh.exe', '-ec', 'AAAA') }
        @{ Case = 'a Windows-style /File';    Arguments = @('pwsh.exe', '/File', 'Yardstick.ps1') }
        @{ Case = 'a script parameter named -NoExit'; Arguments = @('pwsh.exe', '-File', 'Yardstick.ps1', '-NoExit') }
    ) {
        Test-YardstickDedicatedProcess -CommandLineArgs $Arguments | Should -BeTrue
    }

    It "is false for <Case>" -ForEach @(
        @{ Case = 'an interactive console';   Arguments = @('pwsh.exe') }
        @{ Case = 'a console with -NoLogo';   Arguments = @('pwsh.exe', '-NoLogo') }
        @{ Case = '-NoExit before -File';     Arguments = @('pwsh.exe', '-NoExit', '-File', 'Yardstick.ps1') }
        @{ Case = 'an editor host (-NoExit -Command)'; Arguments = @('pwsh.exe', '-NoProfile', '-NoExit', '-Command', 'Import-Module x') }
        @{ Case = '-ExecutionPolicy alone';   Arguments = @('pwsh.exe', '-ExecutionPolicy', 'Bypass') }
    ) {
        Test-YardstickDedicatedProcess -CommandLineArgs $Arguments | Should -BeFalse
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

    It "passes -Dynamic through to BITS" {
        # ltspice's server needs it.
        Mock -ModuleName YardstickSupport Start-BitsTransfer {
            New-FakeBitsJob -StateSequence { 'Transferred' } -ByteSequence { 1000 }
        }
        Mock -ModuleName YardstickSupport Complete-BitsTransfer {} -RemoveParameterType 'BitsJob'

        Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' -PollSeconds 1 -Dynamic
        Should -Invoke Start-BitsTransfer -ModuleName YardstickSupport -Times 1 -ParameterFilter { $Dynamic -and $Asynchronous }
    }

    It "uses the run's stall bound when the caller gives none" {
        # Recipes call it without -StallTimeoutMinutes; bitsStallTimeoutMinutes
        # has to reach them through Set-YardstickBitsStallTimeout.
        Mock -ModuleName YardstickSupport Start-BitsTransfer {
            New-FakeBitsJob -StateSequence { 'Transferring' } -ByteSequence { 0 }
        }
        Mock -ModuleName YardstickSupport Remove-BitsTransfer {} -RemoveParameterType 'BitsJob'
        Mock -ModuleName YardstickSupport Start-Sleep {}
        $Global:FakeClock = [datetime]'2026-11-04T09:00:00'
        Mock -ModuleName YardstickSupport Get-Date {
            $Global:FakeClock = $Global:FakeClock.AddMinutes(2)
            $Global:FakeClock
        }

        try {
            Set-YardstickBitsStallTimeout -Minutes 3
            { Invoke-YardstickBitsDownload -Source 'https://x/y' -Destination 'C:\y' -PollSeconds 1 } |
                Should -Throw -ExpectedMessage '*moved no data for 3 minute(s)*'
        } finally {
            Set-YardstickBitsStallTimeout -Minutes 15
        }
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
    # The runner sets $ErrorActionPreference = 'Stop' at the top, which makes
    # Write-Error a TERMINATING error. Every per-stage failure handler in the main
    # loop was written as `Add-FailedApplication; Write-Error; continue`, so the
    # Write-Error threw, the continue never ran, and control landed in the outer
    # catch - which recorded the same failure a second time under "General
    # Processing" and logged "Unexpected error processing <app>" instead of the
    # real cause. The loop is top-level script code and cannot be executed here,
    # so assert its shape.
    BeforeAll {
        $script:YardstickPath = "$PSScriptRoot\..\..\Private\Invoke-YardstickRun.ps1"
        $script:YardstickAst = [System.Management.Automation.Language.Parser]::ParseFile(
            $script:YardstickPath, [ref]$null, [ref]$null)

        # The main loop is the foreach over each pass of recipes: every recipe
        # first, then the ones GitHub asked to wait.
        $script:MainLoop = $script:YardstickAst.Find({
            param($node)
            $node -is [System.Management.Automation.Language.ForEachStatementAst] -and
            $node.Condition.Extent.Text -match '\$RecipePass\b'
        }, $true)
    }

    It "finds the main processing loop" {
        $script:MainLoop | Should -Not -BeNullOrEmpty
    }

    It "runs the recipes GitHub asked to wait in a later pass" {
        $passLoop = $script:MainLoop.Parent
        while ($passLoop -and $passLoop -isnot [System.Management.Automation.Language.WhileStatementAst]) {
            $passLoop = $passLoop.Parent
        }
        $passLoop | Should -Not -BeNullOrEmpty
        $passLoop.Condition.Extent.Text | Should -Match '\$RecipePass\.Count'

        # The pass loop's last statement fetches the next pass.
        $last = $passLoop.Body.Statements[-1]
        $last.Extent.Text | Should -Match '^\$RecipePass\s*=\s*@\(Get-YardstickDeferredRecipe\)$'
    }

    It "offers the recipe to Add-YardstickGitHubDeferral before failing <Stage>" -ForEach @(
        @{ Stage = 'Pre-Download Script' }
        @{ Stage = 'Version Validation' }
        @{ Stage = 'Download Script' }
        @{ Stage = 'Post-Download Script' }
    ) {
        $deferral = $script:MainLoop.Find({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -eq 'Add-YardstickGitHubDeferral' -and
            $node.Extent.Text -match [regex]::Escape("`"$Stage`"")
        }, $true)
        $deferral | Should -Not -BeNullOrEmpty

        # if (Add-YardstickGitHubDeferral ...) { continue }, ahead of the failure.
        $if = $deferral.Parent
        while ($if -and $if -isnot [System.Management.Automation.Language.IfStatementAst]) { $if = $if.Parent }
        $if.Clauses[0].Item2.Extent.Text | Should -Match '^\{\s*continue\s*\}$'
        $failure = $if.Parent.Statements | Where-Object {
            $_.Extent.StartOffset -gt $if.Extent.EndOffset -and $_.Extent.Text -match '^Add-FailedApplication\b'
        } | Select-Object -First 1
        $failure.Extent.Text | Should -Match ([regex]::Escape("-FailureStage `"$Stage`""))
    }

    It "does not let minRunInterval skip a recipe that was set aside" {
        # The first attempt recorded the run, so the replay would otherwise be
        # skipped as run too recently.
        $guard = $script:MainLoop.Find({
            param($node)
            $node -is [System.Management.Automation.Language.IfStatementAst] -and
            $node.Clauses[0].Item1.Extent.Text -match 'MinRunInterval'
        }, $true)
        $guard.Clauses[0].Item1.Extent.Text | Should -Match '-not \$IsDeferredReplay'
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

    It "gives the backup drain a watchdog stage of its own" {
        # The drain is outside every catch. Changing stage first is what clears an
        # interrupt the watchdog sent to an abandoned stage before it, which would
        # otherwise throw out of the finally and end the run.
        $tryStatement = $script:MainLoop.Body.Find({
            param($node)
            $node -is [System.Management.Automation.Language.TryStatementAst]
        }, $false)
        $tryStatement.Finally.Extent.Text | Should -Match "Set-YardstickWatchdogStage -Stage 'Waiting for backup'\s+Wait-YardstickBackup"
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

    It "pops every location it pushes in a finally, so a failed recipe script cannot leave the stack behind" {
        $pushes = @($script:MainLoop.FindAll({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and
            $node.GetCommandName() -eq 'Push-Location'
        }, $true))
        $pushes.Count | Should -BeGreaterThan 0

        foreach ($push in $pushes) {
            # The push is followed, in the same block, by a try whose finally pops.
            $block = $push.Parent
            while ($block -and $block -isnot [System.Management.Automation.Language.StatementBlockAst] -and
                $block -isnot [System.Management.Automation.Language.NamedBlockAst]) {
                $block = $block.Parent
            }
            $try = $block.Statements | Where-Object {
                $_ -is [System.Management.Automation.Language.TryStatementAst] -and
                $_.Extent.StartOffset -gt $push.Extent.EndOffset
            } | Select-Object -First 1
            $try | Should -Not -BeNullOrEmpty -Because "line $($push.Extent.StartLineNumber) pushes a location"
            $try.Finally.Extent.Text | Should -Match 'Pop-Location' -Because "line $($push.Extent.StartLineNumber) pushes a location"
        }
    }
}
