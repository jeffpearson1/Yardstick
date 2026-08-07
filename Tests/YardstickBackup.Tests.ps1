BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force
}

Describe "Get-YardstickBackupFileName" {
    It "builds AppId_Version_yyyyMMdd-HHmmss.intunewin" {
        Get-YardstickBackupFileName -AppId 'firefox' -Version '120.0.1' -Timestamp ([datetime]'2026-08-06T14:03:09') |
            Should -Be 'firefox_120.0.1_20260806-140309.intunewin'
    }

    It "zero-pads single digit date and time components" {
        Get-YardstickBackupFileName -AppId 'a' -Version '1' -Timestamp ([datetime]'2026-01-02T03:04:05') |
            Should -Be 'a_1_20260102-030405.intunewin'
    }

    It "removes every character the filesystem rejects" {
        $bad = -join [System.IO.Path]::GetInvalidFileNameChars().Where({ $_ -ne [char]0 })
        $result = Get-YardstickBackupFileName -AppId "app$bad" -Version '1.0' -Timestamp (Get-Date)
        foreach ($char in [System.IO.Path]::GetInvalidFileNameChars()) {
            $result.IndexOf($char) | Should -Be -1 -Because "'$char' is not legal in a file name"
        }
    }

    It "keeps exactly three underscore-separated fields when the inputs contain underscores" {
        $result = Get-YardstickBackupFileName -AppId 'my_app' -Version '1_0' -Timestamp ([datetime]'2026-08-06T14:03:09')
        $result | Should -Be 'my-app_1-0_20260806-140309.intunewin'
        ($result -split '_').Count | Should -Be 3
    }

    It "collapses consecutive replacements rather than doubling the separator" {
        Get-YardstickBackupFileName -AppId 'a//b' -Version '1.0' -Timestamp ([datetime]'2026-08-06T14:03:09') |
            Should -Be 'a-b_1.0_20260806-140309.intunewin'
    }

    It "throws on a blank AppId" {
        { Get-YardstickBackupFileName -AppId '  ' -Version '1.0' -Timestamp (Get-Date) } | Should -Throw
    }

    It "throws on a blank Version" {
        { Get-YardstickBackupFileName -AppId 'firefox' -Version '' -Timestamp (Get-Date) } | Should -Throw
    }

    It "is deterministic for identical inputs" {
        $stamp = [datetime]'2026-08-06T14:03:09'
        $a = Get-YardstickBackupFileName -AppId 'firefox' -Version '1.0' -Timestamp $stamp
        $b = Get-YardstickBackupFileName -AppId 'firefox' -Version '1.0' -Timestamp $stamp
        $a | Should -Be $b
    }
}

Describe "Get-YardstickBackupTimestamp" {
    It "round-trips a name produced by Get-YardstickBackupFileName" {
        $stamp = [datetime]'2026-08-06T14:03:09'
        $name = Get-YardstickBackupFileName -AppId 'firefox' -Version '120.0.1' -Timestamp $stamp
        Get-YardstickBackupTimestamp -FileName $name | Should -Be $stamp
    }

    It "parses a full path, not just a leaf name" {
        Get-YardstickBackupTimestamp -FileName 'C:\backups\firefox\firefox_1.0_20260806-140309.intunewin' |
            Should -Be ([datetime]'2026-08-06T14:03:09')
    }

    It "parses using the invariant culture regardless of the current culture" {
        $original = [System.Threading.Thread]::CurrentThread.CurrentCulture
        try {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = [cultureinfo]::GetCultureInfo('de-DE')
            Get-YardstickBackupTimestamp -FileName 'firefox_1.0_20260806-140309.intunewin' |
                Should -Be ([datetime]'2026-08-06T14:03:09')
        } finally {
            [System.Threading.Thread]::CurrentThread.CurrentCulture = $original
        }
    }

    It "returns null without throwing for <Case>" -ForEach @(
        @{ Case = 'a name with no timestamp'; Name = 'legacy.intunewin' }
        @{ Case = 'an unparseable timestamp'; Name = 'firefox_1.0_notadate.intunewin' }
        @{ Case = 'an impossible month';      Name = 'firefox_1.0_20261301-140309.intunewin' }
        @{ Case = 'an empty string';          Name = '' }
        @{ Case = 'the wrong extension';      Name = 'firefox_1.0_20260806-140309.zip' }
    ) {
        $script:result = 'sentinel'
        { $script:result = Get-YardstickBackupTimestamp -FileName $Name } | Should -Not -Throw
        $script:result | Should -BeNullOrEmpty
    }
}

Describe "Invoke-YardstickBackupCopy" {
    BeforeEach {
        $script:root = Join-Path $TestDrive "backups"
        $script:src = Join-Path $TestDrive "package.intunewin"
        Set-Content -LiteralPath $script:src -Value ("x" * 4096) -NoNewline
        $script:appDir = Join-Path $script:root "firefox"
    }

    AfterEach {
        Remove-Item -LiteralPath $script:root -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:src -Force -ErrorAction SilentlyContinue
    }

    It "copies the package into a subfolder named for the recipe id" {
        $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
            -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

        $result.Success | Should -BeTrue
        $result.Status | Should -Match '^ok'
        $result.BackupPath | Should -Be (Join-Path $script:appDir 'firefox_1.0_20260101-010101.intunewin')
        Test-Path -LiteralPath $result.BackupPath | Should -BeTrue
    }

    It "copies the file byte for byte" {
        $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
            -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

        (Get-Item -LiteralPath $result.BackupPath).Length | Should -Be (Get-Item -LiteralPath $script:src).Length
    }

    It "creates the whole directory chain when nothing exists yet" {
        $deep = Join-Path $script:root "a\b\c"
        $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $deep `
            -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

        $result.Success | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $deep "firefox") | Should -BeTrue
    }

    It "returns an object carrying the full documented contract" {
        $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
            -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

        foreach ($prop in @('AppId', 'Success', 'Status', 'BackupPath', 'Removed', 'SourceRemoved', 'Log')) {
            $result.PSObject.Properties[$prop] | Should -Not -BeNullOrEmpty -Because "Wait-YardstickBackup reads $prop"
        }
        $result.AppId | Should -Be 'firefox'
        $result.Log.Count | Should -BeGreaterThan 0
    }

    Context "retention" {
        BeforeEach {
            New-Item -ItemType Directory -Path $script:appDir -Force | Out-Null
        }

        It "keeps the newest by embedded timestamp, not by last write time" {
            # LastWriteTime is written in the opposite order to the embedded
            # timestamps, so a naive sort on file metadata fails this test.
            $stamps = @('20260105-010101', '20260101-010101', '20260104-010101', '20260102-010101', '20260103-010101')
            for ($i = 0; $i -lt $stamps.Count; $i++) {
                $f = Join-Path $script:appDir "firefox_1.0_$($stamps[$i]).intunewin"
                Set-Content -LiteralPath $f -Value "old"
                (Get-Item -LiteralPath $f).LastWriteTime = (Get-Date).AddDays(-$i)
            }

            $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260106-010101.intunewin' -VersionsToKeep 3

            $result.Status | Should -Match '3 kept'
            $kept = @(Get-ChildItem -LiteralPath $script:appDir -File | Where-Object Extension -eq '.intunewin' | Select-Object -ExpandProperty Name | Sort-Object)
            $kept | Should -Be @(
                'firefox_1.0_20260104-010101.intunewin',
                'firefox_1.0_20260105-010101.intunewin',
                'firefox_1.0_20260106-010101.intunewin'
            )
        }

        It "keeps only the file just written when VersionsToKeep is 1" {
            foreach ($s in @('20260101-010101', '20260102-010101')) {
                Set-Content -LiteralPath (Join-Path $script:appDir "firefox_1.0_$s.intunewin") -Value "old"
            }

            Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260103-010101.intunewin' -VersionsToKeep 1 | Out-Null

            $kept = @(Get-ChildItem -LiteralPath $script:appDir -File | Where-Object Extension -eq '.intunewin')
            $kept.Count | Should -Be 1
            $kept[0].Name | Should -Be 'firefox_1.0_20260103-010101.intunewin'
        }

        It "prunes nothing when VersionsToKeep is <Keep>" -ForEach @(
            @{ Keep = 0 }, @{ Keep = -1 }
        ) {
            foreach ($s in @('20260101-010101', '20260102-010101')) {
                Set-Content -LiteralPath (Join-Path $script:appDir "firefox_1.0_$s.intunewin") -Value "old"
            }

            $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260103-010101.intunewin' -VersionsToKeep $Keep

            $result.Status | Should -Match 'retention disabled'
            @(Get-ChildItem -LiteralPath $script:appDir -File | Where-Object Extension -eq '.intunewin').Count | Should -Be 3
        }

        It "sorts unparseable names by last write time without throwing" {
            Set-Content -LiteralPath (Join-Path $script:appDir "legacy-a.intunewin") -Value "old"
            Set-Content -LiteralPath (Join-Path $script:appDir "legacy-b.intunewin") -Value "old"
            Set-Content -LiteralPath (Join-Path $script:appDir "firefox_1.0_20260101-010101.intunewin") -Value "old"

            $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260102-010101.intunewin' -VersionsToKeep 2

            $result.Success | Should -BeTrue
            @(Get-ChildItem -LiteralPath $script:appDir -File | Where-Object Extension -eq '.intunewin').Count | Should -Be 2
        }

        It "never deletes files that are not .intunewin backups" {
            # Guards against -Filter "*.intunewin", whose 8.3 short-name matching
            # would sweep up names it should not.
            foreach ($s in @('20260101-010101', '20260102-010101', '20260103-010101')) {
                Set-Content -LiteralPath (Join-Path $script:appDir "firefox_1.0_$s.intunewin") -Value "old"
            }
            Set-Content -LiteralPath (Join-Path $script:appDir "notes.txt") -Value "keep"
            Set-Content -LiteralPath (Join-Path $script:appDir "recent.intunewin.tmp") -Value "keep"

            Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260104-010101.intunewin' -VersionsToKeep 1 | Out-Null

            Test-Path -LiteralPath (Join-Path $script:appDir "notes.txt") | Should -BeTrue
            Test-Path -LiteralPath (Join-Path $script:appDir "recent.intunewin.tmp") | Should -BeTrue
        }

        It "sweeps abandoned .tmp files older than the cutoff but keeps recent ones" {
            $stale = Join-Path $script:appDir "abandoned.intunewin.tmp"
            Set-Content -LiteralPath $stale -Value "partial"
            (Get-Item -LiteralPath $stale).LastWriteTime = (Get-Date).AddHours(-48)
            $fresh = Join-Path $script:appDir "inflight.intunewin.tmp"
            Set-Content -LiteralPath $fresh -Value "partial"

            Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin' -StaleTmpHours 24 | Out-Null

            Test-Path -LiteralPath $stale | Should -BeFalse
            Test-Path -LiteralPath $fresh | Should -BeTrue
        }
    }

    Context "failure handling" {
        It "reports a missing source without throwing" {
            # Assigned with the script: modifier because a Should -Not -Throw
            # scriptblock runs in a child scope.
            $script:result = $null
            { $script:result = Invoke-YardstickBackupCopy -SourcePath (Join-Path $TestDrive "gone.intunewin") `
                -BackupRoot $script:root -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin' } |
                Should -Not -Throw

            $script:result.Success | Should -BeFalse
            $script:result.Status | Should -Match 'source missing'
        }

        It "reports an unreachable backup root without throwing" {
            # A directory path underneath an existing file can never be created.
            $blocker = Join-Path $TestDrive "blocker.txt"
            Set-Content -LiteralPath $blocker -Value "i am a file"

            $script:result = $null
            { $script:result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot (Join-Path $blocker "sub") `
                -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin' } | Should -Not -Throw

            $script:result.Success | Should -BeFalse
            $script:result.Status | Should -Match '^failed'
            $script:result.Log.Count | Should -BeGreaterThan 0
        }

        It "leaves a .tmp and no .intunewin when the commit rename fails" {
            Mock -ModuleName YardstickSupport Move-Item { throw "interrupted" }

            $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

            $result.Success | Should -BeFalse
            # An interrupted copy must never leave something retention counts as a backup.
            Test-Path -LiteralPath (Join-Path $script:appDir 'firefox_1.0_20260101-010101.intunewin') | Should -BeFalse
        }

        It "detects a truncated copy before committing it" {
            Mock -ModuleName YardstickSupport Copy-Item {
                Set-Content -LiteralPath $Destination -Value "short" -NoNewline
            }

            $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

            $result.Success | Should -BeFalse
            $result.Status | Should -Match '^failed'
            Test-Path -LiteralPath (Join-Path $script:appDir 'firefox_1.0_20260101-010101.intunewin') | Should -BeFalse
        }
    }

    Context "source removal" {
        It "deletes the source after a successful copy" {
            $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin' -RemoveSource

            $result.SourceRemoved | Should -BeTrue
            Test-Path -LiteralPath $script:src | Should -BeFalse
        }

        It "deletes the source even when the copy failed" {
            # Invoke-Cleanup would delete it at the end of the run anyway, so
            # holding on to it would only fill up the Published folder.
            $blocker = Join-Path $TestDrive "blocker2.txt"
            Set-Content -LiteralPath $blocker -Value "i am a file"

            $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot (Join-Path $blocker "sub") `
                -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin' -RemoveSource

            $result.Success | Should -BeFalse
            $result.SourceRemoved | Should -BeTrue
            Test-Path -LiteralPath $script:src | Should -BeFalse
        }

        It "leaves the source alone when RemoveSource is not requested" {
            $result = Invoke-YardstickBackupCopy -SourcePath $script:src -BackupRoot $script:root `
                -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

            $result.SourceRemoved | Should -BeFalse
            Test-Path -LiteralPath $script:src | Should -BeTrue
        }
    }
}

Describe "Start-YardstickBackup and Wait-YardstickBackup" {
    BeforeEach {
        $script:root = Join-Path $TestDrive "jobbackups"
        $script:src = Join-Path $TestDrive "jobpackage.intunewin"
        Set-Content -LiteralPath $script:src -Value ("x" * 65536) -NoNewline
    }

    AfterEach {
        Wait-YardstickBackup -TimeoutSeconds 60 | Out-Null
        Remove-Item -LiteralPath $script:root -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:src -Force -ErrorAction SilentlyContinue
    }

    It "is a no-op when nothing is registered, however many times it is called" {
        { Wait-YardstickBackup -TimeoutSeconds 5 } | Should -Not -Throw
        Wait-YardstickBackup -TimeoutSeconds 5 | Should -BeNullOrEmpty
        Get-YardstickBackupInFlight | Should -BeNullOrEmpty
    }

    It "copies on a background thread and deletes the source when done" {
        $name = 'firefox_1.0_20260101-010101.intunewin'
        Start-YardstickBackup -SourcePath $script:src -BackupRoot $script:root -AppId 'firefox' -FileName $name

        $results = Wait-YardstickBackup -TimeoutSeconds 120

        $results.Count | Should -Be 1
        $results[0].Success | Should -BeTrue
        Test-Path -LiteralPath (Join-Path (Join-Path $script:root 'firefox') $name) | Should -BeTrue
        Test-Path -LiteralPath $script:src | Should -BeFalse
    }

    It "runs the copy inside the thread rather than silently failing to find it" {
        # A thread job runspace inherits nothing from its parent, so this asserts
        # the module really was importable inside the job: a real result object
        # comes back, not a command-not-found ErrorRecord.
        Start-YardstickBackup -SourcePath $script:src -BackupRoot $script:root `
            -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

        $results = Wait-YardstickBackup -TimeoutSeconds 120

        $results[0] | Should -BeOfType [PSCustomObject]
        $results[0] | Should -Not -BeOfType [System.Management.Automation.ErrorRecord]
        $results[0].Status | Should -Match '^ok'
        $results[0].Log.Count | Should -BeGreaterThan 0
    }

    It "reports the source as in flight until it is drained" {
        Start-YardstickBackup -SourcePath $script:src -BackupRoot $script:root `
            -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'

        Wait-YardstickBackup -TimeoutSeconds 120 | Out-Null
        Get-YardstickBackupInFlight | Should -BeNullOrEmpty
    }

    It "records the outcome on the matching successful application entry" {
        Initialize-ApplicationTracker
        Add-SuccessfulApplication -ApplicationId 'firefox' -DisplayName 'Mozilla Firefox' -Version '1.0'

        Start-YardstickBackup -SourcePath $script:src -BackupRoot $script:root `
            -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'
        Wait-YardstickBackup -TimeoutSeconds 120 | Out-Null

        $tracked = (Get-Module YardstickSupport) | ForEach-Object { & $_ { $Script:SuccessfulApplications } }
        $tracked[0].BackupStatus | Should -Match '^ok'
    }

    It "leaves no job objects behind in the session" {
        Start-YardstickBackup -SourcePath $script:src -BackupRoot $script:root `
            -AppId 'firefox' -FileName 'firefox_1.0_20260101-010101.intunewin'
        Wait-YardstickBackup -TimeoutSeconds 120 | Out-Null

        @(Get-Job -Name 'YardstickBackup_*' -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
}
