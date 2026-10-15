BeforeAll {
    $script:modulePath = [IO.Path]::GetFullPath("$PSScriptRoot\..\..\Modules\YardstickLog.psm1")
    Import-Module $script:modulePath -Force

    function ConvertFrom-Fixture {
        param([string]$Text)
        $run = Get-YardstickLogRun -Lines ($Text -split "`r?`n") -Index -1
        $entries = ConvertTo-YardstickLogEntry -Lines $run.Lines
        ConvertFrom-YardstickLog -Entries $entries
    }

    # What Send-YardstickLogReport.ps1 takes from each app, in a comparable form.
    function ConvertTo-AppSummary {
        param($Apps)
        foreach ($app in $Apps) {
            $failures = @($app.Failures)
            if ($failures.Count -eq 0 -and -not $app.Succeeded -and $app.ErrorLines.Count -gt 0) {
                $failures = @([PSCustomObject]@{ ErrorMessage = $app.ErrorLines[-1] })
            }
            [PSCustomObject]@{
                Id          = $app.Id
                DisplayName = $app.DisplayName
                Version     = $app.Version
                Action      = $app.Action
                Succeeded   = $app.Succeeded
                Stages      = @($failures | ForEach-Object {
                    if ($_.Stage) { $_.Stage } else { Get-YardstickFailureStage -ErrorMessage $_.ErrorMessage }
                }) -join ','
                Dependents  = @($app.Dependents.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ';'
                AutoParts   = @($app.AutoParts) -join ','
                Backup      = $app.BackupStatus
            } | ConvertTo-Json -Compress
        }
    }

    # The same run, as the runner logged it before and after the log cleanup.
    $script:oldLog = @'
#######################################################
LOGGING STARTED AT 10/01/2026 08:00:00
#######################################################
10/01/2026 08:00:01 - Starting update for 7zip.arm64...
10/01/2026 08:00:02 - Retrieving all versions of application with display name: 7-Zip (ARM64)
10/01/2026 08:00:03 - Checking if 7-Zip (ARM64) 26.04 is a new version...
10/01/2026 08:00:04 - Attaching supersedence (Update) from 7-Zip (ARM64) 26.04 to 2 target(s)
10/01/2026 08:00:05 - Auto-update (enabled) is set on 3 assignment(s) for app-1
10/01/2026 08:00:06 - Updated dependent app Bundle to reference 7-Zip (ARM64) 26.04.
10/01/2026 08:00:07 - Tracked successful application: 7-Zip (ARM64) 26.04
10/01/2026 08:00:08 - Backed up 7zip.arm64 to X:\Backup\7zip.arm64\a.intunewin
10/01/2026 08:00:08 - Pruned old backup b.intunewin
10/01/2026 08:00:09 - Starting update for putty...
10/01/2026 08:00:10 - Retrieving all versions of application with display name: PuTTY
10/01/2026 08:00:11 - Checking if PuTTY 0.83 is a new version...
10/01/2026 08:00:12 - PuTTY 0.83 already exists in Intune. Running maintenance sweep.
10/01/2026 08:00:13 - Pinning PuTTY 0.83 (p-1) as DETECT anchor for PuTTY
10/01/2026 08:00:14 - Tracked successful application: PuTTY 0.83
10/01/2026 08:00:15 - Starting update for firefox...
10/01/2026 08:00:16 - Tracked failed application: firefox - Error downloading file: 404 Not Found
10/01/2026 08:00:16 - ERROR: Error downloading file: 404 Not Found
10/01/2026 08:00:17 - Starting update for vlc...
10/01/2026 08:00:18 - Failed to update dependent app Media Pack on attempt 1: timeout
10/01/2026 08:00:19 - ERROR: Something broke
   at line 3
10/01/2026 08:00:20 - WARNING: Backup of vlc failed: disk full
'@

    $script:newLog = @'
#######################################################
LOGGING STARTED AT 10/15/2026 08:00:00
#######################################################
10/15/2026 08:00:00 - Processing 4 recipe(s).
10/15/2026 08:00:01 - Starting update for 7zip.arm64...
10/15/2026 08:00:02 - [7zip.arm64] Retrieving all versions of application with display name: 7-Zip (ARM64)
10/15/2026 08:00:03 - [7zip.arm64] Checking if 7-Zip (ARM64) 26.04 is a new version...
10/15/2026 08:00:04 - [7zip.arm64] Attaching supersedence (Update) from 7-Zip (ARM64) 26.04 to 2 target(s)
10/15/2026 08:00:05 - [7zip.arm64] Auto-update (enabled) is set on 3 assignment(s) for app-1.
10/15/2026 08:00:06 - [7zip.arm64] Updated dependent app Bundle to reference 7-Zip (ARM64) 26.04.
10/15/2026 08:00:07 - [7zip.arm64] Done: 7-Zip (ARM64) 26.04 (Updated).
10/15/2026 08:00:08 - [7zip.arm64] Backed up 7zip.arm64 to X:\Backup\7zip.arm64\a.intunewin
10/15/2026 08:00:08 - [7zip.arm64] Pruned old backup b.intunewin.
10/15/2026 08:00:09 - Starting update for putty...
10/15/2026 08:00:10 - [putty] Retrieving all versions of application with display name: PuTTY
10/15/2026 08:00:11 - [putty] Checking if PuTTY 0.83 is a new version...
10/15/2026 08:00:12 - [putty] PuTTY 0.83 is already in Intune. Running maintenance sweep.
10/15/2026 08:00:13 - [putty] Pinning PuTTY 0.83 (p-1) as DETECT anchor for PuTTY
10/15/2026 08:00:14 - [putty] Done: PuTTY 0.83 (Maintained).
10/15/2026 08:00:15 - Starting update for firefox...
10/15/2026 08:00:16 - [firefox] ERROR: Download failed: 404 Not Found
10/15/2026 08:00:17 - Starting update for vlc...
10/15/2026 08:00:18 - [vlc] WARNING: Failed to update dependent app Media Pack on attempt 1: timeout
10/15/2026 08:00:19 - [vlc] ERROR: Something broke
   at line 3
10/15/2026 08:00:20 - [vlc] WARNING: Backup of vlc failed: disk full
10/15/2026 08:00:21 - Finished processing 4 recipe(s).
'@
}

Describe "ConvertFrom-YardstickLog" {
    It "reads the same results from the old and the new log formats" {
        $old = @(ConvertTo-AppSummary (ConvertFrom-Fixture $script:oldLog))
        $new = @(ConvertTo-AppSummary (ConvertFrom-Fixture $script:newLog))

        $old.Count | Should -Be 4
        $new | Should -Be $old
    }

    Context "new format" {
        BeforeAll {
            $script:apps = @(ConvertFrom-Fixture $script:newLog)
        }

        It "reads the name, version and action from the Done line" {
            $app = $script:apps | Where-Object Id -eq '7zip.arm64'
            $app.Succeeded | Should -BeTrue
            $app.DisplayName | Should -Be '7-Zip (ARM64)'
            $app.Version | Should -Be '26.04'
            $app.Action | Should -Be 'Updated'
            $app.AutoParts | Should -Be @('supersedes 2', 'auto-update on 3 assignment(s)')
            $app.BackupStatus | Should -Be 'ok (1 pruned)'
        }

        It "reads a maintenance sweep" {
            $app = $script:apps | Where-Object Id -eq 'putty'
            $app.Action | Should -Be 'Maintained'
            $app.AutoParts | Should -Be @('anchor: pinned')
        }

        It "takes the stage from the failure line" {
            $app = $script:apps | Where-Object Id -eq 'firefox'
            $app.Succeeded | Should -BeFalse
            $app.Failures.Count | Should -Be 1
            $app.Failures[0].Id | Should -Be 'firefox'
            $app.Failures[0].Stage | Should -Be 'Download'
            $app.Failures[0].ErrorMessage | Should -Be '404 Not Found'
        }

        It "reads a dependent-app failure and a backup warning that carry the Warning level" {
            $app = $script:apps | Where-Object Id -eq 'vlc'
            $app.Dependents['Media Pack'] | Should -Be 'Failed (timeout)'
            $app.BackupStatus | Should -Be 'Backup of vlc failed: disk full' -Because 'the report shows the warning without its prefix'
            $app.ErrorLines[-1] | Should -Be "Something broke`n   at line 3"
        }
    }

    It "does not stop a stage name at a shorter stage it starts with" {
        $apps = @(ConvertFrom-Fixture @'
10/15/2026 08:00:01 - Starting update for foo...
10/15/2026 08:00:02 - [foo] ERROR: Download Script failed: exit code 1
'@)
        $apps[0].Failures[0].Stage | Should -Be 'Download Script'
        $apps[0].Failures[0].ErrorMessage | Should -Be 'exit code 1'
    }

    It "leaves a bracketed prefix alone when it does not name the current app" {
        $apps = @(ConvertFrom-Fixture @'
10/15/2026 08:00:01 - Starting update for foo...
10/15/2026 08:00:02 - [Retrieve applications for Foo] Done: Bar 1.0 (Updated).
'@)
        $apps[0].Succeeded | Should -BeFalse
    }
}

Describe "Get-YardstickFailureStage" {
    It "maps old-format messages to their stage" {
        Get-YardstickFailureStage -ErrorMessage 'Error downloading file: 404' | Should -Be 'Download'
        Get-YardstickFailureStage -ErrorMessage 'Failed to upload application to Intune: 500' | Should -Be 'Intune Upload'
        Get-YardstickFailureStage -ErrorMessage 'Something else' | Should -Be 'Processing'
    }
}
