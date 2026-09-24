BeforeAll {
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickGraph.psm1" -Force
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickIntune.psm1" -Force
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

Describe 'Yardstick Intune install time settings' {
    # An 11:00 PM deployment moved forward 14 days came back as 4:00 AM the next
    # day, because the wall clock Intune stores was being treated as an instant
    # and converted into UTC by the build host's offset.

    It 'writes an install time as the wall clock it was given' {
        InModuleScope YardstickIntune {
            # Same wall clock, different Kind. Both must serialize identically:
            # a value that changes when .Kind changes is being converted.
            $local = [datetime]::SpecifyKind([datetime]'2026-08-25 23:00', [System.DateTimeKind]::Local)
            $utc = [datetime]::SpecifyKind([datetime]'2026-08-25 23:00', [System.DateTimeKind]::Utc)
            ConvertTo-YardstickGraphDate $local | Should -Be '2026-08-25T23:00:00.000Z'
            ConvertTo-YardstickGraphDate $utc | Should -Be '2026-08-25T23:00:00.000Z'
        }
    }

    It 'still converts to UTC for values that name an instant' {
        InModuleScope YardstickIntune {
            $local = [datetime]::SpecifyKind([datetime]'2026-08-25 23:00', [System.DateTimeKind]::Local)
            $expected = $local.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [cultureinfo]::InvariantCulture)
            ConvertTo-YardstickGraphDate $local -AsUtc | Should -Be $expected
        }
    }

    It 'uses UTC for detection rule dates' {
        $value = [datetime]::SpecifyKind([datetime]'2026-08-25 23:00', [System.DateTimeKind]::Local)
        $rule = New-YardstickWin32AppDetectionRuleFile -DateModified -Path 'C:\App' -FileOrFolder 'app.exe' `
            -Operator greaterThanOrEqual -DateTimeValue $value
        $rule.detectionValue | Should -Be ($value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fffZ', [cultureinfo]::InvariantCulture))
    }

    It 'reads back every shape Graph reports a date in without shifting it' {
        # PowerShell 7 ConvertFrom-Json yields Kind Utc, 5.1 yields the raw
        # string, and a DateTimeOffset carries its offset separately. Left
        # alone, a [datetime] cast of the string reinterprets it into the host
        # timezone and the three disagree about what .Hour means.
        $fromJson = ('{"d":"2026-08-25T23:00:00.0000000Z"}' | ConvertFrom-Json).d
        $inputs = @($fromJson, '2026-08-25T23:00:00.0000000Z', [datetimeoffset]'2026-08-25T23:00:00.0000000Z')
        foreach ($value in $inputs) {
            $wallClock = ConvertTo-YardstickWallClock $value
            $wallClock.Hour | Should -Be 23
            $wallClock.Minute | Should -Be 0
            $wallClock.Kind | Should -Be ([System.DateTimeKind]::Unspecified)
        }
    }

    It 'round-trips a stored install time through read and write unchanged' {
        InModuleScope YardstickIntune {
            $stored = ('{"d":"2026-08-25T23:00:00.0000000Z"}' | ConvertFrom-Json).d
            $wallClock = ConvertTo-YardstickWallClock $stored
            $rebased = [datetime]::SpecifyKind((Get-Date).Date.AddDays(14), [System.DateTimeKind]::Unspecified).AddHours($wallClock.Hour).AddMinutes($wallClock.Minute)
            $body = New-YardstickAssignmentBody -TargetType '#microsoft.graph.groupAssignmentTarget' -GroupID 'g' `
                -Intent required -AvailableTime $rebased -UseLocalTime $true
            $body.settings.installTimeSettings.useLocalTime | Should -BeTrue
            # 23:00 in, 23:00 out, on any host. The offset applies to the date
            # alone - it must not leak into the time of day.
            $body.settings.installTimeSettings.startDateTime | Should -Match 'T23:00:00\.000Z$'
            $body.settings.installTimeSettings.startDateTime | Should -BeLike "$((Get-Date).Date.AddDays(14).ToString('yyyy-MM-dd'))*"
        }
    }
}

Describe 'Yardstick Intune relationships' {    It 'preserves supersedence while replacing dependencies' {
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

Describe 'Yardstick Intune content upload' {
    BeforeAll {
        # The body Azure returned for the 2026-09-15 Evernote failure, byte order
        # mark and all.
        $Script:SasRejectionBody = "$([char]0xFEFF)AuthenticationFailedServer failed to authenticate the request. Make sure the value of Authorization header is formed correctly including the signature.`nRequestId:c9c11ea2-601e-0023-5630-454314000000`nTime:2026-09-15T16:35:40.2987536ZSAS identifier cannot be found for specified signed identifier"

        # Defined in the module's own scope so the mocked Invoke-WebRequest can
        # reach it; a helper declared out here is invisible from in there.
        InModuleScope YardstickIntune {
            function Script:New-HttpErrorRecord {
                param([int]$StatusCode, [string]$Body)
                $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$StatusCode)
                $exception = [Microsoft.PowerShell.Commands.HttpResponseException]::new(
                    "Response status code does not indicate success: $StatusCode.", $response)
                $record = [System.Management.Automation.ErrorRecord]::new(
                    $exception, 'WebCmdletWebResponseException', [System.Management.Automation.ErrorCategory]::InvalidOperation, $null)
                if ($Body) { $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($Body) }
                return $record
            }
        }
    }

    Context 'classifying blob failures' {
        It 'recognises a rejected SAS in the body Azure actually sends' {
            InModuleScope YardstickIntune -Parameters @{ SasBody = $Script:SasRejectionBody } {
                param($SasBody)
                Test-YardstickSasRejection -Detail $SasBody -StatusCode 403 | Should -BeTrue
            }
        }

        It 'leaves a plain 403 alone so it stays fatal' {
            InModuleScope YardstickIntune {
                Test-YardstickSasRejection -Detail 'This request is not authorized to perform this operation.' -StatusCode 403 | Should -BeFalse
            }
        }

        It 'does not read a 404 as a SAS problem however it is worded' {
            InModuleScope YardstickIntune {
                Test-YardstickSasRejection -Detail 'AuthenticationFailed' -StatusCode 404 | Should -BeFalse
            }
        }

        It 'prefers the response body over the status line and strips the byte order mark' {
            InModuleScope YardstickIntune -Parameters @{ SasBody = $Script:SasRejectionBody } {
                param($SasBody)
                $record = New-HttpErrorRecord -StatusCode 403 -Body $SasBody
                $detail = Get-YardstickHttpErrorDetail -ErrorRecord $record
                $detail | Should -BeLike 'AuthenticationFailed*'
                $detail | Should -Match 'SAS identifier cannot be found'
            }
        }

        It 'falls back to the status line when there is no response body' {
            InModuleScope YardstickIntune {
                $record = New-HttpErrorRecord -StatusCode 500
                Get-YardstickHttpErrorDetail -ErrorRecord $record | Should -Match 'Response status code does not indicate success'
            }
        }
    }

    Context 'recovering from a rejected SAS' {
        It 'renews the signature and replays the request' {
            InModuleScope YardstickIntune -Parameters @{ SasBody = $Script:SasRejectionBody } {
                param($SasBody)
                $Script:calls = 0
                Mock Invoke-WebRequest {
                    $Script:calls++
                    if ($Uri -like 'https://blob/stale*') { throw (New-HttpErrorRecord -StatusCode 403 -Body $SasBody) }
                }
                Mock Update-YardstickIntuneUploadUri { [pscustomobject]@{ Uri = 'https://blob/fresh?sig=new'; ExpiresAt = $null } }
                Mock Write-YardstickIntuneLog {}

                $stale = [pscustomobject]@{ Uri = 'https://blob/stale?sig=old'; ExpiresAt = $null }
                $result = Invoke-YardstickBlobRequestWithRetry -Method Put -UploadUri $stale -FileResource 'files/1' `
                    -Query 'comp=block&blockid=AAA' -Body ([byte[]](1, 2, 3)) -Label 'Block 1 of 2'

                $result.Uri | Should -Be 'https://blob/fresh?sig=new'
                $Script:calls | Should -Be 2
                Should -Invoke Update-YardstickIntuneUploadUri -Times 1 -Exactly
            }
        }

        It 'gives up on a plain 403 without renewing' {
            InModuleScope YardstickIntune {
                Mock Invoke-WebRequest { throw (New-HttpErrorRecord -StatusCode 403 -Body 'This request is not authorized.') }
                Mock Update-YardstickIntuneUploadUri {}
                Mock Write-YardstickIntuneLog {}

                $upload = [pscustomobject]@{ Uri = 'https://blob/one?sig=old'; ExpiresAt = $null }
                { Invoke-YardstickBlobRequestWithRetry -Method Put -UploadUri $upload -FileResource 'files/1' `
                        -Query 'comp=blocklist' -Body ([byte[]](1)) -Label 'Block list commit' } |
                    Should -Throw '*Block list commit failed after 1 attempt(s)*'
                Should -Invoke Update-YardstickIntuneUploadUri -Times 0 -Exactly
            }
        }

        It 'names the failing block so the log says which one died' {
            InModuleScope YardstickIntune {
                Mock Invoke-WebRequest { throw (New-HttpErrorRecord -StatusCode 403 -Body 'This request is not authorized.') }
                Mock Write-YardstickIntuneLog {}
                $path = Join-Path $TestDrive 'content.bin'
                [io.file]::WriteAllBytes($path, [byte[]]::new(12))

                $upload = [pscustomobject]@{ Uri = 'https://blob/one?sig=old'; ExpiresAt = $null }
                { Send-YardstickIntuneContentBlob -FilePath $path -UploadUri $upload -FileResource 'files/1' -ChunkSize 5 } |
                    Should -Throw '*Block 1 of 3 failed*'
            }
        }
    }

    Context 'renewing the upload signature' {
        It 'retries a failed renewal and succeeds on the second attempt' {
            InModuleScope YardstickIntune {
                $Script:waits = 0
                Mock Invoke-YardstickGraphRequest {}
                Mock Wait-YardstickIntuneFileProcessing {
                    $Script:waits++
                    if ($Script:waits -eq 1) { throw "Intune file operation 'azureStorageUriRenewal' failed." }
                    [pscustomobject]@{ azureStorageUri = 'https://blob/fresh?sig=new'; azureStorageUriExpirationDateTime = $null }
                }
                Mock Write-YardstickIntuneLog {}
                Mock Start-Sleep {}

                $result = Update-YardstickIntuneUploadUri -FileResource 'files/1'

                $result.Uri | Should -Be 'https://blob/fresh?sig=new'
                $Script:waits | Should -Be 2
                # Each attempt must re-POST renewUpload; the service state machine
                # only restarts when it is asked again.
                Should -Invoke Invoke-YardstickGraphRequest -Times 2 -Exactly
            }
        }

        It 'still throws when every renewal attempt fails' {
            InModuleScope YardstickIntune {
                Mock Invoke-YardstickGraphRequest {}
                Mock Wait-YardstickIntuneFileProcessing { throw "Intune file operation 'azureStorageUriRenewal' failed." }
                Mock Write-YardstickIntuneLog {}
                Mock Start-Sleep {}

                { Update-YardstickIntuneUploadUri -FileResource 'files/1' } |
                    Should -Throw "*azureStorageUriRenewal*"
                Should -Invoke Wait-YardstickIntuneFileProcessing -Times 3 -Exactly
            }
        }

        It 'does not retry when the first attempt succeeds' {
            InModuleScope YardstickIntune {
                Mock Invoke-YardstickGraphRequest {}
                Mock Wait-YardstickIntuneFileProcessing {
                    [pscustomobject]@{ azureStorageUri = 'https://blob/fresh?sig=new'; azureStorageUriExpirationDateTime = $null }
                }
                Mock Write-YardstickIntuneLog {}

                Update-YardstickIntuneUploadUri -FileResource 'files/1' | Out-Null
                Should -Invoke Invoke-YardstickGraphRequest -Times 1 -Exactly
            }
        }
    }

    Context 'scheduling renewals' {
        It 'reads the expiry Intune published as UTC' {
            InModuleScope YardstickIntune {
                $file = [pscustomobject]@{
                    azureStorageUri                   = 'https://blob/one?sig=x'
                    azureStorageUriExpirationDateTime = '2026-09-15T17:35:00.000Z'
                }
                $upload = ConvertTo-YardstickUploadUri -File $file
                $upload.Uri | Should -Be 'https://blob/one?sig=x'
                $upload.ExpiresAt.Kind | Should -Be ([datetimekind]::Utc)
                $upload.ExpiresAt.Hour | Should -Be 17
            }
        }

        It 'renews ahead of the published expiry rather than on a fixed timer' {
            InModuleScope YardstickIntune {
                $upload = [pscustomobject]@{ Uri = 'https://blob/one'; ExpiresAt = [datetime]::UtcNow.AddMinutes(40) }
                $renewAt = Get-YardstickUploadUriRenewalTime -UploadUri $upload
                # Five minutes of margin ahead of a 40 minute window.
                [math]::Round($renewAt.Subtract([datetime]::UtcNow).TotalMinutes) | Should -Be 35
            }
        }

        It 'falls back to a fixed lifetime when the service omits the expiry' {
            InModuleScope YardstickIntune {
                $upload = ConvertTo-YardstickUploadUri -File ([pscustomobject]@{ azureStorageUri = 'https://blob/one' })
                $upload.ExpiresAt | Should -BeNullOrEmpty
                $renewAt = Get-YardstickUploadUriRenewalTime -UploadUri $upload
                [math]::Round($renewAt.Subtract([datetime]::UtcNow).TotalMinutes) | Should -Be 7
            }
        }

        It 'never schedules a renewal in the past when the window is already short' {
            InModuleScope YardstickIntune {
                $upload = [pscustomobject]@{ Uri = 'https://blob/one'; ExpiresAt = [datetime]::UtcNow.AddMinutes(1) }
                Get-YardstickUploadUriRenewalTime -UploadUri $upload | Should -BeGreaterThan ([datetime]::UtcNow)
            }
        }
    }
}
