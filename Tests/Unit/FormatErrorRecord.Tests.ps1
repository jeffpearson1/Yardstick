BeforeAll {
    Get-Module -Name YardstickSupport -All | Remove-Module -Force -ErrorAction SilentlyContinue
    Import-Module "$PSScriptRoot\..\..\Modules\YardstickSupport.psm1" -Force

    function New-WebError([int]$Status, [string]$Body) {
        $message = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$Status)
        $exception = [Microsoft.PowerShell.Commands.HttpResponseException]::new("Response status code does not indicate success: $Status ($($message.ReasonPhrase)).", $message)
        $record = [System.Management.Automation.ErrorRecord]::new($exception, 'WebCmdletWebResponseException', 'InvalidOperation', $null)
        if ($Body) { $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($Body) }
        $record
    }
}

Describe 'Format-YardstickErrorRecord' {
    It 'shows the status and the text of an HTML outage page (hxd, 2026-11-23)' {
        $body = "`n`n<html><head><title>The page is temporarily unavailable</title><style>body { font-family: Tahoma; }</style></head>`n<body>`n`n`n<p>The page you are looking for is temporarily unavailable.</p>`n<p>Please try again later.</p></body></html>"
        $text = Format-YardstickErrorRecord (New-WebError 502 $body)
        $text | Should -BeLike 'Response status code does not indicate success: 502 (Bad Gateway).*temporarily unavailable*'
        $text | Should -Not -Match '<|\r|\n'
        $text | Should -Not -Match 'font-family'
    }

    It 'caps a long body' {
        $text = Format-YardstickErrorRecord (New-WebError 500 ('x' * 1000))
        $text.Length | Should -BeLessThan 320
        $text | Should -BeLike '*...'
    }

    It 'returns the message alone when there is no body' {
        $record = [System.Management.Automation.ErrorRecord]::new([System.Exception]::new("Release 'v2.3.0' must have exactly one file matching 'x'."), 'x', 'NotSpecified', $null)
        Format-YardstickErrorRecord $record | Should -Be "Release 'v2.3.0' must have exactly one file matching 'x'."
    }

    It 'does not repeat a body the message already contains' {
        $record = [System.Management.Automation.ErrorRecord]::new([System.Exception]::new('boom'), 'x', 'NotSpecified', $null)
        $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('boom')
        Format-YardstickErrorRecord $record | Should -Be 'boom'
    }

    It 'passes plain strings through on one line' {
        Format-YardstickErrorRecord "a`n  b" | Should -Be 'a b'
    }
}

Describe 'Error text at the failure sites' {
    It 'formats every recipe-stage failure with Format-YardstickErrorRecord' {
        $text = Get-Content "$PSScriptRoot\..\..\Private\Invoke-YardstickRun.ps1" -Raw
        foreach ($stage in 'Pre-Download Script', 'Download Script', 'Download', 'Post-Download Script', 'Intune Upload', 'Post-Run Script') {
            $text | Should -Match ([regex]::Escape("-ErrorMessage (Format-YardstickErrorRecord `$_) -FailureStage `"$stage`""))
        }
    }
}
