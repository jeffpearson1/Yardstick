BeforeAll {
    # The breaker lives in Invoke-YardstickRun.ps1, which cannot be dot-sourced (its
    # top level runs a whole update pass), so extract it via the AST. It reads the
    # runner's $IntuneBreaker and $Script:DisplayName, which each test sets.
    $runnerPath = "$PSScriptRoot\..\..\Private\Invoke-YardstickRun.ps1"
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($runnerPath, [ref]$null, [ref]$null)
    foreach ($name in 'Test-YardstickTransientIntuneError', 'Register-YardstickIntuneLookupResult', 'Test-YardstickIntuneBreaker', 'Test-YardstickIntuneAvailable') {
        $fn = $ast.Find({ param($n)
            $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if (-not $fn) { throw "$name not found in $runnerPath" }
        . ([scriptblock]::Create($fn.Extent.Text))
    }

    function Write-Log { param([string]$Content, [string]$Level) }
    function Invoke-YardstickGraphRequest { [CmdletBinding()] param($ApiVersion, $Resource, $MaxAttempts) }

    function New-HttpError([int]$Status) {
        $message = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$Status)
        $http = [Microsoft.PowerShell.Commands.HttpResponseException]::new("Response status code does not indicate success: $Status.", $message)
        # Shaped like Get-SameAppAllVersions' error: the HTTP error is the inner one.
        $wrapped = [System.Exception]::new("Could not retrieve the versions of App from Intune: $($http.Message)", $http)
        [System.Management.Automation.ErrorRecord]::new($wrapped, 'x', 'NotSpecified', $null)
    }

    function New-Breaker {
        @{ Threshold = 2; CoolDownMinutes = 10; Consecutive = 0; OpenUntil = $null; Opened = 0 }
    }
}

Describe 'Test-YardstickTransientIntuneError' {
    It 'treats a wrapped <status> as transient: <expected>' -ForEach @(
        @{ status = 504; expected = $true }
        @{ status = 503; expected = $true }
        @{ status = 429; expected = $true }
        @{ status = 400; expected = $false }
        @{ status = 404; expected = $false }
    ) {
        Test-YardstickTransientIntuneError -ErrorRecord (New-HttpError $status) | Should -Be $expected
    }

    It 'treats a timeout as transient' {
        $record = [System.Management.Automation.ErrorRecord]::new([System.TimeoutException]::new('slow'), 'x', 'NotSpecified', $null)
        Test-YardstickTransientIntuneError -ErrorRecord $record | Should -BeTrue
    }
}

Describe 'Intune circuit breaker' {
    BeforeEach {
        $IntuneBreaker = New-Breaker
        $Script:DisplayName = 'App'
        Mock Write-Log {}
    }

    It 'opens after two transient failures in a row and logs it once' {
        Register-YardstickIntuneLookupResult -ErrorRecord (New-HttpError 504)
        $IntuneBreaker.OpenUntil | Should -BeNullOrEmpty
        Register-YardstickIntuneLookupResult -ErrorRecord (New-HttpError 504)
        $IntuneBreaker.OpenUntil | Should -BeGreaterThan (Get-Date)
        Register-YardstickIntuneLookupResult -ErrorRecord (New-HttpError 504)
        $IntuneBreaker.Opened | Should -Be 1
        Should -Invoke Write-Log -Times 1 -Exactly -ParameterFilter { $Level -eq 'Warning' }
    }

    It 'does not count a failure that is not transient' {
        Register-YardstickIntuneLookupResult -ErrorRecord (New-HttpError 400)
        Register-YardstickIntuneLookupResult -ErrorRecord (New-HttpError 400)
        $IntuneBreaker.OpenUntil | Should -BeNullOrEmpty
    }

    It 'resets the count after a successful lookup' {
        Register-YardstickIntuneLookupResult -ErrorRecord (New-HttpError 504)
        Register-YardstickIntuneLookupResult -Succeeded
        Register-YardstickIntuneLookupResult -ErrorRecord (New-HttpError 504)
        $IntuneBreaker.OpenUntil | Should -BeNullOrEmpty
    }

    It 'lets recipes through while closed, without calling Intune' {
        Mock Invoke-YardstickGraphRequest {}
        Test-YardstickIntuneBreaker | Should -BeTrue
        Should -Invoke Invoke-YardstickGraphRequest -Times 0 -Exactly
    }

    It 'fails recipes fast while open, without calling Intune' {
        Mock Invoke-YardstickGraphRequest {}
        $IntuneBreaker.OpenUntil = (Get-Date).AddMinutes(5)
        Test-YardstickIntuneBreaker | Should -BeFalse
        Should -Invoke Invoke-YardstickGraphRequest -Times 0 -Exactly
    }

    It 'closes after the cool-down when one single-attempt lookup answers' {
        Mock Invoke-YardstickGraphRequest { @() }
        $IntuneBreaker.Consecutive = 2
        $IntuneBreaker.OpenUntil = (Get-Date).AddMinutes(-1)
        Test-YardstickIntuneBreaker | Should -BeTrue
        $IntuneBreaker.OpenUntil | Should -BeNullOrEmpty
        $IntuneBreaker.Consecutive | Should -Be 0
        Should -Invoke Invoke-YardstickGraphRequest -Times 1 -Exactly -ParameterFilter {
            $MaxAttempts -eq 1 -and $Resource -like "*displayName eq 'App'"
        }
    }

    It 'stays open for another cool-down when the test lookup fails' {
        Mock Invoke-YardstickGraphRequest { throw '504' }
        $IntuneBreaker.OpenUntil = (Get-Date).AddMinutes(-1)
        Test-YardstickIntuneBreaker | Should -BeFalse
        $IntuneBreaker.OpenUntil | Should -BeGreaterThan (Get-Date).AddMinutes(9)
    }
}

Describe 'Breaker placement in the run' {
    BeforeAll {
        $script:text = Get-Content $runnerPath -Raw
    }

    It 'checks the breaker after the minRunInterval due check and before the run is recorded and the pre-download script' {
        $due = $script:text.IndexOf('Test-YardstickRecipeRunDue -Path')
        $breaker = $script:text.IndexOf('if (-not (Test-YardstickIntuneBreaker))')
        $record = $script:text.IndexOf('Set-YardstickRecipeLastRun -Path')
        $preDownload = $script:text.IndexOf('# Run the pre-download script')
        $due | Should -BeGreaterThan 0
        $breaker | Should -BeGreaterThan $due
        $record | Should -BeGreaterThan $breaker
        $preDownload | Should -BeGreaterThan $record
    }

    It 'feeds every version lookup result to the breaker' {
        $script:text | Should -Match 'Register-YardstickIntuneLookupResult -ErrorRecord \$_\s+Add-FailedApplication[^\n]+-FailureStage "Intune Lookup"'
        $script:text | Should -Match 'Register-YardstickIntuneLookupResult -Succeeded'
    }
}
