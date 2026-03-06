BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $modulePath = "$PSScriptRoot\..\Modules\YardstickSupport.psm1"
    Import-Module $modulePath -Force
}

Describe "Invoke-WithRetry" {
    It "Returns result on first success" {
        $result = Invoke-WithRetry -Label "test" -ScriptBlock { "hello" }
        $result | Should -Be "hello"
    }

    It "Retries on exception and succeeds" {
        $script:callCount = 0
        $result = Invoke-WithRetry -Label "test" -MaxRetries 3 -DelaySeconds 0 -ScriptBlock {
            $script:callCount++
            if ($script:callCount -lt 3) { throw "fail" }
            return "success"
        }
        $result | Should -Be "success"
        $script:callCount | Should -Be 3
    }

    It "Returns null after all retries exhausted" {
        $result = Invoke-WithRetry -Label "test" -MaxRetries 2 -DelaySeconds 0 -ScriptBlock {
            throw "always fails"
        }
        $result | Should -BeNullOrEmpty
    }

    It "Calls OnFailure when all retries exhausted" {
        $script:failureCalled = $false
        Invoke-WithRetry -Label "test" -MaxRetries 1 -DelaySeconds 0 -ScriptBlock {
            throw "fail"
        } -OnFailure {
            $script:failureCalled = $true
        }
        $script:failureCalled | Should -BeTrue
    }

    It "Uses VerifyBlock to determine success" {
        $script:verifyCount = 0
        $result = Invoke-WithRetry -Label "test" -MaxRetries 3 -DelaySeconds 0 -ScriptBlock {
            "output"
        } -VerifyBlock {
            $script:verifyCount++
            return ($script:verifyCount -ge 2)
        }
        $result | Should -Be "output"
        $script:verifyCount | Should -Be 2
    }

    It "Does not call OnFailure on success" {
        $script:failureCalled = $false
        Invoke-WithRetry -Label "test" -MaxRetries 3 -DelaySeconds 0 -ScriptBlock {
            "ok"
        } -OnFailure {
            $script:failureCalled = $true
        }
        $script:failureCalled | Should -BeFalse
    }

    It "Defaults to MaxRetries 3" {
        $script:callCount = 0
        Invoke-WithRetry -Label "test" -DelaySeconds 0 -ScriptBlock {
            $script:callCount++
            throw "fail"
        }
        $script:callCount | Should -Be 3
    }

    It "Respects TimeoutSeconds" {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $result = Invoke-WithRetry -Label "test" -MaxRetries 100 -DelaySeconds 1 -TimeoutSeconds 3 -ScriptBlock {
            throw "fail"
        }
        $sw.Stop()
        $result | Should -BeNullOrEmpty
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 10
    }
}
