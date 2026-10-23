BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    Import-Module (Join-Path $script:repoRoot 'Modules\YardstickSupport.psm1') -Force

    function New-HistoryPath {
        Join-Path $TestDrive "history-$([guid]::NewGuid().ToString('N'))\recipeRunHistory.json"
    }
}

Describe "Set-YardstickRecipeLastRun" {
    It "Creates the file and its folder" {
        $path = New-HistoryPath
        Set-YardstickRecipeLastRun -Path $path -Id '7zip' -Time '2026-10-01 02:15'
        $history = Read-YardstickRecipeRunHistory -Path $path
        $history['7zip'] | Should -Be ([datetime]'2026-10-01 02:15')
    }

    It "Keeps the other recipes' entries" {
        $path = New-HistoryPath
        Set-YardstickRecipeLastRun -Path $path -Id 'a' -Time '2026-10-01'
        Set-YardstickRecipeLastRun -Path $path -Id 'b' -Time '2026-10-02'
        Set-YardstickRecipeLastRun -Path $path -Id 'a' -Time '2026-10-03'
        $history = Read-YardstickRecipeRunHistory -Path $path
        $history.Count | Should -Be 2
        $history['a'] | Should -Be ([datetime]'2026-10-03')
        $history['b'] | Should -Be ([datetime]'2026-10-02')
    }
}

Describe "Read-YardstickRecipeRunHistory" {
    It "Returns an empty history when the file does not exist" {
        (Read-YardstickRecipeRunHistory -Path (New-HistoryPath)).Count | Should -Be 0
    }

    It "Looks up recipe ids case-insensitively" {
        $path = New-HistoryPath
        Set-YardstickRecipeLastRun -Path $path -Id '7Zip' -Time '2026-10-01'
        (Read-YardstickRecipeRunHistory -Path $path)['7zip'] | Should -Be ([datetime]'2026-10-01')
    }

    It "Warns and returns an empty history for a corrupt file" {
        $path = New-HistoryPath
        New-Item -ItemType Directory -Path (Split-Path -Parent $path) | Out-Null
        Set-Content -LiteralPath $path -Value '{ not json'
        Mock Write-Log {} -ModuleName YardstickSupport
        (Read-YardstickRecipeRunHistory -Path $path).Count | Should -Be 0
        Should -Invoke Write-Log -ModuleName YardstickSupport -ParameterFilter { $Level -eq 'Warning' }
    }

    It "Skips an entry that is not a date" {
        $path = New-HistoryPath
        New-Item -ItemType Directory -Path (Split-Path -Parent $path) | Out-Null
        Set-Content -LiteralPath $path -Value '{ "a": "soon", "b": "2026-10-02T00:00:00" }'
        $history = Read-YardstickRecipeRunHistory -Path $path
        $history.ContainsKey('a') | Should -BeFalse
        $history['b'] | Should -Be ([datetime]'2026-10-02')
    }
}

Describe "Test-YardstickRecipeRunDue" {
    BeforeEach {
        $script:path = New-HistoryPath
    }

    It "Is due for a recipe that has never run" {
        $result = Test-YardstickRecipeRunDue -Path $script:path -Id '7zip' -Days 7
        $result.IsDue | Should -BeTrue
        $result.LastRun | Should -BeNullOrEmpty
        $result.NextRun | Should -BeNullOrEmpty
    }

    It "Is not due before the interval has passed" {
        Set-YardstickRecipeLastRun -Path $script:path -Id '7zip' -Time '2026-10-01 02:15'
        $result = Test-YardstickRecipeRunDue -Path $script:path -Id '7zip' -Days 7 -Now '2026-10-07 23:59'
        $result.IsDue | Should -BeFalse
        $result.LastRun | Should -Be ([datetime]'2026-10-01 02:15')
        $result.NextRun | Should -Be ([datetime]'2026-10-08')
    }

    It "Counts calendar days, so a run reached earlier in the day is still due" {
        Set-YardstickRecipeLastRun -Path $script:path -Id '7zip' -Time '2026-10-01 02:15'
        $result = Test-YardstickRecipeRunDue -Path $script:path -Id '7zip' -Days 7 -Now '2026-10-08 02:05'
        $result.IsDue | Should -BeTrue
        $result.NextRun | Should -BeNullOrEmpty
    }

    It "Is always due with an interval of 0" {
        Set-YardstickRecipeLastRun -Path $script:path -Id '7zip' -Time '2026-10-01 02:15'
        (Test-YardstickRecipeRunDue -Path $script:path -Id '7zip' -Days 0 -Now '2026-10-01 02:16').IsDue | Should -BeTrue
    }

    It "Ignores other recipes' runs" {
        Set-YardstickRecipeLastRun -Path $script:path -Id 'git' -Time '2026-10-01'
        (Test-YardstickRecipeRunDue -Path $script:path -Id '7zip' -Days 7 -Now '2026-10-02').IsDue | Should -BeTrue
    }
}
