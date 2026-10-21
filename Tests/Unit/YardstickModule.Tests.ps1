BeforeAll {
    # GetFullPath, not Resolve-Path: on a UNC share Resolve-Path keeps the "..\..",
    # and a module imported by a differently spelled path loads as a second copy.
    $script:repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
    $script:manifestPath = Join-Path $script:repoRoot 'Yardstick.psd1'
    $script:savedLogLocation = $Global:LogLocation
    $script:savedLogFile = $Global:LogFile
    $Global:LogLocation = $TestDrive
    $Global:LogFile = 'test.log'
    Import-Module powershell-yaml -ErrorAction Stop

    # Start from nothing, so the instance counts below measure this module alone
    # rather than whatever earlier test files happened to leave loaded.
    Get-Module -Name 'Yardstick*' -All | Remove-Module -Force -ErrorAction SilentlyContinue
    Import-Module $script:manifestPath -ErrorAction Stop

    function Get-ParameterSetShape {
        # Parameter set name -> sorted parameter names, without common parameters.
        param([System.Management.Automation.CommandInfo]$Command, [string[]]$Exclude = @())
        $common = [System.Management.Automation.PSCmdlet]::CommonParameters +
            [System.Management.Automation.PSCmdlet]::OptionalCommonParameters
        $shape = [ordered]@{}
        foreach ($set in $Command.ParameterSets | Sort-Object Name) {
            $names = $set.Parameters.Name | Where-Object { $_ -notin $common -and $_ -notin $Exclude } | Sort-Object
            $shape[$set.Name] = $names -join ','
        }
        $shape
    }
}

AfterAll {
    # OnRemove takes the component modules with it, so the next test file's own
    # Import-Module -Force is the only copy and Mock -ModuleName stays unambiguous.
    Get-Module -Name 'Yardstick' -All | Remove-Module -Force -ErrorAction SilentlyContinue
    Get-Module -Name 'Yardstick*' -All | Remove-Module -Force -ErrorAction SilentlyContinue
    $Global:LogLocation = $script:savedLogLocation
    $Global:LogFile = $script:savedLogFile
}

Describe "Yardstick.psd1" {
    It "is a valid manifest pointing at the root module" {
        $manifest = Test-ModuleManifest -Path $script:manifestPath -ErrorAction Stop
        $manifest.RootModule | Should -Be 'Yardstick.psm1'
        $manifest.PowerShellVersion | Should -Be ([version]'7.0')
    }

    It "exports exactly the commands it lists" {
        $listed = (Import-PowerShellDataFile -Path $script:manifestPath).FunctionsToExport | Sort-Object
        $exported = (Get-Module -Name Yardstick).ExportedCommands.Keys | Sort-Object
        $exported | Should -Be $listed
        $listed | Should -Be @('Install-Yardstick', 'Invoke-Yardstick', 'Test-YardstickSetup')
    }
}

Describe "Loading the module" {
    It "loads each component module once" {
        @(Get-Module -Name 'Yardstick*' -All | Group-Object Name | Where-Object Count -gt 1).Name |
            Should -BeNullOrEmpty
    }

    It "still holds one copy after the imports the scripts and components do" {
        # Entry scripts import the manifest again, and YardstickSupport, YardstickSetup
        # and YardstickIntune import their dependencies. None of that may load a
        # second copy: a second YardstickGraph has an empty token cache, and a
        # duplicate makes Mock -ModuleName refuse to choose.
        Import-Module $script:manifestPath
        Import-Module (Join-Path $script:repoRoot 'Modules\YardstickGraph.psm1')
        Import-Module (Join-Path $script:repoRoot 'Modules\YardstickSupport.psm1') -Global

        @(Get-Module -Name 'Yardstick*' -All | Group-Object Name | Where-Object Count -gt 1).Name |
            Should -BeNullOrEmpty
    }

    It "reloads every component on -Force, as Yardstick.ps1 does" {
        $before = Get-Module -Name YardstickGraph
        Import-Module $script:manifestPath -Force
        $after = @(Get-Module -Name YardstickGraph -All)

        $after.Count | Should -Be 1
        [object]::ReferenceEquals($before, $after[0]) | Should -BeFalse -Because 'edits to a component must be picked up by a fresh run'
    }

    It "makes the component functions available to the session" {
        # Recipes, Custom modules and the backup and watchdog threads call these
        # directly, so they cannot be private yet.
        foreach ($name in 'Write-Log', 'Invoke-YardstickGraphRequest', 'Import-YardstickPreferences', 'Add-NativeBitnessGuard') {
            Get-Command -Name $name -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty -Because "$name is called outside the module"
        }
    }

    Context "when Test-YardstickSetup.ps1 imports it" {
        BeforeAll {
            $script:setupRoot = Join-Path $TestDrive 'SetupProject'
            New-Item -ItemType Directory -Path $script:setupRoot -Force | Out-Null
            Copy-Item -LiteralPath (Join-Path $script:repoRoot 'Yardstick.Project.psd1') -Destination $script:setupRoot
        }

        It "reloads a copy the session imported earlier, so updated code runs" {
            $before = Get-Module -Name YardstickSetup
            & (Join-Path $script:repoRoot 'Test-YardstickSetup.ps1') -ProjectRoot $script:setupRoot -Offline 6>$null | Out-Null

            [object]::ReferenceEquals($before, (Get-Module -Name YardstickSetup)) | Should -BeFalse
            @(Get-Module -Name 'Yardstick*' -All | Group-Object Name | Where-Object Count -gt 1).Name | Should -BeNullOrEmpty
        }

        It "keeps the module loaded when the module's own command runs the script" {
            $before = Get-Module -Name Yardstick
            Test-YardstickSetup -ProjectRoot $script:setupRoot -Offline -ErrorAction SilentlyContinue 6>$null | Out-Null

            [object]::ReferenceEquals($before, (Get-Module -Name Yardstick)) | Should -BeTrue
        }
    }
}

Describe "Command parameters" {
    It "gives Invoke-Yardstick the same parameter sets as Yardstick.ps1, plus -ProjectRoot" {
        $script = Get-Command (Join-Path $script:repoRoot 'Yardstick.ps1')
        $function = Get-Command Invoke-Yardstick

        $function.Parameters.Keys | Should -Contain 'ProjectRoot'
        $scriptShape = Get-ParameterSetShape $script
        $functionShape = Get-ParameterSetShape $function -Exclude 'ProjectRoot'
        $functionShape.Keys | Should -Be $scriptShape.Keys
        foreach ($set in $scriptShape.Keys) {
            $functionShape[$set] | Should -Be $scriptShape[$set] -Because "parameter set $set"
        }
        $function.Parameters['ApplicationId'].Aliases | Should -Be $script.Parameters['ApplicationId'].Aliases
    }

    It "gives <Command> the parameters of <Script>" -ForEach @(
        @{ Command = 'Install-Yardstick'; Script = 'Install-Yardstick.ps1' }
        @{ Command = 'Test-YardstickSetup'; Script = 'Test-YardstickSetup.ps1' }
    ) {
        $scriptShape = Get-ParameterSetShape (Get-Command (Join-Path $script:repoRoot $Script))
        $functionShape = Get-ParameterSetShape (Get-Command $Command)
        foreach ($set in $scriptShape.Keys) {
            $functionShape[$set] | Should -Be $scriptShape[$set] -Because "parameter set $set"
        }
    }
}

Describe "The recipe scope contract" {
    # Recipes assign the runner's variables unqualified ($version, $url, ...) from
    # scriptblocks run with Invoke-Command -NoNewScope, and the runner uses exit.
    # Both only work because the runner is a script invoked with &, even when a
    # module function is the caller. If someone turns the runner into a function,
    # this is the test that explains why that breaks every recipe.
    BeforeAll {
        $script:runnerFixture = Join-Path $TestDrive 'runner-fixture.ps1'
        Set-Content -LiteralPath $script:runnerFixture -Value @'
param([int]$ExitCode = 0)
$version = '1.0'
$url = $null
$recipe = [scriptblock]::Create('$version = "2.0"; $url = "https://example.com/app.msi"')
Invoke-Command -ScriptBlock $recipe -NoNewScope
"$version|$url"
if ($ExitCode) { exit $ExitCode }
'@
        $script:caller = New-Module -Name 'YardstickScopeFixture' -ScriptBlock {
            function Invoke-RunnerFixture { param($Path, $ExitCode) & $Path -ExitCode $ExitCode }
        }
    }

    AfterAll {
        Remove-Module -ModuleInfo $script:caller -Force -ErrorAction SilentlyContinue
    }

    It "lets a recipe scriptblock assign the runner's own variables" {
        & $script:caller { param($p) Invoke-RunnerFixture -Path $p -ExitCode 0 } $script:runnerFixture |
            Should -Be '2.0|https://example.com/app.msi'
    }

    It "ends only the runner on exit, leaving the caller running with the exit code" {
        $global:LASTEXITCODE = 0
        $output = & $script:caller { param($p) Invoke-RunnerFixture -Path $p -ExitCode 3 } $script:runnerFixture
        $code = $global:LASTEXITCODE
        $output | Should -Be '2.0|https://example.com/app.msi'
        $code | Should -Be 3
    }
}

Describe "Invoke-Yardstick" {
    BeforeAll {
        # A project folder with no preferences.yaml: the runner stops at loading
        # preferences, before anything reaches the network.
        $script:projectRoot = Join-Path $TestDrive 'Project'
        New-Item -ItemType Directory -Path $script:projectRoot -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $script:repoRoot 'Yardstick.Project.psd1') -Destination $script:projectRoot
    }

    AfterEach {
        $Global:LogLocation = $TestDrive
        $Global:LogFile = 'test.log'
        $Global:PSDefaultParameterValues.Remove('Invoke-WebRequest:TimeoutSec')
        $Global:PSDefaultParameterValues.Remove('Invoke-RestMethod:TimeoutSec')
    }

    It "turns a failed run into an error instead of closing the session" {
        { Invoke-Yardstick -ApplicationId 'nothing-here' -NoEmail -ProjectRoot $script:projectRoot } |
            Should -Throw '*exit code 1*'
        $LASTEXITCODE | Should -Be 1
        Get-Content -LiteralPath (Join-Path $script:projectRoot 'YLog.log') -Raw |
            Should -Match 'Unable to load Yardstick preferences'
    }

    It "puts the caller back in their own folder" {
        Push-Location $TestDrive
        try {
            { Invoke-Yardstick -ApplicationId 'nothing-here' -NoEmail -ProjectRoot $script:projectRoot } | Should -Throw
            (Get-Location).ProviderPath | Should -Be (Get-Item -LiteralPath $TestDrive).FullName
        } finally {
            Pop-Location
        }
    }

    It "restores the caller's web request timeouts, including ones the run set" {
        # The run sets these from preferences for every recipe's web calls. Stand in
        # for that with a run that changes them and then fails.
        $Global:PSDefaultParameterValues['Invoke-WebRequest:TimeoutSec'] = 77
        Mock -ModuleName Yardstick Import-YardstickPreferences {
            $Global:PSDefaultParameterValues['Invoke-WebRequest:TimeoutSec'] = 5
            $Global:PSDefaultParameterValues['Invoke-RestMethod:TimeoutSec'] = 5
            throw 'preferences unavailable'
        }

        { Invoke-Yardstick -ApplicationId 'nothing-here' -NoEmail -ProjectRoot $script:projectRoot } | Should -Throw

        $Global:PSDefaultParameterValues['Invoke-WebRequest:TimeoutSec'] | Should -Be 77
        $Global:PSDefaultParameterValues.Contains('Invoke-RestMethod:TimeoutSec') | Should -BeFalse
    }

    It "asks for an app, a group or -All when given none, rather than failing to bind" {
        { Invoke-Yardstick -NoEmail -ProjectRoot $script:projectRoot } | Should -Throw '*exit code 1*'
        Get-Content -LiteralPath (Join-Path $script:projectRoot 'YLog.log') -Raw |
            Should -Match 'Specify -ApplicationId, -Group or -All\.'
        (Get-Command Invoke-Yardstick).DefaultParameterSet | Should -Be 'SingleApp'
        foreach ($script in 'Yardstick.ps1', 'Private\Invoke-YardstickRun.ps1') {
            $binding = (Get-Command (Join-Path $script:repoRoot $script)).ScriptBlock.Attributes |
                Where-Object { $_ -is [System.Management.Automation.CmdletBindingAttribute] }
            $binding.DefaultParameterSetName | Should -Be 'SingleApp' -Because "$script binds no arguments the same way"
        }
    }

    It "leaves no Graph secrets behind in the session" {
        $Global:ClientSecret = 'left-over'
        $Global:AuthenticationHeader = @{ Authorization = 'Bearer left-over' }
        { Invoke-Yardstick -ApplicationId 'nothing-here' -NoEmail -ProjectRoot $script:projectRoot } | Should -Throw
        Get-Variable -Name ClientSecret, AuthenticationHeader, Token -Scope Global -ErrorAction SilentlyContinue |
            Should -BeNullOrEmpty
    }
}
