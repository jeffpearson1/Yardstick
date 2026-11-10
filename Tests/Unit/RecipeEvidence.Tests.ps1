BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    $script:checkScript = Join-Path $script:repoRoot 'Tests\RecipeHarness\Test-RecipeEvidence.ps1'
    Import-Module powershell-yaml -ErrorAction Stop
    Import-Module (Join-Path $script:repoRoot 'Modules\YardstickRecipeEvidence.psm1') -Force

    $script:allPhases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue')

    function New-RecipeText {
        param([string]$Id, [string]$NewLine = "`n")
        @(
            "id: $Id"
            "displayName: Test App $Id"
            'detectionType: msi'
            "iconFile: $Id.png"
            'description: A test application'
            'publisher: Test Publisher'
            'installScript: msiexec /i installer.msi /qn'
            'uninstallScript: msiexec /x {GUID} /qn'
            ''
        ) -join $NewLine
    }

    # Evidence as the Studio writes it, for a recipe with the given contents.
    function New-EvidenceJson {
        param(
            [string]$RecipePath,
            [string]$RecipeText,
            [string]$RecipeSha256,
            [object]$Validated = $true,
            [string]$Verdict = 'Validated',
            [string]$Reason = '',
            [object[]]$Phases,
            [object]$SchemaVersion = 1,
            [string]$RecipeId
        )
        if (-not $RecipeSha256) {
            $RecipeSha256 = Get-RecipeContentHash -Bytes ([System.Text.Encoding]::UTF8.GetBytes($RecipeText))
        }
        if (-not $RecipeId) { $RecipeId = Get-RecipeEvidenceId -RecipePath $RecipePath }
        if ($null -eq $Phases) {
            $Phases = foreach ($name in $script:allPhases) { @{ name = $name; status = 'Pass'; detail = '' } }
        }
        [ordered]@{
            schemaVersion     = $SchemaVersion
            recipeId          = $RecipeId
            recipePath        = $RecipePath
            recipeSha256      = $RecipeSha256
            validated         = $Validated
            verdict           = $Verdict
            reason            = $Reason
            testedAt          = '2026-11-09T10:00:00Z'
            studioVersion     = '0.5.0'
            installExperience = 'system'
            installer         = @{ fileName = 'test.msi' }
            phases            = @($Phases)
        } | ConvertTo-Json -Depth 5
    }
}

Describe 'Get-RecipeContentHash' {
    It 'matches the Studio test vector <Name>' -TestCases @(
        @{ Name = 'LF'; Hex = ''; Text = "id: 7zip`n"; Expected = 'dac9a5b6e5a249a1a674bd890731d93ada96e9fcb4b480a3659416695f52486b' }
        @{ Name = 'CRLF'; Hex = ''; Text = "id: 7zip`r`n"; Expected = 'dac9a5b6e5a249a1a674bd890731d93ada96e9fcb4b480a3659416695f52486b' }
        @{ Name = 'BOM and CRLF'; Hex = 'EFBBBF'; Text = "id: 7zip`r`ndisplayName: 7-Zip`r`n"; Expected = 'bdd1fbad1c78330b232b808ca82235b077d59638a7ac4bec6cff4a734df27d06' }
        @{ Name = 'two lines LF'; Hex = ''; Text = "id: 7zip`ndisplayName: 7-Zip`n"; Expected = 'bdd1fbad1c78330b232b808ca82235b077d59638a7ac4bec6cff4a734df27d06' }
    ) {
        $bytes = [byte[]]([Convert]::FromHexString($Hex) + [System.Text.Encoding]::UTF8.GetBytes($Text))
        Get-RecipeContentHash -Bytes $bytes | Should -BeExactly $Expected
    }

    It 'keeps a lone CR' {
        $lf = Get-RecipeContentHash -Bytes ([System.Text.Encoding]::UTF8.GetBytes("id: 7zip`n"))
        Get-RecipeContentHash -Bytes ([System.Text.Encoding]::UTF8.GetBytes("id: 7zip`r")) | Should -Not -Be $lf
    }

    It 'hashes an empty file' {
        Get-RecipeContentHash -Bytes ([byte[]]@()) | Should -BeExactly 'e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855'
    }
}

Describe 'Test-RecipeEvidenceNeeded' {
    It '<Path> needs evidence: <Expected>' -TestCases @(
        @{ Path = 'Recipes/7zip.yaml'; Expected = $true }
        @{ Path = 'Recipes/Jetbrains/clion.yaml'; Expected = $true }
        @{ Path = 'Recipes/legacy.yml'; Expected = $true }
        @{ Path = 'Recipes/Development/new.yaml'; Expected = $false }
        @{ Path = 'Recipes/Disabled/old.yaml'; Expected = $false }
        @{ Path = 'Recipes/Jetbrains/Disabled/old.yaml'; Expected = $false }
        @{ Path = 'Recipes/zoom.common.yaml'; Expected = $false }
        @{ Path = 'Recipes/notes.txt'; Expected = $false }
        @{ Path = 'RecipeGroups.yaml'; Expected = $false }
        @{ Path = 'Config/Preferences.example.yaml'; Expected = $false }
    ) {
        Test-RecipeEvidenceNeeded -Path $Path | Should -Be $Expected
    }

    It 'maps a subfolder recipe to its evidence file' {
        Get-RecipeEvidencePath -RecipePath 'Recipes/Jetbrains/clion.yaml' | Should -BeExactly 'Tests/RecipeEvidence/clion.json'
    }
}

Describe 'Test-RecipeEvidenceContent' {
    BeforeAll {
        $script:path = 'Recipes/7zip.yaml'
        $script:text = New-RecipeText -Id '7zip'
        $script:hash = Get-RecipeContentHash -Bytes ([System.Text.Encoding]::UTF8.GetBytes($script:text))

        function Test-Evidence {
            param([string]$Json)
            Test-RecipeEvidenceContent -Json $Json -RecipePath $script:path -RecipeSha256 $script:hash
        }
    }

    It 'passes valid evidence and reports its verdict and test time' {
        $result = Test-Evidence (New-EvidenceJson -RecipePath $script:path -RecipeText $script:text)
        $result.Failures | Should -BeNullOrEmpty
        $result.Verdict | Should -Be 'Validated'
        $result.TestedAt | Should -Be '2026-11-09T10:00:00Z'
    }

    It 'passes ValidatedWithWarnings with Warn phases' {
        $phases = foreach ($name in $script:allPhases) { @{ name = $name; status = 'Warn' } }
        $json = New-EvidenceJson -RecipePath $script:path -RecipeText $script:text -Verdict 'ValidatedWithWarnings' -Phases $phases
        (Test-Evidence $json).Failures | Should -BeNullOrEmpty
    }

    It 'accepts the older detection phase names' {
        $phases = foreach ($name in $script:allPhases) {
            $old = @{ PostInstallDetect = 'PostDetect'; PostUninstallDetect = 'PostUninstall' }[$name]
            @{ name = if ($old) { $old } else { $name }; status = 'Pass' }
        }
        (Test-Evidence (New-EvidenceJson -RecipePath $script:path -RecipeText $script:text -Phases $phases)).Failures | Should -BeNullOrEmpty
    }

    It 'compares phase names and statuses case-insensitively and ignores extra phases' {
        $phases = @(foreach ($name in $script:allPhases) { @{ Name = $name.ToLowerInvariant(); Status = 'PASS' } }) +
            @(@{ name = 'Reinstall'; status = 'Fail' })
        (Test-Evidence (New-EvidenceJson -RecipePath $script:path -RecipeText $script:text -Phases $phases)).Failures | Should -BeNullOrEmpty
    }

    It 'fails <Name>' -TestCases @(
        @{ Name = 'text that is not JSON'; Change = @{}; Json = '{ not json'; Pattern = 'not valid JSON' }
        @{ Name = 'a JSON array'; Change = @{}; Json = '[1]'; Pattern = 'not a JSON object' }
        @{ Name = 'schemaVersion 2'; Change = @{ SchemaVersion = 2 }; Pattern = 'schemaVersion' }
        @{ Name = 'schemaVersion as text'; Change = @{ SchemaVersion = '1' }; Pattern = 'schemaVersion' }
        @{ Name = 'a different recipeId'; Change = @{ RecipeId = '7Zip' }; Pattern = 'recipeId' }
        @{ Name = 'a different recipePath'; Change = @{ RecipePath = 'Recipes/Other/7zip.yaml' }; Pattern = 'recipePath' }
        @{ Name = 'a different hash'; Change = @{ RecipeSha256 = ('0' * 64) }; Pattern = 'Hash mismatch.*re-run the live test' }
        @{ Name = 'a malformed hash'; Change = @{ RecipeSha256 = 'abc' }; Pattern = 'not a SHA-256' }
        @{ Name = 'validated false'; Change = @{ Validated = $false; Verdict = 'NotValidated'; Reason = 'the recipe has changed since' }; Pattern = 'validated is not true.*the recipe has changed since' }
        @{ Name = 'validated as text'; Change = @{ Validated = 'true' }; Pattern = 'validated is not true' }
        @{ Name = 'verdict NotValidated'; Change = @{ Verdict = 'NotValidated' }; Pattern = "verdict is 'NotValidated'" }
        @{ Name = 'a missing phase'; Change = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect' | ForEach-Object { @{ name = $_; status = 'Pass' } }) }; Pattern = 'missing.*Residue' }
        @{ Name = 'no phases'; Change = @{ Phases = @() }; Pattern = 'missing.*Schema' }
        @{ Name = 'a failed phase'; Change = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue' | ForEach-Object { @{ name = $_; status = if ($_ -eq 'Install') { 'Fail' } else { 'Pass' } } }) }; Pattern = "Install status is 'Fail'" }
        @{ Name = 'a skipped phase'; Change = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue' | ForEach-Object { @{ name = $_; status = if ($_ -eq 'Uninstall') { 'Skip' } else { 'Pass' } } }) }; Pattern = "Uninstall status is 'Skip'" }
        @{ Name = 'a phase with no status'; Change = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue' | ForEach-Object { if ($_ -eq 'Install') { @{ name = $_ } } else { @{ name = $_; status = 'Pass' } } }) }; Pattern = 'Install status is missing' }
        @{ Name = 'a phase with a null status'; Change = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue' | ForEach-Object { @{ name = $_; status = if ($_ -eq 'Install') { $null } else { 'Pass' } } }) }; Pattern = 'Install status is missing' }
        @{ Name = 'a phase with an empty status'; Change = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue' | ForEach-Object { @{ name = $_; status = if ($_ -eq 'Install') { '' } else { 'Pass' } } }) }; Pattern = 'Install status is missing' }
        @{ Name = 'phases with no statuses at all'; Change = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue' | ForEach-Object { @{ name = $_ } }) }; Pattern = 'Schema status is missing' }
    ) {
        if (-not $Json) {
            $params = @{ RecipePath = $script:path; RecipeText = $script:text; RecipeId = '7zip' }
            foreach ($key in $Change.Keys) { $params[$key] = $Change[$key] }
            $Json = New-EvidenceJson @params
        }
        $failures = @((Test-Evidence $Json).Failures)
        $failures | Should -Not -BeNullOrEmpty
        ($failures -join "`n") | Should -Match $Pattern
    }
}

Describe 'Test-RecipeEvidence.ps1' {
    BeforeAll {
        function Invoke-TestGit {
            param([string[]]$GitArgs, [byte[]]$InputBytes)
            $output = & git -C $script:repo -c user.name=Test -c user.email=test@example.com -c commit.gpgsign=false @GitArgs 2>&1
            if ($LASTEXITCODE -ne 0) { throw "git $($GitArgs -join ' ') failed: $output" }
        }

        function Set-RepoFile {
            param([string]$Path, [string]$Text)
            $full = Join-Path $script:repo $Path
            New-Item -ItemType Directory -Path (Split-Path $full) -Force | Out-Null
            [System.IO.File]::WriteAllText($full, $Text, [System.Text.UTF8Encoding]::new($false))
        }

        # A recipe and its evidence, as the Studio publishes them.
        function Add-Recipe {
            param([string]$Path, [string]$NewLine = "`n", [hashtable]$Evidence = @{}, [switch]$NoEvidence)
            $id = Get-RecipeEvidenceId -RecipePath $Path
            $text = New-RecipeText -Id $id -NewLine $NewLine
            Set-RepoFile -Path $Path -Text $text
            if (-not $NoEvidence) {
                Set-RepoFile -Path (Get-RecipeEvidencePath -RecipePath $Path) -Text (New-EvidenceJson -RecipePath $Path -RecipeText $text @Evidence)
            }
            $text
        }

        function Save-Commit {
            param([string]$Message = 'change', [string]$AutoCrlf = 'false')
            Invoke-TestGit -GitArgs @('-c', "core.autocrlf=$AutoCrlf", 'add', '-A')
            Invoke-TestGit -GitArgs @('commit', '-q', '--allow-empty', '-m', $Message)
        }

        function Invoke-Check {
            param([switch]$Waived)
            $summary = Join-Path $TestDrive "summary-$([guid]::NewGuid().ToString('N')).md"
            $output = & $script:checkScript -RepositoryPath $script:repo -BaseRef main -HeadRef HEAD -SummaryPath $summary -Waived:$Waived -PassThru 6>$null 3>$null
            [PSCustomObject]@{
                ExitCode = $LASTEXITCODE
                Result   = @($output)[-1]
                Summary  = if (Test-Path -LiteralPath $summary) { Get-Content -LiteralPath $summary -Raw } else { '' }
            }
        }
    }

    AfterAll {
        # Test-RecipeSchema.ps1 imports Yardstick.psd1 into the global session state.
        # Unload it so Mock -ModuleName in the test files that run next sees one copy.
        Get-Module -Name 'Yardstick', 'YardstickSupport', 'YardstickCredential', 'YardstickConfig', 'YardstickSetup', 'YardstickIntune', 'YardstickGraph', 'YardstickGitHub', 'YardstickRecipeEvidence', 'YardstickRecipeSync', 'YardstickLog' -All |
            Remove-Module -Force -ErrorAction SilentlyContinue
    }

    BeforeEach {
        # main holds one tested recipe; each test changes things on a branch.
        $script:repo = Join-Path $TestDrive "repo-$([guid]::NewGuid().ToString('N'))"
        New-Item -ItemType Directory -Path $script:repo | Out-Null
        Invoke-TestGit -GitArgs @('init', '-q', '-b', 'main')
        Add-Recipe -Path 'Recipes/existing.yaml' | Out-Null
        Save-Commit -Message 'base'
        Invoke-TestGit -GitArgs @('checkout', '-q', '-b', 'feature')
    }

    It 'passes when no recipe changed' {
        Set-RepoFile -Path 'README.md' -Text 'docs only'
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 0
        $check.Result.Passed | Should -BeTrue
        $check.Result.Recipes | Should -BeNullOrEmpty
        $check.Summary | Should -Match 'No recipe changes'
    }

    It 'passes a new recipe with valid evidence and writes its summary row' {
        Add-Recipe -Path 'Recipes/foo.yaml' | Out-Null
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 0
        $check.Result.Recipes.Id | Should -Be 'foo'
        $check.Result.Recipes.Result | Should -Be 'Pass'
        $check.Summary | Should -Match '\| `foo` \| Pass \| Validated \| 2026-11-09T10:00:00Z \|'
    }

    It 'fails with a hash mismatch when the recipe changed after testing' {
        $text = Add-Recipe -Path 'Recipes/foo.yaml'
        Set-RepoFile -Path 'Recipes/foo.yaml' -Text $text.Replace('Test App foo', 'Test App fo0')
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 1
        $check.Result.Recipes.Result | Should -Be 'Fail'
        ($check.Result.Recipes.Reasons -join "`n") | Should -Match 'Hash mismatch.*re-run the live test'
    }

    It 'ignores line endings when committed with core.autocrlf=<AutoCrlf>' -TestCases @(
        @{ AutoCrlf = 'true' }
        @{ AutoCrlf = 'false' }
    ) {
        # The Studio hashed the LF text; the file on disk has CRLF.
        Add-Recipe -Path 'Recipes/foo.yaml' -NewLine "`r`n" -Evidence @{ RecipeSha256 = (Get-RecipeContentHash -Bytes ([System.Text.Encoding]::UTF8.GetBytes((New-RecipeText -Id 'foo')))) } | Out-Null
        Save-Commit -AutoCrlf $AutoCrlf
        $check = Invoke-Check
        $check.ExitCode | Should -Be 0
        $check.Result.Recipes.Result | Should -Be 'Pass'
    }

    It 'fails a recipe without an evidence file' {
        Add-Recipe -Path 'Recipes/foo.yaml' -NoEvidence | Out-Null
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 1
        ($check.Result.Recipes.Reasons -join "`n") | Should -Match 'No evidence file: Tests/RecipeEvidence/foo.json'
    }

    It 'fails <Name>, and passes with warnings when waived' -TestCases @(
        @{ Name = 'unvalidated evidence'; Evidence = @{ Validated = $false; Verdict = 'NotValidated'; Reason = 'published without a test' }; Pattern = 'validated is not true' }
        @{ Name = 'a missing phase'; Evidence = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect' | ForEach-Object { @{ name = $_; status = 'Pass' } }) }; Pattern = 'missing.*Residue' }
        @{ Name = 'a failed phase'; Evidence = @{ Phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue' | ForEach-Object { @{ name = $_; status = if ($_ -eq 'Install') { 'Fail' } else { 'Pass' } } }) }; Pattern = "Install status is 'Fail'" }
    ) {
        Add-Recipe -Path 'Recipes/foo.yaml' -Evidence $Evidence | Out-Null
        Save-Commit

        $check = Invoke-Check
        $check.ExitCode | Should -Be 1
        $check.Result.Recipes.Result | Should -Be 'Fail'
        ($check.Result.Recipes.Reasons -join "`n") | Should -Match $Pattern
        $check.Summary | Should -Match $Pattern

        $waived = Invoke-Check -Waived
        $waived.ExitCode | Should -Be 0
        $waived.Result.Passed | Should -BeTrue
        $waived.Result.Recipes.Result | Should -Be 'Waived'
        $waived.Summary | Should -Match 'evidence-waived'
    }

    It 'fails a recipe that fails the schema check, even when waived' {
        $text = (New-RecipeText -Id 'foo').Replace('id: foo', 'id: Foo')
        Set-RepoFile -Path 'Recipes/foo.yaml' -Text $text
        Set-RepoFile -Path 'Tests/RecipeEvidence/foo.json' -Text (New-EvidenceJson -RecipePath 'Recipes/foo.yaml' -RecipeText $text)
        Save-Commit
        $check = Invoke-Check -Waived
        $check.ExitCode | Should -Be 1
        $check.Result.Recipes.Result | Should -Be 'Fail'
        $check.Result.Recipes.EvidenceFailures | Should -BeNullOrEmpty
        ($check.Result.Recipes.Reasons -join "`n") | Should -Match "Schema: .*must exactly match filename 'foo'"
    }

    It 'checks an inherited recipe against its base as committed, not as edited in the working tree' {
        Add-Recipe -Path 'Recipes/Bases/parent.yaml' | Out-Null
        $child = "base: parent`nid: child`ndisplayName: Child App`n"
        Set-RepoFile -Path 'Recipes/child.yaml' -Text $child
        Set-RepoFile -Path 'Tests/RecipeEvidence/child.json' -Text (New-EvidenceJson -RecipePath 'Recipes/child.yaml' -RecipeText $child)
        Save-Commit
        # Uncommitted, and unreadable as YAML: it must not be used.
        Set-RepoFile -Path 'Recipes/Bases/parent.yaml' -Text "id: [broken"
        $check = Invoke-Check
        $check.ExitCode | Should -Be 0
        ($check.Result.Recipes | Where-Object Id -eq 'child').Result | Should -Be 'Pass'
    }

    It 'fails an inherited recipe whose base is only in the working tree, even when waived' {
        $child = "base: ghost`nid: child`ndisplayName: Child App`n"
        Set-RepoFile -Path 'Recipes/child.yaml' -Text $child
        Set-RepoFile -Path 'Tests/RecipeEvidence/child.json' -Text (New-EvidenceJson -RecipePath 'Recipes/child.yaml' -RecipeText $child)
        Save-Commit
        Set-RepoFile -Path 'Recipes/ghost.yaml' -Text (New-RecipeText -Id 'ghost')
        $check = Invoke-Check -Waived
        $check.ExitCode | Should -Be 1
        ($check.Result.Recipes.Reasons -join "`n") | Should -Match "Schema: Base recipe 'ghost' is not in HEAD"
    }

    It 'matches a recipe in a subfolder to its evidence by id' {
        Add-Recipe -Path 'Recipes/Jetbrains/clion.yaml' | Out-Null
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 0
        $check.Result.Recipes.EvidencePath | Should -Be 'Tests/RecipeEvidence/clion.json'
        $check.Result.Recipes.Result | Should -Be 'Pass'
    }

    It 'requires a changed recipe to be re-tested, not only new ones' {
        Set-RepoFile -Path 'Recipes/existing.yaml' -Text ((New-RecipeText -Id 'existing') + "notes: changed`n")
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 1
        ($check.Result.Recipes.Reasons -join "`n") | Should -Match 'Hash mismatch'
    }

    It 'checks a recipe moved to a new folder against its new path' {
        New-Item -ItemType Directory -Path (Join-Path $script:repo 'Recipes/Moved') | Out-Null
        Invoke-TestGit -GitArgs @('mv', 'Recipes/existing.yaml', 'Recipes/Moved/existing.yaml')
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 1
        $check.Result.Recipes.RecipePath | Should -Be 'Recipes/Moved/existing.yaml'
        ($check.Result.Recipes.Reasons -join "`n") | Should -Match 'recipePath'
    }

    It 'skips development and disabled recipes' {
        Set-RepoFile -Path 'Recipes/Development/wip.yaml' -Text (New-RecipeText -Id 'wip')
        Set-RepoFile -Path 'Recipes/Disabled/old.yaml' -Text (New-RecipeText -Id 'old')
        Invoke-TestGit -GitArgs @('add', '-f', 'Recipes')
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 0
        $check.Result.Recipes | Should -BeNullOrEmpty
    }

    It 'warns about evidence left behind by a deleted recipe' {
        Invoke-TestGit -GitArgs @('rm', '-q', 'Recipes/existing.yaml')
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 0
        $check.Result.Warnings | Should -Match 'existing.json belongs to a recipe this branch deletes'
        $check.Summary | Should -Match 'Warnings'
    }

    It 'warns about evidence that matches no recipe' {
        Set-RepoFile -Path 'Tests/RecipeEvidence/orphan.json' -Text '{}'
        Save-Commit
        $check = Invoke-Check
        $check.ExitCode | Should -Be 0
        $check.Result.Warnings | Should -Match 'orphan.json matches no recipe'
    }

    It 'reads the committed recipe, not the working tree' {
        Add-Recipe -Path 'Recipes/foo.yaml' | Out-Null
        Save-Commit
        Set-RepoFile -Path 'Recipes/foo.yaml' -Text 'uncommitted edit'
        (Invoke-Check).ExitCode | Should -Be 0
    }
}
