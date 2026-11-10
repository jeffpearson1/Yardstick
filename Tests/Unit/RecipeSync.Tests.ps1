BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    Import-Module powershell-yaml -ErrorAction Stop
    foreach ($module in 'YardstickSupport', 'YardstickRecipeEvidence', 'YardstickConfig', 'YardstickRecipeSync') {
        Import-Module (Join-Path $script:repoRoot "Modules\$module.psm1") -Force
    }

    $script:utf8 = [System.Text.UTF8Encoding]::new($false)
    $script:phases = @('Schema', 'Download', 'PreDetect', 'Install', 'PostInstallDetect', 'Uninstall', 'PostUninstallDetect', 'Residue')

    # Fake GitHub repositories, by owner/name: path -> bytes. The mock serves the
    # tree listing and blobs from here and records every call.
    $Global:RecipeSyncTestRepos = @{}
    $Global:RecipeSyncTestCalls = [System.Collections.Generic.List[string]]::new()
    $Global:RecipeSyncTestFailBlob = $null

    function Set-FakeRepo {
        param([string]$Repository, [hashtable]$Files)
        $content = @{}
        foreach ($path in $Files.Keys) {
            $value = $Files[$path]
            $content[$path] = if ($value -is [byte[]]) { $value } else { $script:utf8.GetBytes([string]$value) }
        }
        $Global:RecipeSyncTestRepos[$Repository] = $content
    }

    function New-TestEvidence {
        param([string]$Path, [string]$Text, [string]$Hash)
        if (-not $Hash) { $Hash = Get-RecipeContentHash -Bytes $script:utf8.GetBytes($Text) }
        @{
            schemaVersion = 1
            recipeId      = Get-RecipeEvidenceId -RecipePath $Path
            recipePath    = $Path
            recipeSha256  = $Hash
            validated     = $true
            verdict       = 'Validated'
            testedAt      = '2026-11-09T10:00:00Z'
            phases        = @($script:phases | ForEach-Object { @{ name = $_; status = 'Pass' } })
        } | ConvertTo-Json -Depth 4
    }

    function Get-BlobCallCount {
        @($Global:RecipeSyncTestCalls | Where-Object { $_ -match '/git/blobs/' }).Count
    }

    function Invoke-TestSync {
        param([hashtable]$Change = @{})
        $prefs = @{
            Recipes                      = $script:recipes
            syncRecipes                  = $true
            recipeRepository             = 'contoso/yardstick@main'
            additionalRecipeRepositories = @()
            requireRecipeEvidence        = $false
        }
        foreach ($key in $Change.Keys) { $prefs[$key] = $Change[$key] }
        $Global:RecipeSyncTestCalls.Clear()
        @(Invoke-YardstickRecipeSync -Preferences $prefs -ProjectRoot $script:root)
    }

    function Get-RecipeText {
        param([string]$Name)
        [System.IO.File]::ReadAllText((Join-Path $script:recipes $Name))
    }
}

AfterAll {
    Remove-Variable -Name RecipeSyncTestRepos, RecipeSyncTestCalls, RecipeSyncTestFailBlob -Scope Global -ErrorAction SilentlyContinue
}

Describe 'ConvertFrom-YardstickRecipeRepositorySpec' {
    It 'parses <Spec>' -TestCases @(
        @{ Spec = 'contoso/recipes'; Owner = 'contoso'; Name = 'recipes'; Ref = ''; Key = 'contoso-recipes' }
        @{ Spec = ' contoso/recipes@main '; Owner = 'contoso'; Name = 'recipes'; Ref = 'main'; Key = 'contoso-recipes@main' }
        @{ Spec = 'jeffpearson1/Yardstick@release/1.2'; Owner = 'jeffpearson1'; Name = 'Yardstick'; Ref = 'release/1.2'; Key = 'jeffpearson1-Yardstick@release_1.2' }
    ) {
        $parsed = ConvertFrom-YardstickRecipeRepositorySpec -Spec $Spec
        $parsed.Owner | Should -BeExactly $Owner
        $parsed.Name | Should -BeExactly $Name
        $parsed.Ref | Should -BeExactly $Ref
        $parsed.Key | Should -BeExactly $Key
    }

    It 'refuses <Spec>' -TestCases @(
        @{ Spec = 'https://github.com/contoso/recipes' }
        @{ Spec = 'contoso' }
        @{ Spec = 'contoso/recipes/extra' }
        @{ Spec = '../recipes' }
    ) {
        { ConvertFrom-YardstickRecipeRepositorySpec -Spec $Spec } | Should -Throw '*owner/name*'
    }
}

Describe 'Get-YardstickGitBlobSha' {
    It 'matches git hash-object' {
        Get-YardstickGitBlobSha -Bytes $script:utf8.GetBytes("hello`n") | Should -BeExactly 'ce013625030ba8dba906f756967f9e9ca394464a'
        Get-YardstickGitBlobSha -Bytes ([byte[]]@()) | Should -BeExactly 'e69de29bb2d1d6434b8b29ae775ad8c2e48c5391'
    }
}

Describe 'Invoke-YardstickRecipeSync' {
    BeforeAll {
        Mock Invoke-YardstickGitHubApi -ModuleName YardstickRecipeSync {
            param($Uri)
            $Global:RecipeSyncTestCalls.Add($Uri)
            if ($Uri -notmatch '^repos/(?<repo>[^/]+/[^/]+)(?<rest>/.*)?$') { throw "Unexpected URI $Uri" }
            $files = $Global:RecipeSyncTestRepos[$Matches.repo]
            if ($null -eq $files) { throw 'Response status code does not indicate success: 404 (Not Found).' }
            $rest = [string]$Matches.rest
            if (-not $rest) { return [pscustomobject]@{ default_branch = 'main' } }
            $entries = foreach ($path in ($files.Keys | Sort-Object)) {
                [pscustomobject]@{ path = $path; type = 'blob'; sha = (Get-YardstickGitBlobSha -Bytes $files[$path]) }
            }
            if ($rest -match '^/git/trees/') {
                $treeSha = Get-YardstickGitBlobSha -Bytes ([System.Text.Encoding]::UTF8.GetBytes((@($entries | ForEach-Object { "$($_.path)=$($_.sha)" }) -join "`n")))
                return [pscustomobject]@{ sha = $treeSha; truncated = $false; tree = @($entries) }
            }
            if ($rest -match '^/git/blobs/(?<sha>[0-9a-f]{40})$') {
                if ($Global:RecipeSyncTestFailBlob -and $Matches.sha -eq $Global:RecipeSyncTestFailBlob) {
                    throw 'GitHub asked Yardstick to wait.'
                }
                $entry = $entries | Where-Object sha -eq $Matches.sha | Select-Object -First 1
                return [pscustomobject]@{ encoding = 'base64'; content = [Convert]::ToBase64String($files[$entry.path]) }
            }
            throw "Unexpected URI $Uri"
        }
    }

    BeforeEach {
        $script:root = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $script:recipes = Join-Path $script:root 'Recipes'
        New-Item -ItemType Directory -Path $script:recipes -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $script:repoRoot 'Yardstick.Project.psd1') -Destination $script:root
        $Global:RecipeSyncTestRepos.Clear()
        $Global:RecipeSyncTestFailBlob = $null
        Set-FakeRepo -Repository 'contoso/yardstick' -Files @{
            'Recipes/7zip.yaml'                 = "id: 7zip`n"
            'Recipes/Jetbrains/clion.yaml'      = "id: clion`n"
            'Recipes/Development/wip.yaml'      = "id: wip`n"
            'Recipes/Disabled/old.yaml'         = "id: old`n"
            'Recipes/zoom.common.yaml'          = "id: zoom.common`n"
            'Recipes/README.md'                 = 'not a recipe'
            'Modules/YardstickSupport.psm1'     = '# runtime'
        }
    }

    It 'does nothing when syncRecipes is off' {
        $messages = Invoke-TestSync -Change @{ syncRecipes = $false }
        $messages | Should -BeNullOrEmpty
        $Global:RecipeSyncTestCalls | Should -BeNullOrEmpty
    }

    It 'brings in only recipes, keeping their folders' {
        $messages = Invoke-TestSync
        Get-RecipeText '7zip.yaml' | Should -Be "id: 7zip`n"
        Get-RecipeText 'Jetbrains\clion.yaml' | Should -Be "id: clion`n"
        Join-Path $script:recipes 'Development' | Should -Not -Exist
        Join-Path $script:recipes 'Disabled' | Should -Not -Exist
        Join-Path $script:recipes 'zoom.common.yaml' | Should -Not -Exist
        Join-Path $script:recipes 'README.md' | Should -Not -Exist
        Join-Path $script:root 'Modules' | Should -Not -Exist
        $messages[0].Level | Should -Be 'Info'
        $messages[0].Message | Should -Match '2 added, 0 updated, 0 removed, 0 unchanged'
    }

    It 'makes one call when the repository has not changed since the last sync' {
        Invoke-TestSync | Out-Null
        $messages = Invoke-TestSync
        $Global:RecipeSyncTestCalls.Count | Should -Be 1
        $messages.Message | Should -Match 'up to date'
    }

    It 'asks GitHub for the default branch when no branch is given' {
        Invoke-TestSync -Change @{ recipeRepository = 'contoso/yardstick' } | Out-Null
        $Global:RecipeSyncTestCalls[0] | Should -Be 'repos/contoso/yardstick'
        $Global:RecipeSyncTestCalls[1] | Should -Match '^repos/contoso/yardstick/git/trees/main\?recursive=1$'
    }

    It 'downloads only files that differ from the local copy' {
        Set-Content -LiteralPath (Join-Path $script:recipes '7zip.yaml') -Value "id: 7zip`n" -NoNewline
        $messages = Invoke-TestSync
        Get-BlobCallCount | Should -Be 1
        $messages[0].Message | Should -Match '1 added, 0 updated, 0 removed, 1 unchanged'
    }

    It 'updates a recipe changed in the repository and backs up the old copy' {
        Invoke-TestSync | Out-Null
        Set-FakeRepo -Repository 'contoso/yardstick' -Files @{ 'Recipes/7zip.yaml' = "id: 7zip`nversion: 2`n"; 'Recipes/Jetbrains/clion.yaml' = "id: clion`n" }
        $messages = Invoke-TestSync
        Get-RecipeText '7zip.yaml' | Should -Be "id: 7zip`nversion: 2`n"
        $messages[0].Message | Should -Match '0 added, 1 updated'
        $backup = Get-ChildItem -LiteralPath (Join-Path $script:root 'Artifacts\RecipeSyncBackups') -Recurse -File -Filter '7zip.yaml'
        [System.IO.File]::ReadAllText($backup.FullName) | Should -Be "id: 7zip`n"
    }

    It 'keeps a recipe edited here since it was synced, with a warning' {
        Invoke-TestSync | Out-Null
        Set-Content -LiteralPath (Join-Path $script:recipes '7zip.yaml') -Value "id: 7zip`nlocal: edit`n" -NoNewline
        Set-FakeRepo -Repository 'contoso/yardstick' -Files @{ 'Recipes/7zip.yaml' = "id: 7zip`nversion: 2`n"; 'Recipes/Jetbrains/clion.yaml' = "id: clion`n" }
        $messages = Invoke-TestSync
        Get-RecipeText '7zip.yaml' | Should -Be "id: 7zip`nlocal: edit`n"
        (@($messages | Where-Object Level -eq 'Warning').Message -join "`n") | Should -Match 'Kept recipe 7zip.yaml, which was edited here'
    }

    It 'replaces a different copy it has never synced, keeping a backup' {
        Set-Content -LiteralPath (Join-Path $script:recipes '7zip.yaml') -Value "id: 7zip`nold: version`n" -NoNewline
        Invoke-TestSync | Out-Null
        Get-RecipeText '7zip.yaml' | Should -Be "id: 7zip`n"
        Get-ChildItem -LiteralPath (Join-Path $script:root 'Artifacts\RecipeSyncBackups') -Recurse -File -Filter '7zip.yaml' | Should -Not -BeNullOrEmpty
    }

    It 'removes a recipe removed from the repository, but never a local-only one' {
        Set-Content -LiteralPath (Join-Path $script:recipes 'mine.yaml') -Value "id: mine`n"
        Invoke-TestSync | Out-Null
        Set-FakeRepo -Repository 'contoso/yardstick' -Files @{ 'Recipes/7zip.yaml' = "id: 7zip`n" }
        $messages = Invoke-TestSync
        Join-Path $script:recipes 'Jetbrains\clion.yaml' | Should -Not -Exist
        Join-Path $script:recipes 'mine.yaml' | Should -Exist
        $messages[0].Message | Should -Match '1 removed'
        Get-ChildItem -LiteralPath (Join-Path $script:root 'Artifacts\RecipeSyncBackups') -Recurse -File -Filter 'clion.yaml' | Should -Not -BeNullOrEmpty
    }

    It 'warns and changes nothing when the repository cannot be reached' {
        Invoke-TestSync | Out-Null
        $messages = Invoke-TestSync -Change @{ recipeRepository = 'contoso/missing@main' }
        $messages.Level | Should -Be 'Warning'
        $messages.Message | Should -Match 'Could not sync recipes from contoso/missing@main: .*404.*Nothing has been synced from it yet'
        Get-RecipeText '7zip.yaml' | Should -Be "id: 7zip`n"
    }

    It 'keeps what it synced when GitHub stops it part way, and carries on next time' {
        $Global:RecipeSyncTestFailBlob = Get-YardstickGitBlobSha -Bytes $script:utf8.GetBytes("id: clion`n")
        $messages = Invoke-TestSync
        $messages.Message | Should -Match 'asked Yardstick to wait'
        Join-Path $script:recipes '7zip.yaml' | Should -Exist

        $Global:RecipeSyncTestFailBlob = $null
        $messages = Invoke-TestSync
        Get-BlobCallCount | Should -Be 1
        Get-RecipeText 'Jetbrains\clion.yaml' | Should -Be "id: clion`n"
        $messages[0].Message | Should -Match '1 added, 0 updated, 0 removed, 1 unchanged'
    }

    Context 'Linked repositories' {
        BeforeEach {
            Set-FakeRepo -Repository 'contoso/lab' -Files @{
                'Recipes/labapp.yaml' = "id: labapp`n"
                'Recipes/7zip.yaml'   = "id: 7zip`nlab: copy`n"
            }
            $script:labFolder = Join-Path $script:root 'Local\RecipeRepos\contoso-lab\Recipes'
            $script:linkedPrefs = @{ additionalRecipeRepositories = @('contoso/lab') }
        }

        It 'mirrors a linked repository into its own folder' {
            Invoke-TestSync -Change $script:linkedPrefs | Out-Null
            Join-Path $script:labFolder 'labapp.yaml' | Should -Exist
            Get-RecipeText '7zip.yaml' | Should -Be "id: 7zip`n"
        }

        It 'removes anything else from the mirror' {
            Invoke-TestSync -Change $script:linkedPrefs | Out-Null
            Set-Content -LiteralPath (Join-Path $script:labFolder 'stray.yaml') -Value 'id: stray'
            Set-FakeRepo -Repository 'contoso/lab' -Files @{ 'Recipes/labapp.yaml' = "id: labapp`nversion: 2`n" }
            Invoke-TestSync -Change $script:linkedPrefs | Out-Null
            Join-Path $script:labFolder 'stray.yaml' | Should -Not -Exist
            Join-Path $script:labFolder '7zip.yaml' | Should -Not -Exist
            [System.IO.File]::ReadAllText((Join-Path $script:labFolder 'labapp.yaml')) | Should -Be "id: labapp`nversion: 2`n"
        }

        It 'searches the mirror after Recipes, so the Recipes copy wins' {
            Invoke-TestSync -Change $script:linkedPrefs | Out-Null
            $prefs = @{ Recipes = $script:recipes; syncRecipes = $true; recipeRepository = 'contoso/yardstick@main'; additionalRecipeRepositories = @('contoso/lab') }
            $paths = @(Get-YardstickRecipeSearchPath -Preferences $prefs -ProjectRoot $script:root)
            $paths | Should -Be @($script:recipes, $script:labFolder)

            @(Get-YardstickRecipeFile -RecipesPath $paths -Id '7zip')[0].FullName | Should -Be (Join-Path $script:recipes '7zip.yaml')
            @(Get-YardstickRecipeFile -RecipesPath $paths -Id 'labapp')[0].FullName | Should -Be (Join-Path $script:labFolder 'labapp.yaml')

            $prefs.syncRecipes = $false
            @(Get-YardstickRecipeSearchPath -Preferences $prefs -ProjectRoot $script:root) | Should -Be @($script:recipes)
        }

        It 'lets a linked recipe use a base recipe from Recipes' {
            Invoke-TestSync -Change $script:linkedPrefs | Out-Null
            $merged = Merge-RecipeWithBase -Recipe @{ base = '7zip'; id = 'labapp' } -RecipesPath @($script:recipes, $script:labFolder)
            $merged.id | Should -Be 'labapp'
        }
    }

    Context 'requireRecipeEvidence' {
        BeforeEach {
            $script:text = "id: 7zip`nversion: 2`n"
            $script:evidencePrefs = @{ requireRecipeEvidence = $true }
        }

        It 'takes a recipe whose evidence proves these contents' {
            Set-FakeRepo -Repository 'contoso/yardstick' -Files @{
                'Recipes/7zip.yaml'              = $script:text
                'Tests/RecipeEvidence/7zip.json' = New-TestEvidence -Path 'Recipes/7zip.yaml' -Text $script:text
            }
            Invoke-TestSync -Change $script:evidencePrefs | Out-Null
            Get-RecipeText '7zip.yaml' | Should -Be $script:text
        }

        It 'keeps the copy here when the evidence is for other contents' {
            Set-Content -LiteralPath (Join-Path $script:recipes '7zip.yaml') -Value "id: 7zip`n" -NoNewline
            Set-FakeRepo -Repository 'contoso/yardstick' -Files @{
                'Recipes/7zip.yaml'              = $script:text
                'Tests/RecipeEvidence/7zip.json' = New-TestEvidence -Path 'Recipes/7zip.yaml' -Text "id: 7zip`nversion: 1`n"
            }
            $messages = Invoke-TestSync -Change $script:evidencePrefs
            Get-RecipeText '7zip.yaml' | Should -Be "id: 7zip`n"
            (@($messages | Where-Object Level -eq 'Warning').Message -join "`n") | Should -Match 'Did not take recipe 7zip.yaml: Hash mismatch'
        }

        It 'does not take a recipe without evidence' {
            $messages = Invoke-TestSync -Change $script:evidencePrefs
            Join-Path $script:recipes '7zip.yaml' | Should -Not -Exist
            (@($messages | Where-Object Level -eq 'Warning').Message -join "`n") | Should -Match 'no evidence file \(Tests/RecipeEvidence/7zip.json\)'
            Get-BlobCallCount | Should -Be 0
        }
    }
}

Describe 'Get-YardstickRecipeFile with several folders' {
    BeforeAll {
        $script:first = Join-Path $TestDrive 'first'
        $script:second = Join-Path $TestDrive 'second'
        New-Item -ItemType Directory -Path (Join-Path $script:first 'Disabled'), $script:second -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:first 'Disabled\shared.yaml') -Value 'id: shared'
        Set-Content -LiteralPath (Join-Path $script:second 'shared.yaml') -Value 'id: shared'
        Set-Content -LiteralPath (Join-Path $script:second 'only.yaml') -Value 'id: only'
    }

    It 'prefers an enabled recipe in a later folder over a disabled one in an earlier folder' {
        @(Get-YardstickRecipeFile -RecipesPath @($script:first, $script:second) -Id 'shared' -IncludeDisabled)[0].FullName |
            Should -Be (Join-Path $script:second 'shared.yaml')
    }

    It 'skips a later folder that does not exist' {
        @(Get-YardstickRecipeFile -RecipesPath @($script:second, (Join-Path $TestDrive 'missing'))).Count | Should -Be 2
    }

    It 'still requires the first folder to exist' {
        { Get-YardstickRecipeFile -RecipesPath (Join-Path $TestDrive 'missing') } | Should -Throw
    }
}

Describe 'Recipe repository preferences' {
    It 'defaults to Yardstick''s own repository with sync off' {
        $defaults = Get-YardstickDefaultPreferences -ProjectRoot $script:repoRoot
        $defaults.syncRecipes | Should -BeFalse
        $defaults.recipeRepository | Should -Be 'jeffpearson1/Yardstick@main'
        @($defaults.additionalRecipeRepositories).Count | Should -Be 0
        $defaults.requireRecipeEvidence | Should -BeFalse
    }

    It 'reports <Name>' -TestCases @(
        @{ Name = 'a repository that is not owner/name'; Change = @{ additionalRecipeRepositories = @('https://github.com/contoso/lab') }; Errors = "is not a GitHub repository"; Warnings = $null }
        @{ Name = 'sync with no repository'; Change = @{ syncRecipes = $true; recipeRepository = '' }; Errors = $null; Warnings = 'nothing to sync' }
        @{ Name = 'evidence without sync'; Change = @{ requireRecipeEvidence = $true }; Errors = $null; Warnings = 'syncRecipes is off' }
    ) {
        $prefs = Get-YardstickDefaultPreferences -ProjectRoot $script:repoRoot
        $prefs.Recipes = $TestDrive
        foreach ($key in $Change.Keys) { $prefs[$key] = $Change[$key] }
        $result = Test-YardstickPreferences -Preferences $prefs
        if ($Errors) { ($result.Errors -join "`n") | Should -Match $Errors } else { $result.Errors | Should -BeNullOrEmpty }
        if ($Warnings) { ($result.Warnings -join "`n") | Should -Match $Warnings }
    }
}
