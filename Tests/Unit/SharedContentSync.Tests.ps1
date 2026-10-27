BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    Import-Module powershell-yaml -ErrorAction Stop
    Import-Module (Join-Path $script:repoRoot 'Modules\YardstickSupport.psm1') -Force
    Add-Type -AssemblyName System.Drawing

    function New-TestFolder {
        param([string]$Name)
        (New-Item -ItemType Directory -Path (Join-Path $TestDrive "$Name-$([guid]::NewGuid().ToString('N'))")).FullName
    }

    # A real image, so it passes Get-YardstickImageProblem. A larger Size gives a
    # different file.
    function New-TestIcon {
        param([string]$Folder, [string]$Name, [int]$Size = 1, [datetime]$Modified = '2026-01-01')
        $path = Join-Path $Folder $Name
        $bitmap = [System.Drawing.Bitmap]::new($Size, $Size)
        try {
            $format = if ($Name -match '\.jpe?g$') { [System.Drawing.Imaging.ImageFormat]::Jpeg } else { [System.Drawing.Imaging.ImageFormat]::Png }
            $bitmap.Save($path, $format)
        } finally {
            $bitmap.Dispose()
        }
        (Get-Item -LiteralPath $path).LastWriteTimeUtc = $Modified
        $path
    }

    function Set-TestFile {
        param([string]$Path, [string[]]$Line)
        [IO.File]::WriteAllText($Path, (($Line -join "`n") + "`n"))
    }
}

Describe "Get-YardstickImageProblem" {
    BeforeEach {
        $script:folder = New-TestFolder 'images'
    }

    It "Accepts a complete PNG and JPEG" {
        Get-YardstickImageProblem -Path (New-TestIcon $script:folder 'a.png' -Size 4) | Should -BeNullOrEmpty
        Get-YardstickImageProblem -Path (New-TestIcon $script:folder 'b.jpg' -Size 4) | Should -BeNullOrEmpty
    }

    It "Reports a PNG that was cut short" {
        $path = New-TestIcon $script:folder 'a.png' -Size 16
        $bytes = [IO.File]::ReadAllBytes($path)
        [IO.File]::WriteAllBytes($path, $bytes[0..($bytes.Length - 13)])
        Get-YardstickImageProblem -Path $path | Should -BeLike '*cut short*'
    }

    It "Reports a JPEG that was cut short" {
        $path = New-TestIcon $script:folder 'a.jpg' -Size 16
        $bytes = [IO.File]::ReadAllBytes($path)
        [IO.File]::WriteAllBytes($path, $bytes[0..($bytes.Length - 20)])
        Get-YardstickImageProblem -Path $path | Should -BeLike '*cut short*'
    }

    It "Reports an empty file and a file that is not an image" {
        $empty = Join-Path $script:folder 'empty.png'
        [IO.File]::WriteAllBytes($empty, [byte[]]@())
        Get-YardstickImageProblem -Path $empty | Should -BeLike '*empty*'
        $text = Join-Path $script:folder 'text.png'
        Set-Content -LiteralPath $text -Value '<html>404</html>'
        Get-YardstickImageProblem -Path $text | Should -BeLike '*not a PNG or JPEG*'
    }
}

Describe "Sync-YardstickSharedIcons" {
    BeforeEach {
        $script:share = New-TestFolder 'share'
        $script:local = New-TestFolder 'icons'
    }

    It "Copies an icon the local folder does not have" {
        $source = New-TestIcon $script:share '7zip.png' -Size 3
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        $result.Pulled | Should -Be @('7zip.png')
        $result.Pushed | Should -BeNullOrEmpty
        $copy = Join-Path $script:local '7zip.png'
        (Get-FileHash $copy).Hash | Should -Be (Get-FileHash $source).Hash
        (Get-Item $copy).LastWriteTimeUtc | Should -Be (Get-Item $source).LastWriteTimeUtc
    }

    It "Copies an icon the share does not have" {
        New-TestIcon $script:local 'cmtraceopen.png' | Out-Null
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        $result.Pushed | Should -Be @('cmtraceopen.png')
        Test-Path -LiteralPath (Join-Path $script:share 'cmtraceopen.png') | Should -BeTrue
    }

    It "Leaves an icon with the same size and date alone" {
        New-TestIcon $script:share '7zip.png' | Out-Null
        New-TestIcon $script:local '7zip.png' | Out-Null
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        $result.Pulled | Should -BeNullOrEmpty
        $result.Pushed | Should -BeNullOrEmpty
        $result.Unchanged | Should -Be 1
    }

    It "Replaces the older copy with the newer one, in either direction" {
        $newShared = New-TestIcon $script:share '7zip.png' -Size 5 -Modified '2026-06-01'
        New-TestIcon $script:local '7zip.png' -Size 2 | Out-Null
        New-TestIcon $script:share 'jasp.png' -Size 2 | Out-Null
        $newLocal = New-TestIcon $script:local 'jasp.png' -Size 6 -Modified '2026-06-01'
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        $result.Pulled | Should -Be @('7zip.png')
        $result.Pushed | Should -Be @('jasp.png')
        (Get-FileHash (Join-Path $script:local '7zip.png')).Hash | Should -Be (Get-FileHash $newShared).Hash
        (Get-FileHash (Join-Path $script:share 'jasp.png')).Hash | Should -Be (Get-FileHash $newLocal).Hash
    }

    It "Leaves both copies alone when they differ but share a date" {
        New-TestIcon $script:share '7zip.png' -Size 5 | Out-Null
        $local = New-TestIcon $script:local '7zip.png' -Size 2
        $before = (Get-FileHash $local).Hash
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        $result.Failed.Name | Should -Be @('7zip.png')
        (Get-FileHash $local).Hash | Should -Be $before
    }

    It "Never copies a damaged icon, and never replaces a good one with it" {
        $good = New-TestIcon $script:local '7zip.png' -Size 4
        $before = (Get-FileHash $good).Hash
        $bad = New-TestIcon $script:share '7zip.png' -Size 16 -Modified '2026-06-01'
        $bytes = [IO.File]::ReadAllBytes($bad)
        [IO.File]::WriteAllBytes($bad, $bytes[0..($bytes.Length - 13)])
        (Get-Item $bad).LastWriteTimeUtc = '2026-06-01'
        Set-Content -LiteralPath (Join-Path $script:local 'notes.png') -Value 'not an image'

        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        @($result.Failed.Name | Sort-Object) | Should -Be @('7zip.png', 'notes.png')
        (Get-FileHash $good).Hash | Should -Be $before
        Test-Path -LiteralPath (Join-Path $script:share 'notes.png') | Should -BeFalse
    }

    It "Considers only image files and deletes nothing" {
        New-TestIcon $script:share 'a.png' | Out-Null
        New-TestIcon $script:share 'b.jpg' | Out-Null
        New-TestIcon $script:local 'c.jpeg' | Out-Null
        Set-Content -LiteralPath (Join-Path $script:share 'Thumbs.db') -Value 'x'
        Set-Content -LiteralPath (Join-Path $script:local '.gitkeep') -Value ''
        Sync-YardstickSharedIcons -Source $script:share -Destination $script:local | Out-Null
        @(Get-ChildItem -LiteralPath $script:local -Force | ForEach-Object Name | Sort-Object) | Should -Be @('.gitkeep', 'a.png', 'b.jpg', 'c.jpeg')
        @(Get-ChildItem -LiteralPath $script:share -Force | ForEach-Object Name | Sort-Object) | Should -Be @('a.png', 'b.jpg', 'c.jpeg', 'Thumbs.db')
    }

    It "Leaves no temporary files behind" {
        New-TestIcon $script:share '7zip.png' | Out-Null
        New-TestIcon $script:local 'jasp.png' | Out-Null
        Sync-YardstickSharedIcons -Source $script:share -Destination $script:local | Out-Null
        Get-ChildItem -LiteralPath $script:local, $script:share -Filter '*.tmp' -Force | Should -BeNullOrEmpty
    }

    It "Throws when the shared folder is not reachable" {
        { Sync-YardstickSharedIcons -Source (Join-Path $TestDrive 'no-such-share') -Destination $script:local } |
            Should -Throw '*not reachable*'
    }
}

Describe "Sync-YardstickSharedRecipeOverrides" {
    BeforeEach {
        $script:shareFolder = New-TestFolder 'share'
        $script:share = Join-Path $script:shareFolder 'recipeOverrides.yaml'
        $localFolder = New-TestFolder 'local'
        $script:local = Join-Path $localFolder 'recipeOverrides.yaml'
        $script:cache = Join-Path $localFolder 'recipeOverrides.shared.yaml'

        function Invoke-Sync {
            Sync-YardstickSharedRecipeOverrides -LocalPath $script:local -SharedPath $script:share -CachePath $script:cache
        }
        function Get-Entries {
            param([string]$Path)
            Import-RecipeOverrides -Path $Path
        }
    }

    It "Adds local entries to the shared file in order, keeping its comments" {
        Set-TestFile $script:share '# Team overrides', '7zip:', '  scopeTags:', "    - 'Labs'", 'zoom:', '  scopeTags:', "    - 'Labs'"
        Set-TestFile $script:local '# Mine', 'jasp:', '  scopeTags:', "    - 'Mine'"

        $result = Invoke-Sync
        $result.Pushed | Should -Be @('jasp: scopeTags')
        $result.SharedEntries | Should -Be 3
        Get-Content -LiteralPath $script:share | Should -Be @(
            '# Team overrides', '7zip:', '  scopeTags:', "    - 'Labs'", 'jasp:', '  scopeTags:', "    - 'Mine'", 'zoom:', '  scopeTags:', "    - 'Labs'")
        (Get-FileHash $script:cache).Hash | Should -Be (Get-FileHash $script:share).Hash
        Test-Path -LiteralPath "$script:share.bak" | Should -BeTrue
    }

    It "Does nothing on a second sync" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'Labs'"
        Set-TestFile $script:local 'jasp:', '  scopeTags:', "    - 'Mine'"
        Invoke-Sync | Out-Null
        $before = (Get-FileHash $script:share).Hash
        $result = Invoke-Sync
        $result.Pushed + $result.Pulled + $result.Removed + $result.Conflicts | Should -BeNullOrEmpty
        (Get-FileHash $script:share).Hash | Should -Be $before
    }

    It "Creates the shared file when the folder has none yet" {
        Set-TestFile $script:local 'jasp:', '  scopeTags:', "    - 'Mine'"
        $result = Invoke-Sync
        $result.Pushed | Should -Be @('jasp: scopeTags')
        (Get-Entries $script:share)['jasp'].scopeTags | Should -Be @('Mine')
    }

    It "Brings a shared change into the local file, keeping its comments" {
        Set-TestFile $script:share '7zip:', '  scopeTags:', "    - 'Old'"
        Set-TestFile $script:local '# My overrides', '7zip:', '  scopeTags:', "    - 'Old'"
        Invoke-Sync | Out-Null
        Set-TestFile $script:share '7zip:', '  scopeTags:', "    - 'New'", '    - Second'

        $result = Invoke-Sync
        $result.Pulled | Should -Be @('7zip: scopeTags')
        Get-Content -LiteralPath $script:local | Should -Be @('# My overrides', '7zip:', '  scopeTags:', "    - 'New'", "    - 'Second'")
    }

    It "Does not copy entries that only the shared file has into the local file" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'Labs'"
        Set-TestFile $script:local 'jasp:', '  scopeTags:', "    - 'Mine'"
        Invoke-Sync | Out-Null
        (Get-Entries $script:local).Keys | Should -Be @('jasp')
    }

    It "Sends a local change made after the last sync" {
        Set-TestFile $script:share '7zip:', '  scopeTags:', "    - 'Old'", '  notes: shared'
        Set-TestFile $script:local '7zip:', '  scopeTags:', "    - 'Old'"
        Invoke-Sync | Out-Null
        Set-TestFile $script:local '7zip:', '  scopeTags:', "    - 'New'"

        $result = Invoke-Sync
        $result.Pushed | Should -Be @('7zip: scopeTags')
        $entry = (Get-Entries $script:share)['7zip']
        $entry.scopeTags | Should -Be @('New')
        $entry.notes | Should -Be 'shared'
    }

    It "Removes locally a field or entry that was removed from the shared file" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'", "  notes: 'x'", 'jasp:', '  scopeTags:', "    - 'A'", 'keep:', '  notes: kept'
        Set-TestFile $script:local 'zoom:', '  scopeTags:', "    - 'A'", "  notes: 'x'", 'jasp:', '  scopeTags:', "    - 'A'"
        Invoke-Sync | Out-Null
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'", 'keep:', '  notes: kept'

        $result = Invoke-Sync
        @($result.Removed | Sort-Object) | Should -Be @('jasp: scopeTags', 'zoom: notes')
        $local = Get-Entries $script:local
        $local.Keys | Should -Be @('zoom')
        $local['zoom'].Keys | Should -Be @('scopeTags')
    }

    It "Never removes an entry from the shared file because it was removed locally" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'"
        Set-TestFile $script:local 'zoom:', '  scopeTags:', "    - 'A'"
        Invoke-Sync | Out-Null
        Set-TestFile $script:local '# nothing here'
        Invoke-Sync | Out-Null
        (Get-Entries $script:share)['zoom'].scopeTags | Should -Be @('A')
    }

    It "Reports a field changed on both sides until they agree, changing neither" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'"
        Set-TestFile $script:local 'zoom:', '  scopeTags:', "    - 'A'"
        Invoke-Sync | Out-Null
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'Shared'"
        Set-TestFile $script:local 'zoom:', '  scopeTags:', "    - 'Local'"

        (Invoke-Sync).Conflicts | Should -Be @('zoom: scopeTags')
        (Invoke-Sync).Conflicts | Should -Be @('zoom: scopeTags')
        (Get-Entries $script:share)['zoom'].scopeTags | Should -Be @('Shared')
        (Get-Entries $script:local)['zoom'].scopeTags | Should -Be @('Local')

        Set-TestFile $script:local 'zoom:', '  scopeTags:', "    - 'Shared'"
        (Invoke-Sync).Conflicts | Should -BeNullOrEmpty
        Set-TestFile $script:local 'zoom:', '  scopeTags:', "    - 'Later'"
        (Invoke-Sync).Pushed | Should -Be @('zoom: scopeTags')
    }

    It "Writes values that need quoting so they read back the same" {
        Set-TestFile $script:share '# empty'
        Set-TestFile $script:local "o'brien:", '  scopeTags:', "    - 'O''Brien Lab'", '    - "yes"', '  allowUserUninstall: false',
            '  minimumFreeDiskSpaceInMB: 1024', "  notes: 'a: b # c'", '  requirementRule:', '    type: registry', '    keys: [a, b]'
        Invoke-Sync | Out-Null
        $shared = Get-Entries $script:share
        $local = Get-Entries $script:local
        (ConvertTo-YardstickCanonicalJson $shared) | Should -Be (ConvertTo-YardstickCanonicalJson $local)
    }

    It "Keeps the local file's line endings" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'"
        [IO.File]::WriteAllText($script:local, "zoom:`r`n  scopeTags:`r`n    - 'A'`r`n")
        Invoke-Sync | Out-Null
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'B'"
        Invoke-Sync | Out-Null
        [IO.File]::ReadAllText($script:local) | Should -Be "zoom:`r`n  scopeTags:`r`n    - 'B'`r`n"
    }

    It "Changes nothing when the shared file is malformed" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'"
        Set-TestFile $script:local 'jasp:', '  scopeTags:', "    - 'Mine'"
        Invoke-Sync | Out-Null
        $cacheBefore = (Get-FileHash $script:cache).Hash
        Set-TestFile $script:share 'zoom: [unclosed'
        Set-TestFile $script:local 'jasp:', '  scopeTags:', "    - 'Changed'"

        { Invoke-Sync } | Should -Throw '*malformed*'
        (Get-FileHash $script:cache).Hash | Should -Be $cacheBefore
        (Get-Entries $script:local)['jasp'].scopeTags | Should -Be @('Changed')
    }

    It "Changes nothing when the local file is malformed" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'"
        Set-TestFile $script:local 'jasp: Mine'
        $before = (Get-FileHash $script:share).Hash
        { Invoke-Sync } | Should -Throw '*local overrides file*malformed*'
        (Get-FileHash $script:share).Hash | Should -Be $before
        Test-Path -LiteralPath $script:cache | Should -BeFalse
    }

    It "Treats recipe ids that differ only in case as malformed" {
        Set-TestFile $script:share 'zoom:', '  notes: a', 'Zoom:', '  notes: b'
        Set-TestFile $script:local '# none'
        { Invoke-Sync } | Should -Throw '*more than once*'
        Set-TestFile $script:share 'zoom:', '  notes: a', '  requirementRule:', '    notes: nested is fine', '  Notes: b'
        { Invoke-Sync } | Should -Throw "*sets 'Notes' more than once*"
    }

    It "Leaves everything alone when the shared file was emptied after a sync" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'"
        Set-TestFile $script:local 'jasp:', '  scopeTags:', "    - 'Mine'"
        Invoke-Sync | Out-Null
        Set-TestFile $script:share '# oops'
        { Invoke-Sync } | Should -Throw '*is now empty*'
        Get-Content -LiteralPath $script:share | Should -Be @('# oops')
        (Get-Entries $script:cache).Keys | Should -Contain 'zoom'
    }

    It "Throws when the shared folder is not reachable and leaves the cache alone" {
        Set-TestFile $script:cache 'zoom:', '  scopeTags:', "    - 'A'"
        Set-TestFile $script:local '# none'
        $script:share = Join-Path $TestDrive 'no-such-share\recipeOverrides.yaml'
        { Invoke-Sync } | Should -Throw '*not reachable*'
        (Get-Entries $script:cache)['zoom'].scopeTags | Should -Be @('A')
    }

    It "Leaves no temporary or lock files behind" {
        Set-TestFile $script:share 'zoom:', '  scopeTags:', "    - 'A'"
        Set-TestFile $script:local 'jasp:', '  scopeTags:', "    - 'Mine'"
        Invoke-Sync | Out-Null
        Get-ChildItem -LiteralPath $script:shareFolder, (Split-Path $script:local) -Force |
            Where-Object { $_.Name -like '*.tmp' -or $_.Name -like '*.lock' } | Should -BeNullOrEmpty
    }
}

Describe "Write-YardstickFileSafely" {
    It "Refuses to replace a file that changed since it was read" {
        $path = Join-Path (New-TestFolder 'safe') 'file.yaml'
        Set-Content -LiteralPath $path -Value 'original'
        { Write-YardstickFileSafely -Path $path -Content ([byte[]][char[]]'new') -ExpectedHash ('0' * 64) } |
            Should -Throw '*changed by someone else*'
        Get-Content -LiteralPath $path | Should -Be 'original'
        Get-ChildItem -LiteralPath (Split-Path $path) -Force | ForEach-Object Name | Should -Be @('file.yaml')
    }
}

Describe "Invoke-YardstickSharedContentSync" {
    BeforeEach {
        $folder = New-TestFolder 'run'
        $script:localOverrides = Join-Path $folder 'recipeOverrides.yaml'
        $script:cache = Join-Path $folder 'recipeOverrides.shared.yaml'
        $script:prefs = @{
            syncSharedContent     = $true
            Icons                 = (New-TestFolder 'icons')
            SharedIcons           = (Join-Path $TestDrive 'missing-icons')
            SharedRecipeOverrides = (Join-Path $TestDrive 'missing-share\recipeOverrides.yaml')
        }
    }

    It "Does nothing when shared sync is off" {
        $script:prefs.syncSharedContent = $false
        Invoke-YardstickSharedContentSync -Preferences $script:prefs -LocalOverridesPath $script:localOverrides -SharedOverridesCachePath $script:cache |
            Should -BeNullOrEmpty
    }

    It "Turns every failure into a warning instead of throwing" {
        Set-TestFile $script:cache 'zoom:', '  scopeTags:', "    - 'A'"
        $messages = @(Invoke-YardstickSharedContentSync -Preferences $script:prefs -LocalOverridesPath $script:localOverrides -SharedOverridesCachePath $script:cache)
        $messages.Count | Should -Be 2
        $messages.Level | Should -Be @('Warning', 'Warning')
        $messages[0].Message | Should -BeLike '*icons*not reachable*'
        $messages[1].Message | Should -BeLike '*Using the shared copy synced on*'
    }

    It "Warns about a conflict and reports what was synced" {
        $shareFolder = New-TestFolder 'share'
        $script:prefs.SharedIcons = $null
        $script:prefs.SharedRecipeOverrides = Join-Path $shareFolder 'recipeOverrides.yaml'
        Set-TestFile $script:prefs.SharedRecipeOverrides 'zoom:', '  scopeTags:', "    - 'A'"
        Set-TestFile $script:localOverrides 'zoom:', '  scopeTags:', "    - 'A'"
        Invoke-YardstickSharedContentSync -Preferences $script:prefs -LocalOverridesPath $script:localOverrides -SharedOverridesCachePath $script:cache | Out-Null
        Set-TestFile $script:prefs.SharedRecipeOverrides 'zoom:', '  scopeTags:', "    - 'B'"
        Set-TestFile $script:localOverrides 'zoom:', '  scopeTags:', "    - 'C'"

        $messages = @(Invoke-YardstickSharedContentSync -Preferences $script:prefs -LocalOverridesPath $script:localOverrides -SharedOverridesCachePath $script:cache)
        $messages[0].Level | Should -Be 'Info'
        $messages[1].Level | Should -Be 'Warning'
        $messages[1].Message | Should -BeLike '*zoom: scopeTags*changed both here and in the shared file*'
    }
}

Describe "Merge-RecipeOverrideSet" {
    It "Lets a local field beat the shared field and keeps the other shared fields" {
        $shared = @{ zoom = @{ scopeTags = @('Shared'); url = 'https://shared.zoom.us' } }
        $local = @{ zoom = @{ scopeTags = @('Local') } }
        $merged = Merge-RecipeOverrideSet -Shared $shared -Local $local
        $merged['zoom'].scopeTags | Should -Be @('Local')
        $merged['zoom'].url | Should -Be 'https://shared.zoom.us'
    }

    It "Passes through ids that only one side has" {
        $shared = @{ '7zip' = @{ scopeTags = @('Shared') } }
        $local = @{ jasp = @{ scopeTags = @('Local') } }
        $merged = Merge-RecipeOverrideSet -Shared $shared -Local $local
        $merged.Keys | Sort-Object | Should -Be @('7zip', 'jasp')
        $merged['7zip'].scopeTags | Should -Be @('Shared')
        $merged['jasp'].scopeTags | Should -Be @('Local')
    }

    It "Matches recipe ids regardless of case" {
        $shared = @{ AmazonCorretto23 = @{ scopeTags = @('Shared'); notes = 'shared' } }
        $local = @{ amazoncorretto23 = @{ scopeTags = @('Local') } }
        $merged = Merge-RecipeOverrideSet -Shared $shared -Local $local
        $merged.Count | Should -Be 1
        $merged['AMAZONCORRETTO23'].scopeTags | Should -Be @('Local')
        $merged['AMAZONCORRETTO23'].notes | Should -Be 'shared'
    }

    It "Does not change the shared entries it was given" {
        $shared = @{ zoom = @{ scopeTags = @('Shared') } }
        Merge-RecipeOverrideSet -Shared $shared -Local @{ zoom = @{ scopeTags = @('Local') } } | Out-Null
        $shared['zoom'].scopeTags | Should -Be @('Shared')
    }
}

Describe "Import-YardstickEffectiveRecipeOverrides" {
    BeforeEach {
        $folder = New-TestFolder 'local'
        $script:localPath = Join-Path $folder 'recipeOverrides.yaml'
        $script:cachePath = Join-Path $folder 'recipeOverrides.shared.yaml'
        Set-Content -LiteralPath $script:localPath -Value "jasp:`n  scopeTags:`n    - Local"
        Set-Content -LiteralPath $script:cachePath -Value "jasp:`n  scopeTags:`n    - Shared`n7zip:`n  scopeTags:`n    - Shared"
    }

    It "Reads only the local file when no shared copy is passed" {
        $result = Import-YardstickEffectiveRecipeOverrides -Path $script:localPath
        $result.Overrides.Keys | Should -Be @('jasp')
        $result.Sources | Should -Be @($script:localPath)
    }

    It "Reads only the local file when the shared copy has not been synced yet" {
        Remove-Item -LiteralPath $script:cachePath
        $result = Import-YardstickEffectiveRecipeOverrides -Path $script:localPath -SharedPath $script:cachePath
        $result.Overrides.Keys | Should -Be @('jasp')
    }

    It "Layers the local file over the shared copy" {
        $result = Import-YardstickEffectiveRecipeOverrides -Path $script:localPath -SharedPath $script:cachePath
        $result.Overrides['jasp'].scopeTags | Should -Be @('Local')
        $result.Overrides['7zip'].scopeTags | Should -Be @('Shared')
        $result.Sources | Should -Be @($script:cachePath, $script:localPath)
    }

    It "Uses the shared copy alone when there is no local file" {
        Remove-Item -LiteralPath $script:localPath
        $result = Import-YardstickEffectiveRecipeOverrides -Path $script:localPath -SharedPath $script:cachePath
        $result.Overrides['jasp'].scopeTags | Should -Be @('Shared')
        $result.Sources | Should -Be @($script:cachePath)
    }

    It "Throws on a malformed shared copy" {
        Set-Content -LiteralPath $script:cachePath -Value "jasp: Shared"
        { Import-YardstickEffectiveRecipeOverrides -Path $script:localPath -SharedPath $script:cachePath } |
            Should -Throw '*shared overrides copy*'
    }

    It "Throws on a malformed local file" {
        Set-Content -LiteralPath $script:localPath -Value "jasp: Local"
        { Import-YardstickEffectiveRecipeOverrides -Path $script:localPath -SharedPath $script:cachePath } |
            Should -Throw
    }
}
