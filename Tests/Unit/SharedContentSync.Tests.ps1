BeforeAll {
    $Global:LogLocation = $TestDrive
    $Global:LogFile = "test.log"
    $script:repoRoot = (Resolve-Path "$PSScriptRoot\..\..").ProviderPath
    Import-Module powershell-yaml -ErrorAction Stop
    Import-Module (Join-Path $script:repoRoot 'Modules\YardstickSupport.psm1') -Force

    function New-TestFolder {
        param([string]$Name)
        (New-Item -ItemType Directory -Path (Join-Path $TestDrive "$Name-$([guid]::NewGuid().ToString('N'))")).FullName
    }

    function New-TestIcon {
        param([string]$Folder, [string]$Name, [string]$Content = 'icon', [datetime]$Modified = '2026-01-01')
        $path = Join-Path $Folder $Name
        Set-Content -LiteralPath $path -Value $Content -NoNewline
        (Get-Item -LiteralPath $path).LastWriteTimeUtc = $Modified
        $path
    }
}

Describe "Sync-YardstickSharedIcons" {
    BeforeEach {
        $script:share = New-TestFolder 'share'
        $script:local = New-TestFolder 'icons'
    }

    It "Copies an icon the local folder does not have" {
        New-TestIcon $script:share '7zip.png' | Out-Null
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        $result.Copied | Should -Be @('7zip.png')
        $result.Unchanged | Should -Be 0
        Get-Content -LiteralPath (Join-Path $script:local '7zip.png') -Raw | Should -Be 'icon'
    }

    It "Leaves an icon with the same size and date alone" {
        New-TestIcon $script:share '7zip.png' | Out-Null
        New-TestIcon $script:local '7zip.png' | Out-Null
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        $result.Copied | Should -BeNullOrEmpty
        $result.Unchanged | Should -Be 1
    }

    It "Replaces an icon that changed on the share" {
        New-TestIcon $script:share '7zip.png' -Content 'new icon' -Modified '2026-06-01' | Out-Null
        New-TestIcon $script:local '7zip.png' -Content 'old icon' | Out-Null
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        $result.Copied | Should -Be @('7zip.png')
        Get-Content -LiteralPath (Join-Path $script:local '7zip.png') -Raw | Should -Be 'new icon'
    }

    It "Copies only image files" {
        New-TestIcon $script:share 'a.png' | Out-Null
        New-TestIcon $script:share 'b.jpg' | Out-Null
        New-TestIcon $script:share 'c.jpeg' | Out-Null
        New-TestIcon $script:share 'Thumbs.db' | Out-Null
        New-TestIcon $script:share '.DS_Store' | Out-Null
        $result = Sync-YardstickSharedIcons -Source $script:share -Destination $script:local
        @($result.Copied | Sort-Object) | Should -Be @('a.png', 'b.jpg', 'c.jpeg')
        Test-Path -LiteralPath (Join-Path $script:local 'Thumbs.db') | Should -BeFalse
    }

    It "Keeps icons that exist only on this machine" {
        New-TestIcon $script:local 'cmtraceopen.png' | Out-Null
        Sync-YardstickSharedIcons -Source $script:share -Destination $script:local | Out-Null
        Test-Path -LiteralPath (Join-Path $script:local 'cmtraceopen.png') | Should -BeTrue
    }

    It "Leaves no temporary files behind" {
        New-TestIcon $script:share '7zip.png' | Out-Null
        Sync-YardstickSharedIcons -Source $script:share -Destination $script:local | Out-Null
        Get-ChildItem -LiteralPath $script:local -Filter '*.tmp' | Should -BeNullOrEmpty
    }

    It "Throws when the shared folder is not reachable" {
        { Sync-YardstickSharedIcons -Source (Join-Path $TestDrive 'no-such-share') -Destination $script:local } |
            Should -Throw '*not reachable*'
    }
}

Describe "Sync-YardstickSharedRecipeOverrides" {
    BeforeEach {
        $script:share = Join-Path (New-TestFolder 'share') 'recipeOverrides.yaml'
        $script:cache = Join-Path (New-TestFolder 'local') 'recipeOverrides.shared.yaml'
    }

    It "Copies a valid shared file to the cache" {
        Set-Content -LiteralPath $script:share -Value "7zip:`n  scopeTags:`n    - Labs`njasp:`n  scopeTags:`n    - Labs"
        Sync-YardstickSharedRecipeOverrides -Source $script:share -CachePath $script:cache | Should -Be 2
        (Import-RecipeOverrides -Path $script:cache)['7zip'].scopeTags | Should -Be @('Labs')
        Test-Path -LiteralPath "$script:cache.tmp" | Should -BeFalse
    }

    It "Keeps the last good copy when the shared file is malformed" {
        Set-Content -LiteralPath $script:cache -Value "7zip:`n  scopeTags:`n    - Labs"
        Set-Content -LiteralPath $script:share -Value "7zip: Labs"
        { Sync-YardstickSharedRecipeOverrides -Source $script:share -CachePath $script:cache } |
            Should -Throw '*malformed*'
        (Import-RecipeOverrides -Path $script:cache)['7zip'].scopeTags | Should -Be @('Labs')
        Test-Path -LiteralPath "$script:cache.tmp" | Should -BeFalse
    }

    It "Throws when the shared file is not reachable and leaves the cache alone" {
        Set-Content -LiteralPath $script:cache -Value "7zip:`n  scopeTags:`n    - Labs"
        { Sync-YardstickSharedRecipeOverrides -Source $script:share -CachePath $script:cache } |
            Should -Throw '*not reachable*'
        Test-Path -LiteralPath $script:cache | Should -BeTrue
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
