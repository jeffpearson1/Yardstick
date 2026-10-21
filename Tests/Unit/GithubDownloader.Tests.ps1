using module ..\..\Modules\Custom\GithubDownloader.psm1

BeforeAll {
    # The class calls Invoke-RestMethod from its own scope, where Mock cannot reach.
    # A global function of the same name comes ahead of the cmdlet.
    function New-TestRelease {
        param([string]$Tag, [string[]]$Assets, [string]$Published = '2026-10-01T00:00:00Z', [bool]$Prerelease = $false)
        [PSCustomObject]@{
            tag_name     = $Tag
            prerelease   = $Prerelease
            published_at = [datetime]$Published
            assets       = @($Assets | ForEach-Object {
                [PSCustomObject]@{ browser_download_url = "https://github.com/owner/app/releases/download/$Tag/$_" }
            })
        }
    }

    function global:Invoke-RestMethod {
        $Global:GithubDownloaderTestUri = $args[0]
        $Global:GithubDownloaderTestReleases
    }
}

AfterAll {
    Remove-Item Function:\Invoke-RestMethod -ErrorAction SilentlyContinue
    Remove-Variable GithubDownloaderTestReleases, GithubDownloaderTestUri -Scope Global -ErrorAction SilentlyContinue
}

Describe "GithubDownloader" {
    BeforeEach {
        $Global:GithubDownloaderTestReleases = @(
            New-TestRelease -Tag 'v2.0.0' -Assets 'app-2.0.0-x64.msi', 'app-2.0.0-arm64.msi' -Published '2026-10-01T00:00:00Z'
            New-TestRelease -Tag 'v1.9.0' -Assets 'app-1.9.0-x64.msi' -Published '2026-09-01T00:00:00Z'
            New-TestRelease -Tag 'v2.1.0-rc1' -Assets 'app-2.1.0-rc1-x64.msi' -Published '2026-10-10T00:00:00Z' -Prerelease $true
        )
    }

    It "Takes the newest release, its version and its one matching file" {
        $downloader = [GithubDownloader]::new('https://github.com/owner/app/releases', 'x64\.msi$', '\d+\.\d+\.\d+')
        $downloader.Update()
        $Global:GithubDownloaderTestUri | Should -Be 'https://api.github.com/repos/owner/app/releases'
        $downloader.LatestVersion | Should -Be '2.0.0'
        $downloader.Filename | Should -Be 'app-2.0.0-x64.msi'
        $downloader.URL | Should -Be 'https://github.com/owner/app/releases/download/v2.0.0/app-2.0.0-x64.msi'
    }

    It "Uses the whole tag as the version when no version pattern is given" {
        $downloader = [GithubDownloader]::new('https://github.com/owner/app', 'x64\.msi$')
        $downloader.Update()
        $downloader.LatestVersion | Should -Be 'v2.0.0'
    }

    It "Takes the newest prerelease when asked" {
        $downloader = [GithubDownloader]@{ URL = 'https://github.com/owner/app'; DownloaderRegex = 'x64\.msi$'; Prerelease = $true }
        $downloader.Update()
        $downloader.LatestVersion | Should -Be 'v2.1.0-rc1'
    }

    It "Rejects a property it does not have" {
        { [GithubDownloader]::new(@{ URL = 'https://github.com/owner/app'; DownloadRegex = 'x64' }) } |
            Should -Throw "*no property named 'DownloadRegex'*"
    }

    It "Fails when the tag does not match the version pattern" {
        $downloader = [GithubDownloader]::new('https://github.com/owner/app', 'x64\.msi$', '^release-\d+')
        { $downloader.Update() } | Should -Throw "*'v2.0.0' does not match the version pattern*"
    }

    It "Fails when more than one file matches, and names them" {
        $downloader = [GithubDownloader]::new('https://github.com/owner/app', '\.msi$')
        { $downloader.Update() } | Should -Throw "*It matches 2: app-2.0.0-x64.msi, app-2.0.0-arm64.msi*"
    }

    It "Fails when no file matches" {
        $downloader = [GithubDownloader]::new('https://github.com/owner/app', '\.zip$')
        { $downloader.Update() } | Should -Throw "*exactly one file matching*"
    }

    It "Fails when there is no release" {
        $Global:GithubDownloaderTestReleases = @()
        $downloader = [GithubDownloader]::new('https://github.com/owner/app', 'x64\.msi$')
        { $downloader.Update() } | Should -Throw "No release was found at *"
    }
}
