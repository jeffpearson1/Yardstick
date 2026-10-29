class GithubDownloader {
    [string]$URL
    [string]$DownloaderRegex
    # Blank means the release's whole tag is the version.
    [string]$VersionRegex
    [string]$LatestVersion
    [string]$Filename
    [bool]$Prerelease = $false

    GithubDownloader() {}

    GithubDownloader([string]$URL, [string]$DownloaderRegex) {
        $this.URL = $URL
        $this.DownloaderRegex = $DownloaderRegex
    }

    GithubDownloader([string]$URL, [string]$DownloaderRegex, [string]$VersionRegex) {
        $this.URL = $URL
        $this.DownloaderRegex = $DownloaderRegex
        $this.VersionRegex = $VersionRegex
    }

    GithubDownloader([hashtable]$Properties) {
        foreach ($key in $Properties.Keys) {
            if (-not $this.PSObject.Properties[$key]) {
                throw "GithubDownloader has no property named '$key'."
            }
            $this.$key = $Properties[$key]
        }
    }

    [void] Update() {
        # Invoke-YardstickGitHubApi (YardstickGitHub.psm1, loaded globally with
        # Yardstick) attaches the stored token and handles GitHub's rate limits.
        $releasePath = "repos/$(($this.URL -replace "https://github.com/|/releases").trimEnd("/"))/releases"
        $releaseEndpoint = "https://api.github.com/$releasePath"
        $releaseResponse = Invoke-YardstickGitHubApi $releasePath
        $releaseLatest = $releaseResponse | Where-Object Prerelease -eq $this.Prerelease |
            Sort-Object published_at -Descending | Select-Object -First 1
        if (-not $releaseLatest) {
            $kind = if ($this.Prerelease) { 'prerelease' } else { 'release' }
            throw "No $kind was found at $releaseEndpoint."
        }

        if ([string]::IsNullOrEmpty($this.VersionRegex)) {
            $this.LatestVersion = $releaseLatest.tag_name
        }
        else {
            $match = [regex]::Match([string]$releaseLatest.tag_name, $this.VersionRegex, 'IgnoreCase')
            if (-not $match.Success) {
                throw "The latest release tag '$($releaseLatest.tag_name)' does not match the version pattern '$($this.VersionRegex)'."
            }
            $this.LatestVersion = $match.Value
        }

        $downloads = @($releaseLatest.assets | Where-Object browser_download_url -match $this.DownloaderRegex |
            ForEach-Object browser_download_url)
        if ($downloads.Count -ne 1) {
            $found = if ($downloads) { " It matches $($downloads.Count): $(($downloads | ForEach-Object { ($_ -split '/')[-1] }) -join ', ')." } else { '' }
            throw "Release '$($releaseLatest.tag_name)' must have exactly one file matching '$($this.DownloaderRegex)'.$found"
        }
        $this.URL = $downloads[0]
        $this.Filename = ($this.URL -split "/")[-1]
    }
}
