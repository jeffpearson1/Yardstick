class GithubDownloader {
    [string]$URL
    [string]$DownloaderRegex
    [string]$VersionRegex = $null
    [string]$LatestVersion
    [string]$Filename
    [bool]$Prerelease = $false

    GithubDownloader() { 
        $this.Init(@{}) 
    }

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
        $this.Init($Properties) 
    }

    [void] Update() {
        $releaseEndpoint = "https://api.github.com/repos/$(($this.URL -replace "https://github.com/|/releases").trimEnd("/"))/releases"
        $releaseResponse = Invoke-RestMethod $releaseEndpoint
        $releaseLatest = ($releaseResponse | Where-Object Prerelease -eq $this.Prerelease | Sort-Object published_at -Descending)[0]
        if ($null -ne $this.VersionRegex) {
            $releaseLatest.tag_name -match "$($this.VersionRegex)"
            $this.LatestVersion = $matches[0]
        }
        else {
            $this.LatestVersion = $releaseLatest.tag_name
        }
        $this.URL = ($releaseLatest[0].assets | Where-Object browser_download_url -match $this.DownloaderRegex).browser_download_url
        $this.Filename = ($this.URL -split "/")[-1]
    }
}