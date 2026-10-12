BeforeAll {
    . (Join-Path $PSScriptRoot '..\..\Get-YardstickIcon.ps1') -LoadOnly

    function New-TestBitmap {
        # 40x30 white canvas with a 20x20 red square at (10,5)
        $bmp = New-Object System.Drawing.Bitmap 40, 30
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.FillRectangle([System.Drawing.Brushes]::Red, 10, 5, 20, 20)
        $g.Dispose()
        $argb = [YardstickIconOps]::ToArgb($bmp)
        $bmp.Dispose()
        $argb
    }

    function New-BingAnchor {
        param([string]$Murl, [string]$Turl = 'https://tse1.mm.bing.net/th?id=OIP.x&pid=15.1', [string]$Title = 'Result', [string]$Dims = '512&#215;512')
        $json = (@{ murl = $Murl; turl = $Turl; t = $Title } | ConvertTo-Json -Compress)
        $encoded = [System.Net.WebUtility]::HtmlEncode($json)
        "<a aria-label=`"Image result`" role=`"link`" class=`"iusc`" style=`"height:190px`" m=`"$encoded`"><img/></a><div class=`"img_info hon`"><span class=`"nowrap`">$Dims</span></div>"
    }

    function New-TestIco {
        # ICO container with one PNG frame per size, in the given order.
        param([int[]]$Sizes)
        $frames = foreach ($s in $Sizes) {
            $bmp = New-Object System.Drawing.Bitmap $s, $s
            $ms = New-Object System.IO.MemoryStream
            $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
            $bmp.Dispose()
            , $ms.ToArray()
        }
        $out = New-Object System.IO.MemoryStream
        $w = New-Object System.IO.BinaryWriter $out
        $w.Write([uint16]0); $w.Write([uint16]1); $w.Write([uint16]$Sizes.Count)
        $offset = 6 + 16 * $Sizes.Count
        for ($i = 0; $i -lt $Sizes.Count; $i++) {
            $w.Write([byte]($Sizes[$i] % 256)); $w.Write([byte]($Sizes[$i] % 256)); $w.Write([byte]0); $w.Write([byte]0)
            $w.Write([uint16]1); $w.Write([uint16]32)
            $w.Write([uint32]$frames[$i].Length); $w.Write([uint32]$offset)
            $offset += $frames[$i].Length
        }
        foreach ($f in $frames) { $w.Write($f) }
        $w.Flush()
        , $out.ToArray()
    }
}

Describe 'ConvertFrom-BingImageHtml' {
    It 'extracts full url, thumbnail, title and dimensions from each result' {
        $html = New-BingAnchor -Murl 'https://example.com/a.png' -Title 'First &amp; best' -Dims '600&#215;400'
        $result = @(ConvertFrom-BingImageHtml -Html $html)
        $result.Count | Should -Be 1
        $result[0].FullUrl | Should -Be 'https://example.com/a.png'
        $result[0].ThumbUrl | Should -Match 'bing\.net'
        $result[0].Width | Should -Be 600
        $result[0].Height | Should -Be 400
    }

    It 'skips duplicate full-size urls' {
        $html = (New-BingAnchor -Murl 'https://example.com/a.png') + (New-BingAnchor -Murl 'https://example.com/a.png') + (New-BingAnchor -Murl 'https://example.com/b.png')
        @(ConvertFrom-BingImageHtml -Html $html).Count | Should -Be 2
    }

    It 'stops at the requested maximum, keeping Bing ranking order' {
        $html = -join (1..30 | ForEach-Object { New-BingAnchor -Murl "https://example.com/$_.png" })
        $result = @(ConvertFrom-BingImageHtml -Html $html -Max 20)
        $result.Count | Should -Be 20
        $result[0].FullUrl | Should -Be 'https://example.com/1.png'
        $result[19].FullUrl | Should -Be 'https://example.com/20.png'
    }

    It 'returns nothing for a page with no results' {
        @(ConvertFrom-BingImageHtml -Html '<html><body>blocked</body></html>').Count | Should -Be 0
    }

    It 'ignores results whose metadata is not valid json' {
        $html = '<a class="iusc" m="{not json"></a>' + (New-BingAnchor -Murl 'https://example.com/ok.png')
        $result = @(ConvertFrom-BingImageHtml -Html $html)
        $result.Count | Should -Be 1
        $result[0].FullUrl | Should -Be 'https://example.com/ok.png'
    }
}

Describe 'Get-BingImageSearchUrl' {
    It 'escapes the query and requests square results' {
        $url = Get-BingImageSearchUrl -Query 'notepad++ app icon'
        $url | Should -Match 'q=notepad%2B%2B%20app%20icon'
        $url | Should -Match 'filterui:aspect-square'
        $url | Should -Not -Match 'photo-transparent'
    }

    It 'adds the transparent filter on request' {
        Get-BingImageSearchUrl -Query 'x' -TransparentOnly | Should -Match 'filterui:photo-transparent'
    }
}

Describe 'YardstickIconOps' {
    It 'finds the bounds of the content when the border is a uniform colour' {
        $bmp = New-TestBitmap
        $r = [YardstickIconOps]::TrimBounds($bmp, 10)
        "$($r.X),$($r.Y),$($r.Width),$($r.Height)" | Should -Be '10,5,20,20'
    }

    It 'trims transparent borders' {
        $bmp = New-TestBitmap
        $clear = [YardstickIconOps]::RemoveColor($bmp, 0, 0, 10, $true)
        $r = [YardstickIconOps]::TrimBounds($clear, 10)
        "$($r.X),$($r.Y),$($r.Width),$($r.Height)" | Should -Be '10,5,20,20'
    }

    It 'returns the full image when there is nothing to trim away' {
        $bmp = New-Object System.Drawing.Bitmap 8, 8
        $r = [YardstickIconOps]::TrimBounds([YardstickIconOps]::ToArgb($bmp), 10)
        $r.Width | Should -Be 8
        $r.Height | Should -Be 8
    }

    It 'crops to the requested rectangle and clamps to the image' {
        $bmp = New-TestBitmap
        $c = [YardstickIconOps]::Crop($bmp, (New-Object System.Drawing.Rectangle 30, 20, 50, 50))
        $c.Width | Should -Be 10
        $c.Height | Should -Be 10
    }

    It 'flood-fills only the contiguous background from the clicked pixel' {
        $bmp = New-TestBitmap
        $result = [YardstickIconOps]::RemoveColor($bmp, 0, 0, 10, $true)
        $result.GetPixel(0, 0).A | Should -Be 0
        $result.GetPixel(15, 10).A | Should -Be 255
    }

    It 'does not remove an enclosed region of the same colour when contiguous' {
        $bmp = New-Object System.Drawing.Bitmap 20, 20
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.FillRectangle([System.Drawing.Brushes]::Black, 5, 5, 10, 10)
        $g.FillRectangle([System.Drawing.Brushes]::White, 8, 8, 4, 4)
        $g.Dispose()
        $argb = [YardstickIconOps]::ToArgb($bmp)
        $contiguous = [YardstickIconOps]::RemoveColor($argb, 0, 0, 10, $true)
        $contiguous.GetPixel(9, 9).A | Should -Be 255
        $everything = [YardstickIconOps]::RemoveColor($argb, 0, 0, 10, $false)
        $everything.GetPixel(9, 9).A | Should -Be 0
    }

    It 'respects tolerance when matching colours' {
        $bmp = New-Object System.Drawing.Bitmap 4, 1
        $bmp.SetPixel(0, 0, [System.Drawing.Color]::FromArgb(255, 100, 100, 100))
        $bmp.SetPixel(1, 0, [System.Drawing.Color]::FromArgb(255, 110, 110, 110))
        $bmp.SetPixel(2, 0, [System.Drawing.Color]::FromArgb(255, 160, 160, 160))
        $bmp.SetPixel(3, 0, [System.Drawing.Color]::FromArgb(255, 100, 100, 100))
        $argb = [YardstickIconOps]::ToArgb($bmp)
        $result = [YardstickIconOps]::RemoveColor($argb, 0, 0, 15, $false)
        $result.GetPixel(1, 0).A | Should -Be 0
        $result.GetPixel(2, 0).A | Should -Be 255
        $result.GetPixel(3, 0).A | Should -Be 0
    }

    It 'removes white with RemoveWhite' {
        $result = [YardstickIconOps]::RemoveWhite((New-TestBitmap), 10)
        $result.GetPixel(0, 0).A | Should -Be 0
        $result.GetPixel(15, 10).A | Should -Be 255
    }

    It 'resizes so the longest side matches, keeping the aspect ratio' {
        $r = [YardstickIconOps]::Resize((New-TestBitmap), 80, $false)
        $r.Width | Should -Be 80
        $r.Height | Should -Be 60
    }

    It 'pads to a square canvas with transparency' {
        $r = [YardstickIconOps]::Resize((New-TestBitmap), 80, $true)
        $r.Width | Should -Be 80
        $r.Height | Should -Be 80
        $r.GetPixel(40, 2).A | Should -Be 0
        $r.GetPixel(40, 40).A | Should -Be 255
    }

    It 'composites over a background colour and rounds the tile corners' {
        $clear = [YardstickIconOps]::RemoveWhite((New-TestBitmap), 10)
        $flat = [YardstickIconOps]::AddBackground($clear, [System.Drawing.Color]::Blue, $false)
        $flat.GetPixel(0, 0).A | Should -Be 255
        $flat.GetPixel(0, 0).B | Should -Be 255
        $rounded = [YardstickIconOps]::AddBackground($clear, [System.Drawing.Color]::Blue, $true)
        $rounded.GetPixel(0, 0).A | Should -Be 0
        $rounded.GetPixel(20, 15).A | Should -Be 255
    }

    It 'round-trips through a PNG file' {
        $path = Join-Path $TestDrive 'icon.png'
        (New-TestBitmap).Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $img = [System.Drawing.Image]::FromFile($path)
        try {
            $img.Width | Should -Be 40
            $img.RawFormat.Guid | Should -Be ([System.Drawing.Imaging.ImageFormat]::Png.Guid)
        } finally { $img.Dispose() }
    }

    It 'fades anti-aliased edge pixels instead of leaving a halo' {
        # Colours within 40 of white go; between 40 and 80 they fade; further out they stay.
        $bmp = New-Object System.Drawing.Bitmap 4, 1
        $bmp.SetPixel(0, 0, [System.Drawing.Color]::White)
        $bmp.SetPixel(1, 0, [System.Drawing.Color]::FromArgb(255, 255, 195, 195))
        $bmp.SetPixel(2, 0, [System.Drawing.Color]::FromArgb(255, 255, 230, 230))
        $bmp.SetPixel(3, 0, [System.Drawing.Color]::Red)
        $argb = [YardstickIconOps]::ToArgb($bmp)
        $result = [YardstickIconOps]::RemoveWhite($argb, 40)

        $result.GetPixel(0, 0).A | Should -Be 0
        $result.GetPixel(2, 0).A | Should -Be 0
        $edge = $result.GetPixel(1, 0)
        $edge.A | Should -BeGreaterThan 0
        $edge.A | Should -BeLessThan 255
        # The white mixed into the edge is taken back out, so it is redder than before.
        $edge.R | Should -Be 255
        $edge.G | Should -BeLessThan 195
        $result.GetPixel(3, 0).ToArgb() | Should -Be ([System.Drawing.Color]::Red.ToArgb())
    }

    It 'softens edge pixels in a flood fill without spreading through them' {
        # white | pinkish edge | white island behind it
        $bmp = New-Object System.Drawing.Bitmap 3, 1
        $bmp.SetPixel(0, 0, [System.Drawing.Color]::White)
        $bmp.SetPixel(1, 0, [System.Drawing.Color]::FromArgb(255, 255, 195, 195))
        $bmp.SetPixel(2, 0, [System.Drawing.Color]::White)
        $result = [YardstickIconOps]::RemoveColor([YardstickIconOps]::ToArgb($bmp), 0, 0, 40, $true)
        $result.GetPixel(0, 0).A | Should -Be 0
        $result.GetPixel(1, 0).A | Should -BeLessThan 255
        $result.GetPixel(1, 0).A | Should -BeGreaterThan 0
        $result.GetPixel(2, 0).A | Should -Be 255
    }

    It 'keeps hard edges when the tolerance is zero' {
        $bmp = New-Object System.Drawing.Bitmap 2, 1
        $bmp.SetPixel(0, 0, [System.Drawing.Color]::White)
        $bmp.SetPixel(1, 0, [System.Drawing.Color]::FromArgb(255, 254, 254, 254))
        $result = [YardstickIconOps]::RemoveWhite([YardstickIconOps]::ToArgb($bmp), 0)
        $result.GetPixel(0, 0).A | Should -Be 0
        $result.GetPixel(1, 0).A | Should -Be 255
    }
}

Describe 'ConvertFrom-ImageBytes' {
    It 'decodes a PNG to 32bpp ARGB, preserving transparency' {
        $path = Join-Path $TestDrive 'clear.png'
        [YardstickIconOps]::RemoveWhite((New-TestBitmap), 10).Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp = ConvertFrom-ImageBytes ([System.IO.File]::ReadAllBytes($path))
        $bmp.Width | Should -Be 40
        $bmp.PixelFormat | Should -Be ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $bmp.GetPixel(0, 0).A | Should -Be 0
        $bmp.GetPixel(15, 10).ToArgb() | Should -Be ([System.Drawing.Color]::Red.ToArgb())
    }

    It 'picks the largest frame of a multi-size ICO' {
        $bytes = New-TestIco -Sizes 16, 48, 32
        $bmp = ConvertFrom-ImageBytes $bytes
        $bmp.Width | Should -Be 48
        $bmp.Height | Should -Be 48
    }

    It 'returns nothing for bytes that are not an image' {
        ConvertFrom-ImageBytes ([System.Text.Encoding]::ASCII.GetBytes('<html>not an image</html>')) | Should -BeNullOrEmpty
        ConvertFrom-ImageBytes $null | Should -BeNullOrEmpty
    }
}

Describe 'Get-ExeIconBitmap' {
    It 'extracts the icon of a program' {
        $bmp = Get-ExeIconBitmap -Path (Get-Process -Id $PID).Path
        $bmp | Should -Not -BeNullOrEmpty
        $bmp.Width | Should -BeGreaterOrEqual 32
        $bmp.PixelFormat | Should -Be ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    }

    It 'returns nothing for a file without icons' {
        $path = Join-Path $TestDrive 'plain.txt'
        Set-Content -LiteralPath $path -Value 'no icons here'
        Get-ExeIconBitmap -Path $path | Should -BeNullOrEmpty
    }
}

Describe 'Get-BingThumbnailUrl' {
    It 'asks Bing for a larger copy of the thumbnail' {
        Get-BingThumbnailUrl -ThumbUrl 'https://tse1.mm.bing.net/th?id=OIP.x&pid=15.1' | Should -Be 'https://tse1.mm.bing.net/th?id=OIP.x&pid=15.1&w=512&h=512&c=7'
        Get-BingThumbnailUrl -ThumbUrl 'https://tse1.mm.bing.net/th/id/OIP.x' -Size 256 | Should -Be 'https://tse1.mm.bing.net/th/id/OIP.x?w=256&h=256&c=7'
    }

    It 'returns an empty string when there is no thumbnail' {
        Get-BingThumbnailUrl -ThumbUrl '' | Should -Be ''
    }
}

Describe 'Recipe lookup' {
    BeforeAll {
        $script:recipeDir = Join-Path $TestDrive 'Recipes'
        $script:devDir = Join-Path $TestDrive 'Development'
        $script:iconDir = Join-Path $TestDrive 'Icons'
        New-Item -ItemType Directory -Path $script:recipeDir, $script:devDir, $script:iconDir | Out-Null
        Set-Content -LiteralPath (Join-Path $script:recipeDir 'brave.yaml') -Value "id: brave`ndisplayName: Brave Browser`niconFile: brave.png`ndescription: |`n  iconFile: nested.png"
        # base: names the recipe file, which need not match the base recipe's id.
        Set-Content -LiteralPath (Join-Path $script:recipeDir '7zip.yaml') -Value "id: 7zip.x64`ndisplayName: `"7-Zip`" # quoted`niconFile: 7zip.png"
        Set-Content -LiteralPath (Join-Path $script:recipeDir '7zip.arm64.yaml') -Value "base: 7zip`nid: 7zip.arm64`ndisplayName: 7-Zip (ARM64)"
        Set-Content -LiteralPath (Join-Path $script:recipeDir 'sketchup2025.yaml') -Value "id: sketchup2025`ndisplayName: SketchUp 2025`niconFile: sketchup.png"
        Set-Content -LiteralPath (Join-Path $script:devDir 'sketchup2026.yaml') -Value "id: sketchup2026`ndisplayName: SketchUp 2026`niconFile: sketchup.png"
        Set-Content -LiteralPath (Join-Path $script:devDir 'noicon.yaml') -Value "id: NoIcon`ndisplayName: No Icon App"
        Set-Content -LiteralPath (Join-Path $script:iconDir '7zip.png') -Value 'png'
        $script:recipes = @(Get-RecipeIconInfo -RecipeFolder $script:recipeDir, $script:devDir, (Join-Path $TestDrive 'missing-folder'))
    }

    It 'reads id, displayName and iconFile from both recipe folders' {
        $script:recipes.Count | Should -Be 6
        $brave = $script:recipes | Where-Object Id -eq 'brave'
        $brave.DisplayName | Should -Be 'Brave Browser'
        $brave.IconFile | Should -Be 'brave.png'
    }

    It 'strips quotes and trailing comments' {
        ($script:recipes | Where-Object Id -eq '7zip.x64').DisplayName | Should -Be '7-Zip'
    }

    It 'inherits iconFile from the base recipe' {
        $arm = $script:recipes | Where-Object Id -eq '7zip.arm64'
        $arm.IconFile | Should -Be '7zip.png'
        $arm.DisplayName | Should -Be '7-Zip (ARM64)'
    }

    It 'suggests a lowercase id-based file name when iconFile is missing' {
        $none = $script:recipes | Where-Object Id -eq 'NoIcon'
        $none.IconFile | Should -BeNullOrEmpty
        $none.FileName | Should -Be 'noicon.png'
    }

    It 'lists each missing icon file once, grouping recipes that share it' {
        $missing = @(Get-MissingIconRecipe -Recipe $script:recipes -IconsFolder $script:iconDir)
        @($missing | ForEach-Object FileName) | Should -Be @('brave.png', 'noicon.png', 'sketchup.png')
        ($missing | Where-Object FileName -eq 'sketchup.png').RecipeIds.Count | Should -Be 2
        ($missing | Where-Object FileName -eq 'noicon.png').HasIconFile | Should -BeFalse
    }

    It 'finds a recipe by id or display name' {
        (Find-Recipe -Recipe $script:recipes -Key 'BRAVE').Id | Should -Be 'brave'
        (Find-Recipe -Recipe $script:recipes -Key 'Brave Browser').Id | Should -Be 'brave'
        (Find-Recipe -Recipe $script:recipes -Key '7zip').Id | Should -Be '7zip.x64'
        Find-Recipe -Recipe $script:recipes -Key 'nope' | Should -BeNullOrEmpty
    }

    It 'drops architecture suffixes from the search text' {
        Get-RecipeSearchText ($script:recipes | Where-Object Id -eq '7zip.arm64') | Should -Be '7-Zip'
    }
}

Describe 'ConvertTo-IconFileName' {
    It 'follows the lowercase recipe id convention' {
        ConvertTo-IconFileName 'Brave Browser' | Should -Be 'bravebrowser.png'
        ConvertTo-IconFileName 'Notepad++' | Should -Be 'notepad.png'
        ConvertTo-IconFileName 'Krita.PNG' | Should -Be 'krita.png'
        ConvertTo-IconFileName '  ' | Should -Be 'icon.png'
    }
}
