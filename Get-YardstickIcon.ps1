<#
.SYNOPSIS
    Small GUI that searches the web for an application icon, lets you clean it up,
    and saves it as a PNG in the Yardstick Icons folder.

.DESCRIPTION
    Searches Bing Images (no API key) for the top 20 candidates and shows them as
    thumbnails. Selecting one loads it into an editor where you can crop (drag a
    rectangle, or Auto-trim), resize, remove a background colour, add a background
    colour, undo, and finally save as a PNG that fits Intune's 512x512 icon limit.

    Open file..., Paste and From EXE... are always available as a manual fallback
    if the search is blocked or returns nothing useful.

    The "Missing icons" list shows every recipe whose iconFile is not in the Icons
    folder. Picking one searches for it and pre-fills the file name the recipe
    expects, so saving is all that is needed.

.PARAMETER RecipeId
    Recipe to fetch an icon for. Searches for the recipe's displayName and saves
    as the recipe's iconFile (following `base:` inheritance).

.PARAMETER Name
    Application name to search for. If it matches a recipe id or displayName that
    recipe is used, otherwise the file name defaults to the lowercase name
    (e.g. "Obsidian" -> obsidian.png).

.PARAMETER OutputFolder
    Where the PNG is saved. Defaults to the IconsPath from Yardstick.Project.psd1,
    falling back to .\Icons beside this script.

.PARAMETER Size
    Maximum width/height of the saved icon. Defaults to 512 (the Intune maximum).

.EXAMPLE
    .\Get-YardstickIcon.ps1 -RecipeId brave
    Searches for "Brave Browser" and saves as Icons\brave.png.

.EXAMPLE
    .\Get-YardstickIcon.ps1 -Name Obsidian
    Opens the picker with "Obsidian" searched for.
#>
[CmdletBinding()]
param(
    [string]$RecipeId,
    [string]$Name,
    [string]$OutputFolder,
    [ValidateRange(16, 1024)]
    [int]$Size = 512,
    # Defines the functions and the image helper, then returns without opening the
    # window. Used by the unit tests.
    [Parameter(DontShow)]
    [switch]$LoadOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 and PowerShell 7 consoles are STA by default; the clipboard
# and the file dialogs do not work without it.
if (-not $LoadOnly -and [Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    throw 'The icon picker needs an STA thread for the clipboard and file dialogs. Run it from a normal PowerShell console, or start PowerShell with -STA.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Net.Http
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { Write-Verbose "Could not enable TLS 1.2: $($_.Exception.Message)" }
# .NET Framework (5.1) allows only 2 connections per host, which serialises the 20
# thumbnail downloads and lets slow ones miss the wait. pwsh ignores this setting.
if ([Net.ServicePointManager]::DefaultConnectionLimit -lt 16) { [Net.ServicePointManager]::DefaultConnectionLimit = 16 }

# WIC (via WPF's imaging classes) decodes WebP and picks the largest ICO frame, which
# GDI+ cannot do. GDI+ stays as the fallback if WPF is unavailable.
$script:WicAvailable = $true
try { Add-Type -AssemblyName PresentationCore, WindowsBase } catch { $script:WicAvailable = $false }

#region Image helper (pixel work)
if (-not ('YardstickIconOps' -as [type])) {
    $helperSource = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

public static class YardstickIconOps
{
    public static Bitmap ToArgb(Image src)
    {
        Bitmap bmp = new Bitmap(src.Width, src.Height, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Transparent);
            g.CompositingMode = CompositingMode.SourceCopy;
            g.DrawImage(src, new Rectangle(0, 0, src.Width, src.Height));
        }
        return bmp;
    }

    // Bgra32 bytes (as produced by WIC) to a 32bpp ARGB bitmap; the memory layouts match.
    public static Bitmap FromBgra32(byte[] bgra, int w, int h)
    {
        Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        BitmapData d = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
        try
        {
            for (int y = 0; y < h; y++)
                Marshal.Copy(bgra, y * w * 4, IntPtr.Add(d.Scan0, y * d.Stride), w * 4);
        }
        finally { bmp.UnlockBits(d); }
        return bmp;
    }

    static int[] GetPixels(Bitmap bmp)
    {
        int w = bmp.Width, h = bmp.Height;
        BitmapData d = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
        try
        {
            int[] px = new int[w * h];
            Marshal.Copy(d.Scan0, px, 0, px.Length);
            return px;
        }
        finally { bmp.UnlockBits(d); }
    }

    static Bitmap FromPixels(int[] px, int w, int h)
    {
        Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        BitmapData d = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
        try { Marshal.Copy(px, 0, d.Scan0, px.Length); }
        finally { bmp.UnlockBits(d); }
        return bmp;
    }

    static bool Close(int p, int r, int g, int b, int tol)
    {
        return Distance(p, r, g, b) <= tol;
    }

    static int Distance(int p, int r, int g, int b)
    {
        int dr = Math.Abs(((p >> 16) & 255) - r);
        int dg = Math.Abs(((p >> 8) & 255) - g);
        int db = Math.Abs((p & 255) - b);
        return Math.Max(dr, Math.Max(dg, db));
    }

    // Background removal for one pixel. Within 'tol' of the colour it becomes fully
    // transparent. Between tol and 2*tol (typically anti-aliased edges) it keeps part
    // of its alpha and has the background colour taken back out, so no halo is left.
    static int Soften(int p, int r, int g, int b, int tol)
    {
        int a = (p >> 24) & 255;
        if (a == 0) return p;
        int d = Distance(p, r, g, b);
        if (d <= tol) return 0;
        if (tol == 0 || d >= 2 * tol) return p;
        double f = (double)(d - tol) / tol;
        int na = (int)Math.Round(a * f);
        if (na == 0) return 0;
        int nr = Unmix((p >> 16) & 255, r, f);
        int ng = Unmix((p >> 8) & 255, g, f);
        int nb = Unmix(p & 255, b, f);
        return (na << 24) | (nr << 16) | (ng << 8) | nb;
    }

    static int Unmix(int c, int bg, double f)
    {
        double v = (c - (1 - f) * bg) / f;
        return (int)Math.Max(0, Math.Min(255, Math.Round(v)));
    }

    public static Bitmap Crop(Bitmap src, Rectangle r)
    {
        int x = Math.Max(0, r.X), y = Math.Max(0, r.Y);
        int w = Math.Min(src.Width - x, r.Width), h = Math.Min(src.Height - y, r.Height);
        if (w < 1 || h < 1) return ToArgb(src);
        return src.Clone(new Rectangle(x, y, w, h), PixelFormat.Format32bppArgb);
    }

    // Bounds of the content once transparent pixels (or pixels close to the opaque
    // top-left colour) are trimmed away. Returns the full image if nothing is found.
    public static Rectangle TrimBounds(Bitmap src, int tolerance)
    {
        int w = src.Width, h = src.Height;
        int[] px = GetPixels(src);
        int corner = px[0];
        bool transparentCorner = ((corner >> 24) & 255) < 16;
        int cr = (corner >> 16) & 255, cg = (corner >> 8) & 255, cb = corner & 255;
        int minX = w, minY = h, maxX = -1, maxY = -1;
        for (int y = 0; y < h; y++)
        {
            for (int x = 0; x < w; x++)
            {
                int p = px[y * w + x];
                bool bg = ((p >> 24) & 255) < 16 || (!transparentCorner && Close(p, cr, cg, cb, tolerance));
                if (bg) continue;
                if (x < minX) minX = x;
                if (x > maxX) maxX = x;
                if (y < minY) minY = y;
                if (y > maxY) maxY = y;
            }
        }
        if (maxX < 0) return new Rectangle(0, 0, w, h);
        return new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
    }

    // Scales so the longest side equals 'size'. When 'square' is set the result is a
    // size x size canvas with the image centred on transparent padding.
    public static Bitmap Resize(Bitmap src, int size, bool square)
    {
        double scale = (double)size / Math.Max(src.Width, src.Height);
        int nw = Math.Max(1, (int)Math.Round(src.Width * scale));
        int nh = Math.Max(1, (int)Math.Round(src.Height * scale));
        int cw = square ? size : nw, ch = square ? size : nh;
        Bitmap bmp = new Bitmap(cw, ch, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp))
        using (ImageAttributes ia = new ImageAttributes())
        {
            g.Clear(Color.Transparent);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            ia.SetWrapMode(WrapMode.TileFlipXY);
            Rectangle dest = new Rectangle((cw - nw) / 2, (ch - nh) / 2, nw, nh);
            g.DrawImage(src, dest, 0, 0, src.Width, src.Height, GraphicsUnit.Pixel, ia);
        }
        return bmp;
    }

    // Makes pixels close to the seed pixel's colour transparent, with soft edges (see
    // Soften). Contiguous mode flood-fills from the seed; edge pixels are softened but
    // the fill does not spread through them. Otherwise every pixel is processed.
    public static Bitmap RemoveColor(Bitmap src, int seedX, int seedY, int tolerance, bool contiguous)
    {
        int w = src.Width, h = src.Height;
        int[] px = GetPixels(src);
        if (seedX < 0 || seedY < 0 || seedX >= w || seedY >= h) return FromPixels(px, w, h);
        int seed = px[seedY * w + seedX];
        if (((seed >> 24) & 255) == 0) return FromPixels(px, w, h);
        int sr = (seed >> 16) & 255, sg = (seed >> 8) & 255, sb = seed & 255;

        if (!contiguous)
        {
            for (int i = 0; i < px.Length; i++) px[i] = Soften(px[i], sr, sg, sb, tolerance);
            return FromPixels(px, w, h);
        }

        bool[] seen = new bool[px.Length];
        Stack<int> stack = new Stack<int>();
        stack.Push(seedY * w + seedX);
        seen[seedY * w + seedX] = true;
        while (stack.Count > 0)
        {
            int i = stack.Pop();
            int x = i % w, y = i / w;
            px[i] = 0;
            if (x > 0) Visit(px, seen, stack, i - 1, sr, sg, sb, tolerance);
            if (x < w - 1) Visit(px, seen, stack, i + 1, sr, sg, sb, tolerance);
            if (y > 0) Visit(px, seen, stack, i - w, sr, sg, sb, tolerance);
            if (y < h - 1) Visit(px, seen, stack, i + w, sr, sg, sb, tolerance);
        }
        return FromPixels(px, w, h);
    }

    // Clears every pixel within 'tolerance' of pure white, with soft edges.
    public static Bitmap RemoveWhite(Bitmap src, int tolerance)
    {
        int[] px = GetPixels(src);
        for (int i = 0; i < px.Length; i++) px[i] = Soften(px[i], 255, 255, 255, tolerance);
        return FromPixels(px, src.Width, src.Height);
    }

    static void Visit(int[] px, bool[] seen, Stack<int> stack, int n, int r, int g, int b, int tol)
    {
        if (seen[n]) return;
        seen[n] = true;
        int p = px[n];
        if (((p >> 24) & 255) == 0 || Close(p, r, g, b, tol)) stack.Push(n);
        else px[n] = Soften(p, r, g, b, tol);
    }

    // Composites the image over a solid colour. With 'rounded' the result is a
    // rounded-corner tile (anti-aliased) instead of a full square.
    public static Bitmap AddBackground(Bitmap src, Color colour, bool rounded)
    {
        int w = src.Width, h = src.Height;
        Bitmap bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (Graphics g = Graphics.FromImage(bmp))
        {
            g.Clear(colour);
            g.DrawImage(src, new Rectangle(0, 0, w, h));
        }
        if (!rounded) return bmp;

        Bitmap mask = new Bitmap(w, h, PixelFormat.Format32bppArgb);
        using (Graphics mg = Graphics.FromImage(mask))
        using (GraphicsPath path = new GraphicsPath())
        using (SolidBrush white = new SolidBrush(Color.White))
        {
            mg.Clear(Color.Transparent);
            mg.SmoothingMode = SmoothingMode.AntiAlias;
            float d = Math.Min(w, h) * 0.4f;
            path.AddArc(0, 0, d, d, 180, 90);
            path.AddArc(w - d - 1, 0, d, d, 270, 90);
            path.AddArc(w - d - 1, h - d - 1, d, d, 0, 90);
            path.AddArc(0, h - d - 1, d, d, 90, 90);
            path.CloseFigure();
            mg.FillPath(white, path);
        }
        int[] px = GetPixels(bmp);
        int[] mp = GetPixels(mask);
        mask.Dispose();
        bmp.Dispose();
        for (int i = 0; i < px.Length; i++)
        {
            int a = ((px[i] >> 24) & 255) * ((mp[i] >> 24) & 255) / 255;
            px[i] = (a << 24) | (px[i] & 0xFFFFFF);
        }
        return FromPixels(px, w, h);
    }
}
'@
    $drawingRefs = @([System.Drawing.Bitmap], [System.Drawing.Color]) | ForEach-Object { $_.Assembly.Location } | Select-Object -Unique
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $drawingRefs += @('System.Runtime', 'System.Collections', 'System.Runtime.InteropServices')
        # Newer .NET builds split Graphics' interfaces into System.Private.Windows.* assemblies the compiler also needs.
        $drawingRefs += @(Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName($drawingRefs[0])) -Filter 'System.Private.Windows*.dll' -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
    }
    Add-Type -TypeDefinition $helperSource -ReferencedAssemblies $drawingRefs
}

function ConvertFrom-ImageBytes {
    # Decodes image bytes to a 32bpp ARGB bitmap, or $null if they are not an image.
    [CmdletBinding()]
    param([byte[]]$Bytes)

    if (-not $Bytes -or $Bytes.Length -lt 16) { return $null }
    if ($script:WicAvailable) {
        $stream = New-Object System.IO.MemoryStream (, $Bytes)
        try {
            $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create($stream,
                [System.Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat,
                [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
            # ICO files hold several sizes; take the largest (then the deepest colour).
            $best = $null
            foreach ($frame in $decoder.Frames) {
                if (-not $best) { $best = $frame; continue }
                $area = $frame.PixelWidth * $frame.PixelHeight
                $bestArea = $best.PixelWidth * $best.PixelHeight
                if ($area -gt $bestArea -or ($area -eq $bestArea -and $frame.Format.BitsPerPixel -gt $best.Format.BitsPerPixel)) { $best = $frame }
            }
            if ($best) {
                $converted = New-Object System.Windows.Media.Imaging.FormatConvertedBitmap ($best, [System.Windows.Media.PixelFormats]::Bgra32, $null, 0)
                $w = $converted.PixelWidth
                $h = $converted.PixelHeight
                $buffer = New-Object byte[] ($w * $h * 4)
                $converted.CopyPixels($buffer, $w * 4, 0)
                return [YardstickIconOps]::FromBgra32($buffer, $w, $h)
            }
        } catch {
            Write-Verbose "WIC could not decode the image, trying GDI+: $($_.Exception.Message)"
        } finally {
            $stream.Dispose()
        }
    }

    $stream = New-Object System.IO.MemoryStream (, $Bytes)
    try {
        $img = [System.Drawing.Image]::FromStream($stream)
        try { return [YardstickIconOps]::ToArgb($img) } finally { $img.Dispose() }
    } catch {
        return $null
    } finally {
        $stream.Dispose()
    }
}

function Get-ExeIconBitmap {
    # Largest embedded icon of an .exe/.dll as a 32bpp ARGB bitmap, or $null.
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not ('Win32IconExtractor' -as [type])) {
        # Same signature as Tests\RecipeHarness\Export-ExeIcon.ps1, so either can load it first.
        $signature = @'
[DllImport("user32.dll", CharSet = CharSet.Unicode)]
public static extern int PrivateExtractIcons(string lpszFile, int nIconIndex, int cxIcon, int cyIcon, IntPtr[] phicon, int[] piconid, int nIcons, int flags);
[DllImport("user32.dll")]
public static extern bool DestroyIcon(IntPtr hIcon);
'@
        Add-Type -MemberDefinition $signature -Name 'Win32IconExtractor' -Namespace '' | Out-Null
    }

    foreach ($px in @(256, 128, 64, 48, 32)) {
        $handles = New-Object IntPtr[] 1
        $ids = New-Object int[] 1
        $count = [Win32IconExtractor]::PrivateExtractIcons($Path, 0, $px, $px, $handles, $ids, 1, 0)
        if ($count -le 0 -or $handles[0] -eq [IntPtr]::Zero) { continue }
        try {
            $icon = [System.Drawing.Icon]::FromHandle($handles[0])
            try {
                $raw = $icon.ToBitmap()
                try { return [YardstickIconOps]::ToArgb($raw) } finally { $raw.Dispose() }
            } finally { $icon.Dispose() }
        } finally {
            [Win32IconExtractor]::DestroyIcon($handles[0]) | Out-Null
        }
    }
    $null
}
#endregion

#region Search
$script:UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'

function ConvertFrom-BingImageHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyString()][string]$Html,
        [int]$Max = 20
    )

    $results = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    $tags = [regex]::Matches($Html, '<a\b[^>]*\bclass="[^"]*\biusc\b[^"]*"[^>]*>')
    for ($i = 0; $i -lt $tags.Count -and $results.Count -lt $Max; $i++) {
        $tag = $tags[$i]
        $attr = [regex]::Match($tag.Value, '\sm="([^"]*)"')
        if (-not $attr.Success) { continue }
        try {
            $meta = [System.Net.WebUtility]::HtmlDecode($attr.Groups[1].Value) | ConvertFrom-Json
        } catch { continue }
        if (-not $meta.PSObject.Properties['murl'] -or -not $meta.murl) { continue }
        if (-not $seen.Add([string]$meta.murl)) { continue }

        $windowStart = $tag.Index + $tag.Length
        $windowEnd = if ($i + 1 -lt $tags.Count) { $tags[$i + 1].Index } else { $Html.Length }
        $window = $Html.Substring($windowStart, [Math]::Min($windowEnd - $windowStart, 1500))
        $dims = [regex]::Match($window, '(\d{2,5})\s*(?:x|&#215;|&times;|[^0-9a-z<>\s])\s*(\d{2,5})')

        $title = ''
        if ($meta.PSObject.Properties['t'] -and $meta.t) {
            $title = [System.Net.WebUtility]::HtmlDecode(([string]$meta.t -replace '</?[^>]+>', ''))
        }
        $results.Add([pscustomobject]@{
            Title    = $title
            FullUrl  = [string]$meta.murl
            ThumbUrl = if ($meta.PSObject.Properties['turl']) { [string]$meta.turl } else { '' }
            Width    = if ($dims.Success) { [int]$dims.Groups[1].Value } else { 0 }
            Height   = if ($dims.Success) { [int]$dims.Groups[2].Value } else { 0 }
        })
    }
    $results.ToArray()
}

function Get-BingImageSearchUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Query,
        [switch]$TransparentOnly
    )
    $filters = '+filterui:aspect-square'
    if ($TransparentOnly) { $filters += '+filterui:photo-transparent' }
    'https://www.bing.com/images/search?q={0}&qft={1}&first=1' -f [Uri]::EscapeDataString($Query), $filters
}

function Get-BingThumbnailUrl {
    # Bing serves its thumbnails at any size; used when the full image will not load.
    [CmdletBinding()]
    param(
        [AllowEmptyString()][string]$ThumbUrl,
        [int]$Size = 512
    )
    if (-not $ThumbUrl) { return '' }
    $separator = if ($ThumbUrl.Contains('?')) { '&' } else { '?' }
    '{0}{1}w={2}&h={2}&c=7' -f $ThumbUrl, $separator, $Size
}
#endregion

#region Recipes
function ConvertTo-IconFileName {
    # Icon files follow the recipe id convention: lowercase, no spaces (brave.png).
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Text)
    $stem = ($Text.Trim().ToLowerInvariant() -replace '\.png$', '') -replace '[^a-z0-9._-]', ''
    if (-not $stem) { $stem = 'icon' }
    "$stem.png"
}

function Get-RecipeIconInfo {
    # id / displayName / iconFile for every recipe, with `base:` inheritance applied
    # (base names a recipe file, as in Merge-RecipeWithBase). Only these top-level keys
    # are needed, so a line regex avoids a YAML module dependency.
    [CmdletBinding()]
    param([string[]]$RecipeFolder)

    $raw = [ordered]@{}
    foreach ($folder in $RecipeFolder) {
        if (-not $folder -or -not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        foreach ($file in Get-ChildItem -LiteralPath $folder -File | Where-Object { $_.Extension -in '.yaml', '.yml' }) {
            if ($raw.Contains($file.BaseName)) { continue }
            $fields = @{ id = ''; displayName = ''; iconFile = ''; base = ''; path = $file.FullName }
            $text = [System.IO.File]::ReadAllText($file.FullName)
            foreach ($m in [regex]::Matches($text, '(?m)^(id|displayName|iconFile|base):[ \t]*(.*?)[ \t]*\r?$')) {
                $value = $m.Groups[2].Value
                if ($value -match '^"([^"]*)"' -or $value -match "^'([^']*)'") { $value = $Matches[1] }
                else { $value = ($value -replace '\s+#.*$', '').Trim() }
                if (-not $fields[$m.Groups[1].Value]) { $fields[$m.Groups[1].Value] = $value }
            }
            $raw[$file.BaseName] = $fields
        }
    }

    $out = foreach ($key in $raw.Keys) {
        $recipe = $raw[$key]
        $id = if ($recipe.id) { $recipe.id } else { $key }
        $displayName = $recipe.displayName
        $iconFile = $recipe.iconFile
        $parent = $recipe.base
        $depth = 0
        while ($parent -and $raw.Contains($parent) -and $depth -lt 10 -and (-not $displayName -or -not $iconFile)) {
            if (-not $displayName) { $displayName = $raw[$parent].displayName }
            if (-not $iconFile) { $iconFile = $raw[$parent].iconFile }
            $parent = $raw[$parent].base
            $depth++
        }
        [pscustomobject]@{
            Id          = $id
            DisplayName = if ($displayName) { $displayName } else { $id }
            IconFile    = $iconFile
            FileName    = if ($iconFile) { $iconFile } else { ConvertTo-IconFileName $id }
            Path        = $recipe.path
        }
    }
    $out | Sort-Object Id
}

function Get-MissingIconRecipe {
    # One entry per icon file that is missing; recipes that share an icon are grouped.
    [CmdletBinding()]
    param(
        [object[]]$Recipe,
        [Parameter(Mandatory)][string]$IconsFolder
    )
    $missing = @($Recipe | Where-Object { $_ -and -not $_.IconFile -or ($_.IconFile -and -not (Test-Path -LiteralPath (Join-Path $IconsFolder $_.IconFile) -PathType Leaf)) })
    foreach ($group in ($missing | Group-Object { $_.FileName.ToLowerInvariant() } | Sort-Object { $_.Group[0].DisplayName })) {
        $first = $group.Group[0]
        [pscustomobject]@{
            FileName    = $first.FileName
            DisplayName = $first.DisplayName
            Id          = $first.Id
            RecipeIds   = @($group.Group | ForEach-Object { $_.Id })
            HasIconFile = [bool]$first.IconFile
        }
    }
}

function Find-Recipe {
    # Matches a recipe by id, file name or displayName (case-insensitive).
    [CmdletBinding()]
    param([object[]]$Recipe, [string]$Key)
    if (-not $Key) { return $null }
    $hit = @($Recipe | Where-Object { $_.Id -eq $Key -or [IO.Path]::GetFileNameWithoutExtension($_.Path) -eq $Key })
    if ($hit.Count -eq 0) { $hit = @($Recipe | Where-Object { $_.DisplayName -eq $Key }) }
    if ($hit.Count -gt 0) { $hit[0] } else { $null }
}

function Get-RecipeSearchText {
    # "7-Zip (ARM64)" -> "7-Zip": the architecture suffix only hurts an image search.
    param($Recipe)
    (($Recipe.DisplayName -replace '\s*\([^)]*\)', '') -replace '\s{2,}', ' ').Trim()
}
#endregion

if ($LoadOnly) { return }

#region Project layout and recipes
$layout = $null
try {
    . (Join-Path $PSScriptRoot 'ProjectLayout.ps1')
    $layout = Get-YardstickProjectLayout -ProjectRoot $PSScriptRoot
} catch {
    Write-Verbose "Project layout unavailable: $($_.Exception.Message)"
}
if (-not $OutputFolder) {
    $OutputFolder = if ($layout -and $layout.PSObject.Properties['IconsPath']) { $layout.IconsPath } else { Join-Path $PSScriptRoot 'Icons' }
}

$recipeFolders = @()
if ($layout) {
    foreach ($key in @('RecipesPath', 'DevelopmentRecipes')) {
        if ($layout.PSObject.Properties[$key]) { $recipeFolders += $layout.$key }
    }
} else {
    $recipeFolders = @(Join-Path $PSScriptRoot 'Recipes')
}
$script:Recipes = @(Get-RecipeIconInfo -RecipeFolder $recipeFolders)

$startRecipe = $null
if ($RecipeId) {
    $startRecipe = Find-Recipe -Recipe $script:Recipes -Key $RecipeId
    if (-not $startRecipe) { throw "Recipe '$RecipeId' was not found in: $($recipeFolders -join ', ')" }
} elseif ($Name) {
    $startRecipe = Find-Recipe -Recipe $script:Recipes -Key $Name
}
#endregion

#region HTTP (UI-thread friendly)
$script:Http = New-Object System.Net.Http.HttpClient
$script:Http.Timeout = [TimeSpan]::FromSeconds(20)
$script:Http.DefaultRequestHeaders.TryAddWithoutValidation('User-Agent', $script:UserAgent) | Out-Null
$script:Http.DefaultRequestHeaders.TryAddWithoutValidation('Accept-Language', 'en-US,en;q=0.9') | Out-Null

function Wait-Tasks {
    # Keeps the window responsive while downloads run; gives up early if it closes.
    param([object[]]$Tasks, [int]$TimeoutMs)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.ElapsedMilliseconds -lt $TimeoutMs -and -not $script:State.Closing) {
        $pending = @($Tasks | Where-Object { $_ -and -not $_.IsCompleted })
        if ($pending.Count -eq 0) { break }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 15
    }
}

function Get-TaskBytes {
    param($Task)
    if ($Task -and $Task.IsCompleted -and $Task.Status -eq [System.Threading.Tasks.TaskStatus]::RanToCompletion) { return , $Task.Result }
    $null
}

function Get-UrlBytes {
    param([string]$Url, [int]$TimeoutMs)
    if (-not $Url) { return $null }
    try { $task = $script:Http.GetByteArrayAsync($Url) } catch { return $null }
    Wait-Tasks @($task) $TimeoutMs
    Get-TaskBytes $task
}
#endregion

#region State
$script:State = @{
    Current     = $null
    Original    = $null
    Undo        = New-Object System.Collections.Generic.List[System.Drawing.Bitmap]
    Selection   = [System.Drawing.Rectangle]::Empty
    DragStart   = $null
    Picking     = $false
    BgColour    = [System.Drawing.Color]::White
    Busy        = $false
    Closing     = $false
    # Result clicked while another download was running; loaded when it finishes.
    Pending     = $null
    # Bumped by every load so a slow download cannot replace a newer image.
    LoadGen     = 0
    NameTouched = $false
    Recipe      = $startRecipe
    Missing     = @()
    Rebuilding  = $false
}
$script:MaxUndo = 20
# Larger images are scaled down on load; editing a 4000px photo is slow and pointless for a 512px icon.
$script:MaxLoadSize = 1024
#endregion

#region UI construction
# Become DPI-aware before any window exists. A DPI-unaware process (Windows PowerShell)
# is bitmap-stretched by Windows until the first WIC decode loads WPF, which makes the
# process DPI-aware mid-run and shrinks the whole window. Sizes below are written for
# 96 DPI and auto-scaled to the real DPI when layout resumes.
if (-not ('YardstickIconDpi' -as [type])) {
    Add-Type -MemberDefinition '[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();' -Name 'YardstickIconDpi' -Namespace '' | Out-Null
}
[void][YardstickIconDpi]::SetProcessDPIAware()
[System.Windows.Forms.Application]::EnableVisualStyles()

$form = New-Object System.Windows.Forms.Form
$form.SuspendLayout()
$form.Text = 'Yardstick Icon Picker'
$form.StartPosition = 'CenterScreen'
$form.ClientSize = New-Object System.Drawing.Size(1240, 780)
$form.MinimumSize = New-Object System.Drawing.Size(1100, 600)

$grid = New-Object System.Windows.Forms.TableLayoutPanel
$grid.Dock = 'Fill'
$grid.ColumnCount = 3
$grid.RowCount = 3
[void]$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 340)))
[void]$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))
[void]$grid.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Absolute', 250)))
[void]$grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
[void]$grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100)))
[void]$grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize')))
$form.Controls.Add($grid)

# Top bar. Bar control sizes, margins and padding are set once the DPI is known (see Set-BarLayout).
$top = New-Object System.Windows.Forms.FlowLayoutPanel
$top.Dock = 'Fill'
$top.WrapContents = $false
$top.AutoSize = $true
$top.AutoSizeMode = 'GrowAndShrink'
$grid.Controls.Add($top, 0, 0)
$grid.SetColumnSpan($top, 3)

$cmbMissing = New-Object System.Windows.Forms.ComboBox
$cmbMissing.DropDownStyle = 'DropDownList'
$cmbMissing.Width = 200
$cmbMissing.DropDownWidth = 380
$cmbMissing.MaxDropDownItems = 25
$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Width = 240
$txtSearch.Text = if ($Name) { $Name } elseif ($startRecipe) { Get-RecipeSearchText $startRecipe } else { '' }
$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = 'Search'
$chkAppend = New-Object System.Windows.Forms.CheckBox
$chkAppend.Text = "Append 'app icon'"
$chkAppend.Checked = $true
$chkAppend.AutoSize = $true
$chkTransparent = New-Object System.Windows.Forms.CheckBox
$chkTransparent.Text = 'Transparent only'
$chkTransparent.AutoSize = $true
$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = 'Open file...'
$btnPaste = New-Object System.Windows.Forms.Button
$btnPaste.Text = 'Paste'
$btnExe = New-Object System.Windows.Forms.Button
$btnExe.Text = 'From EXE...'
foreach ($c in @($cmbMissing, $txtSearch, $btnSearch, $chkAppend, $chkTransparent, $btnOpen, $btnPaste, $btnExe)) { $top.Controls.Add($c) }

# Results list
$imageList = New-Object System.Windows.Forms.ImageList
$imageList.ImageSize = New-Object System.Drawing.Size(96, 96)
$imageList.ColorDepth = 'Depth32Bit'
$lv = New-Object System.Windows.Forms.ListView
$lv.Dock = 'Fill'
$lv.View = 'LargeIcon'
$lv.LargeImageList = $imageList
$lv.MultiSelect = $false
$lv.HideSelection = $false
$grid.Controls.Add($lv, 0, 1)

# Canvas
$canvas = New-Object System.Windows.Forms.Panel
$canvas.Dock = 'Fill'
$canvas.BackColor = [System.Drawing.Color]::FromArgb(64, 64, 64)
$canvas.GetType().GetProperty('DoubleBuffered', [Reflection.BindingFlags]'NonPublic,Instance').SetValue($canvas, $true, $null)
$grid.Controls.Add($canvas, 1, 1)

# Tool panel
$tools = New-Object System.Windows.Forms.FlowLayoutPanel
$tools.Dock = 'Fill'
$tools.FlowDirection = 'TopDown'
$tools.WrapContents = $false
$tools.AutoScroll = $true
$tools.Padding = New-Object System.Windows.Forms.Padding(4)
$grid.Controls.Add($tools, 2, 1)

function New-ToolGroup {
    param([string]$Title)
    $g = New-Object System.Windows.Forms.GroupBox
    $g.Text = $Title
    $g.Width = 226
    $g.Height = 10
    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock = 'Fill'
    $flow.FlowDirection = 'TopDown'
    $flow.WrapContents = $false
    $flow.AutoSize = $false
    $g.Controls.Add($flow)
    $tools.Controls.Add($g)
    [pscustomobject]@{ Group = $g; Flow = $flow }
}

function Set-GroupHeight {
    param($Tool, [int]$Height)
    $Tool.Group.Height = $Height
}

function New-ToolButton {
    param([string]$Text, [bool]$Toggle = $false)
    $b = if ($Toggle) { New-Object System.Windows.Forms.CheckBox } else { New-Object System.Windows.Forms.Button }
    if ($Toggle) { $b.Appearance = 'Button' }
    $b.Text = $Text
    $b.Width = 206
    $b.Height = 26
    if ($Toggle) { $b.TextAlign = 'MiddleCenter' }
    $b
}

$gCrop = New-ToolGroup 'Crop'
$lblCrop = New-Object System.Windows.Forms.Label
$lblCrop.Text = 'Drag on the image to select an area.'
$lblCrop.Width = 206
$lblCrop.Height = 30
$btnCrop = New-ToolButton 'Crop to selection'
$btnTrim = New-ToolButton 'Auto-trim borders'
foreach ($c in @($lblCrop, $btnCrop, $btnTrim)) { $gCrop.Flow.Controls.Add($c) }
Set-GroupHeight $gCrop 115

$gSize = New-ToolGroup 'Resize'
$cmbSize = New-Object System.Windows.Forms.ComboBox
$cmbSize.Width = 206
foreach ($s in @(512, 256, 128, 64)) { [void]$cmbSize.Items.Add([string]$s) }
$cmbSize.Text = [string]$Size
$chkSquare = New-Object System.Windows.Forms.CheckBox
$chkSquare.Text = 'Pad to square canvas'
$chkSquare.Checked = $true
$chkSquare.Width = 206
$btnResize = New-ToolButton 'Resize (longest side)'
foreach ($c in @($cmbSize, $chkSquare, $btnResize)) { $gSize.Flow.Controls.Add($c) }
Set-GroupHeight $gSize 112

$gRemove = New-ToolGroup 'Remove background'
$btnPick = New-ToolButton 'Pick colour to remove' $true
$lblTol = New-Object System.Windows.Forms.Label
$lblTol.Text = 'Tolerance: 30'
$lblTol.Width = 206
$trkTol = New-Object System.Windows.Forms.TrackBar
$trkTol.Minimum = 0
$trkTol.Maximum = 150
$trkTol.Value = 30
$trkTol.TickFrequency = 10
$trkTol.Width = 206
$chkAll = New-Object System.Windows.Forms.CheckBox
$chkAll.Text = 'All matching (not just touching)'
$chkAll.Width = 206
$btnWhite = New-ToolButton 'Remove white'
foreach ($c in @($btnPick, $lblTol, $trkTol, $chkAll, $btnWhite)) { $gRemove.Flow.Controls.Add($c) }
Set-GroupHeight $gRemove 212

$gAdd = New-ToolGroup 'Add background'
$btnColour = New-ToolButton 'Colour...'
$btnColour.BackColor = $script:State.BgColour
$chkRounded = New-Object System.Windows.Forms.CheckBox
$chkRounded.Text = 'Rounded tile'
$chkRounded.Width = 206
$btnApplyBg = New-ToolButton 'Apply background'
foreach ($c in @($btnColour, $chkRounded, $btnApplyBg)) { $gAdd.Flow.Controls.Add($c) }
Set-GroupHeight $gAdd 115

$gHist = New-ToolGroup 'History'
$btnUndo = New-ToolButton 'Undo'
$btnReset = New-ToolButton 'Reset to original'
foreach ($c in @($btnUndo, $btnReset)) { $gHist.Flow.Controls.Add($c) }
Set-GroupHeight $gHist 92

# Bottom bar
$bottom = New-Object System.Windows.Forms.FlowLayoutPanel
$bottom.Dock = 'Fill'
$bottom.WrapContents = $false
$bottom.AutoSize = $true
$bottom.AutoSizeMode = 'GrowAndShrink'
$grid.Controls.Add($bottom, 0, 2)
$grid.SetColumnSpan($bottom, 3)
$lblFile = New-Object System.Windows.Forms.Label
$lblFile.Text = 'Save as:'
$lblFile.AutoSize = $true
$txtFile = New-Object System.Windows.Forms.TextBox
$txtFile.Width = 260
if ($startRecipe) { $txtFile.Text = $startRecipe.FileName }
elseif ($Name) { $txtFile.Text = ConvertTo-IconFileName $Name }
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = 'Save PNG'
$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.AutoSize = $true
$lblStatus.Text = "Output folder: $OutputFolder"
foreach ($c in @($lblFile, $txtFile, $btnSave, $lblStatus)) { $bottom.Controls.Add($c) }

$toolTip = New-Object System.Windows.Forms.ToolTip

$form.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$form.AutoScaleMode = 'Dpi'
$form.ResumeLayout($false)
$form.PerformLayout()
$script:UiScale = $form.CurrentAutoScaleDimensions.Width / 96

function Set-BarLayout {
    # Lays out a top/bottom bar after auto-scaling, in device pixels. AutoSize buttons keep
    # their (scaled) 75x23 default as a minimum, which made them taller than the text boxes
    # and pushed them out of the bar, so bar buttons get one explicit size instead, and every
    # control is centred on the button height. LeftGap (96-DPI pixels) separates groups.
    param(
        [System.Windows.Forms.FlowLayoutPanel]$Bar,
        [System.Drawing.Size]$ButtonSize,
        [hashtable]$LeftGap = @{}
    )
    $px = { param([double]$Value) [int][Math]::Round($Value * $script:UiScale) }
    $Bar.Padding = New-Object System.Windows.Forms.Padding((& $px 6), (& $px 5), (& $px 6), (& $px 5))
    foreach ($c in $Bar.Controls) {
        if ($c -is [System.Windows.Forms.Button]) { $c.AutoSize = $false; $c.Size = $ButtonSize }
        $h = if ($c.AutoSize) { $c.PreferredSize.Height } else { $c.Height }
        $topGap = [Math]::Max(0, [int][Math]::Floor(($ButtonSize.Height - $h) / 2))
        $left = if ($LeftGap.ContainsKey($c.Name)) { & $px $LeftGap[$c.Name] } else { & $px 3 }
        $c.Margin = New-Object System.Windows.Forms.Padding($left, $topGap, (& $px 3), ($ButtonSize.Height - $h - $topGap))
    }
}

# Every bar button shares one size: the widest caption plus padding, a little taller than a text box.
$barButtons = @($btnSearch, $btnOpen, $btnPaste, $btnExe, $btnSave)
$captionWidth = ($barButtons | ForEach-Object { [System.Windows.Forms.TextRenderer]::MeasureText($_.Text, $_.Font).Width } | Measure-Object -Maximum).Maximum
$barButtonSize = New-Object System.Drawing.Size(
    [int][Math]::Max($captionWidth + 24 * $script:UiScale, 88 * $script:UiScale),
    [int]([Math]::Max($txtSearch.Height, $cmbMissing.Height) + 4 * $script:UiScale))
$txtSearch.Name = 'txtSearch'; $chkAppend.Name = 'chkAppend'; $btnOpen.Name = 'btnOpen'; $lblStatus.Name = 'lblStatus'
Set-BarLayout -Bar $top -ButtonSize $barButtonSize -LeftGap @{ txtSearch = 12; chkAppend = 12; btnOpen = 20 }
Set-BarLayout -Bar $bottom -ButtonSize $barButtonSize -LeftGap @{ lblStatus = 15 }

# The ImageList is not a control, so auto-scaling skips it. ImageList tops out at 256.
$tilePx = [int][Math]::Min(256, [Math]::Round(96 * $script:UiScale))
$imageList.ImageSize = New-Object System.Drawing.Size($tilePx, $tilePx)
#endregion

#region Editor logic
function Set-Status {
    param([string]$Text)
    $lblStatus.Text = $Text
    if (-not $script:State.Closing) { [System.Windows.Forms.Application]::DoEvents() }
}

function Get-CanvasLayout {
    $bmp = $script:State.Current
    if (-not $bmp) { return $null }
    $cw = [Math]::Max(1, $canvas.ClientSize.Width - 20)
    $ch = [Math]::Max(1, $canvas.ClientSize.Height - 20)
    $scale = [Math]::Min($cw / $bmp.Width, $ch / $bmp.Height)
    $w = [int][Math]::Max(1, [Math]::Round($bmp.Width * $scale))
    $h = [int][Math]::Max(1, [Math]::Round($bmp.Height * $scale))
    [pscustomobject]@{
        Scale = $scale
        Rect  = New-Object System.Drawing.Rectangle ([int](($canvas.ClientSize.Width - $w) / 2)), ([int](($canvas.ClientSize.Height - $h) / 2)), $w, $h
    }
}

function ConvertTo-ImagePoint {
    param($Layout, [int]$X, [int]$Y)
    $bmp = $script:State.Current
    $ix = [int][Math]::Floor(($X - $Layout.Rect.X) / $Layout.Scale)
    $iy = [int][Math]::Floor(($Y - $Layout.Rect.Y) / $Layout.Scale)
    New-Object System.Drawing.Point ([Math]::Max(0, [Math]::Min($bmp.Width - 1, $ix))), ([Math]::Max(0, [Math]::Min($bmp.Height - 1, $iy)))
}

function Set-CurrentImage {
    param([System.Drawing.Bitmap]$Bitmap, [bool]$KeepUndo = $true)
    if ($KeepUndo -and $script:State.Current) {
        $script:State.Undo.Add($script:State.Current)
        if ($script:State.Undo.Count -gt $script:MaxUndo) {
            $script:State.Undo[0].Dispose()
            $script:State.Undo.RemoveAt(0)
        }
    } elseif ($script:State.Current) {
        $script:State.Current.Dispose()
    }
    $script:State.Current = $Bitmap
    $script:State.Selection = [System.Drawing.Rectangle]::Empty
    $canvas.Invalidate()
    Update-State
}

function Initialize-Image {
    param([System.Drawing.Bitmap]$Bitmap, [string]$Source)
    # Whatever loaded last wins: drop queued results and invalidate running downloads.
    $script:State.LoadGen++
    $script:State.Pending = $null

    $note = ''
    if ($Bitmap.Width -gt $script:MaxLoadSize -or $Bitmap.Height -gt $script:MaxLoadSize) {
        $note = ", scaled down from {0}x{1}" -f $Bitmap.Width, $Bitmap.Height
        $smaller = [YardstickIconOps]::Resize($Bitmap, $script:MaxLoadSize, $false)
        $Bitmap.Dispose()
        $Bitmap = $smaller
    }
    foreach ($b in $script:State.Undo) { $b.Dispose() }
    $script:State.Undo.Clear()
    if ($script:State.Original) { $script:State.Original.Dispose() }
    $script:State.Original = $Bitmap.Clone()
    if ($script:State.Current) { $script:State.Current.Dispose(); $script:State.Current = $null }
    Set-CurrentImage -Bitmap $Bitmap -KeepUndo $false
    Set-Status ("Loaded {0} ({1}x{2}{3})" -f $Source, $Bitmap.Width, $Bitmap.Height, $note)
}

function Update-State {
    $has = [bool]$script:State.Current
    foreach ($c in @($btnCrop, $btnTrim, $btnResize, $btnPick, $btnWhite, $btnApplyBg, $btnSave)) { $c.Enabled = $has }
    $btnUndo.Enabled = $script:State.Undo.Count -gt 0
    $btnReset.Enabled = $has -and [bool]$script:State.Original
    $btnCrop.Enabled = $has -and -not $script:State.Selection.IsEmpty
}

function Update-MissingList {
    $script:State.Missing = @(Get-MissingIconRecipe -Recipe $script:Recipes -IconsFolder $OutputFolder)
    $script:State.Rebuilding = $true
    try {
        $cmbMissing.BeginUpdate()
        $cmbMissing.Items.Clear()
        $header = if ($script:State.Missing.Count) { "Missing icons ($($script:State.Missing.Count))..." } else { 'No missing icons' }
        [void]$cmbMissing.Items.Add($header)
        foreach ($m in $script:State.Missing) {
            $extra = if ($m.RecipeIds.Count -gt 1) { " +$($m.RecipeIds.Count - 1) more" } else { '' }
            $file = if ($m.HasIconFile) { $m.FileName } else { "$($m.FileName), no iconFile set" }
            [void]$cmbMissing.Items.Add(("{0}  ({1}){2}" -f $m.DisplayName, $file, $extra))
        }
        $cmbMissing.SelectedIndex = 0
        $cmbMissing.EndUpdate()
    } finally {
        $script:State.Rebuilding = $false
    }
}

function Select-Recipe {
    param($Recipe)
    $script:State.Recipe = $Recipe
    $script:State.NameTouched = $false
    $txtSearch.Text = Get-RecipeSearchText $Recipe
    $txtFile.Text = $Recipe.FileName
    Invoke-Search
}

$checker = New-Object System.Drawing.Bitmap 16, 16
$cg = [System.Drawing.Graphics]::FromImage($checker)
$cg.Clear([System.Drawing.Color]::FromArgb(235, 235, 235))
$greyBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(200, 200, 200))
$cg.FillRectangle($greyBrush, 0, 0, 8, 8)
$cg.FillRectangle($greyBrush, 8, 8, 8, 8)
$cg.Dispose()
$greyBrush.Dispose()
$checkerBrush = New-Object System.Drawing.TextureBrush $checker

$canvas.Add_Paint({
    $e = $_
    $lay = Get-CanvasLayout
    if (-not $lay) {
        $f = New-Object System.Drawing.Font('Segoe UI', 11)
        $e.Graphics.DrawString('Search for an icon, or use Open file... / Paste / From EXE...', $f, [System.Drawing.Brushes]::Gainsboro, 20, 20)
        $f.Dispose()
        return
    }
    $g = $e.Graphics
    $g.FillRectangle($checkerBrush, $lay.Rect)
    $g.InterpolationMode = if ($lay.Scale -gt 1) { 'NearestNeighbor' } else { 'HighQualityBicubic' }
    $g.PixelOffsetMode = 'Half'
    $g.DrawImage($script:State.Current, $lay.Rect)
    $sel = $script:State.Selection
    if (-not $sel.IsEmpty) {
        $r = New-Object System.Drawing.Rectangle ([int]($lay.Rect.X + $sel.X * $lay.Scale)), ([int]($lay.Rect.Y + $sel.Y * $lay.Scale)), ([int]($sel.Width * $lay.Scale)), ([int]($sel.Height * $lay.Scale))
        $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Black), 1
        $pen2 = New-Object System.Drawing.Pen ([System.Drawing.Color]::White), 1
        $pen2.DashStyle = 'Dash'
        $g.DrawRectangle($pen, $r)
        $g.DrawRectangle($pen2, $r)
        $pen.Dispose(); $pen2.Dispose()
    }
})
$canvas.Add_Resize({ $canvas.Invalidate() })

$canvas.Add_MouseDown({
    $e = $_
    if (-not $script:State.Current -or $e.Button -ne 'Left') { return }
    $lay = Get-CanvasLayout
    $pt = ConvertTo-ImagePoint $lay $e.X $e.Y
    if ($script:State.Picking) {
        $result = [YardstickIconOps]::RemoveColor($script:State.Current, $pt.X, $pt.Y, $trkTol.Value, (-not $chkAll.Checked))
        Set-CurrentImage -Bitmap $result
        Set-Status ("Removed colour at ({0},{1}) with tolerance {2}" -f $pt.X, $pt.Y, $trkTol.Value)
        return
    }
    $script:State.DragStart = $pt
    $script:State.Selection = [System.Drawing.Rectangle]::Empty
    $canvas.Invalidate()
    Update-State
})
$canvas.Add_MouseMove({
    $e = $_
    if (-not $script:State.DragStart -or -not $script:State.Current) { return }
    $lay = Get-CanvasLayout
    $pt = ConvertTo-ImagePoint $lay $e.X $e.Y
    $start = $script:State.DragStart
    $x = [Math]::Min($start.X, $pt.X); $y = [Math]::Min($start.Y, $pt.Y)
    $script:State.Selection = New-Object System.Drawing.Rectangle $x, $y, ([Math]::Abs($pt.X - $start.X) + 1), ([Math]::Abs($pt.Y - $start.Y) + 1)
    $canvas.Invalidate()
})
$canvas.Add_MouseUp({
    if (-not $script:State.DragStart) { return }
    $script:State.DragStart = $null
    $sel = $script:State.Selection
    if ($sel.Width -lt 3 -or $sel.Height -lt 3) { $script:State.Selection = [System.Drawing.Rectangle]::Empty }
    $canvas.Invalidate()
    Update-State
    if (-not $script:State.Selection.IsEmpty) { Set-Status ("Selection {0}x{1} at ({2},{3})" -f $sel.Width, $sel.Height, $sel.X, $sel.Y) }
})

function Invoke-Search {
    if ($script:State.Busy) { Set-Status 'Still downloading; search again in a moment.'; return }
    $text = $txtSearch.Text.Trim()
    if (-not $text) { Set-Status 'Type something to search for.'; return }
    $query = if ($chkAppend.Checked) { "$text app icon" } else { $text }
    $script:State.Busy = $true
    $form.UseWaitCursor = $true
    try {
        Set-Status "Searching for '$query'..."
        $url = Get-BingImageSearchUrl -Query $query -TransparentOnly:$chkTransparent.Checked
        $page = $script:Http.GetStringAsync($url)
        Wait-Tasks @($page) 25000
        if ($script:State.Closing) { return }
        if ($page.Status -ne [System.Threading.Tasks.TaskStatus]::RanToCompletion) {
            $reason = if ($page.IsFaulted) { $page.Exception.GetBaseException().Message } else { 'timed out' }
            Set-Status "Search failed ($reason). Use Open file..., Paste or From EXE... instead."
            return
        }
        $found = @(ConvertFrom-BingImageHtml -Html $page.Result -Max 20)
        if ($found.Count -eq 0) {
            Set-Status 'No results parsed (Bing may have blocked the request or changed its page). Use Open file..., Paste or From EXE...'
            return
        }

        Set-Status ("Downloading {0} thumbnails..." -f $found.Count)
        $tasks = @(foreach ($r in $found) {
            $t = $null
            if ($r.ThumbUrl) { try { $t = $script:Http.GetByteArrayAsync($r.ThumbUrl) } catch { $t = $null } }
            $t
        })
        Wait-Tasks $tasks 10000
        if ($script:State.Closing) { return }

        $lv.BeginUpdate()
        try {
            $lv.Items.Clear()
            $imageList.Images.Clear()
            for ($i = 0; $i -lt $found.Count; $i++) {
                $bytes = Get-TaskBytes $tasks[$i]
                $thumb = ConvertFrom-ImageBytes $bytes
                if (-not $thumb) { continue }
                $found[$i] | Add-Member -NotePropertyName ThumbBytes -NotePropertyValue $bytes -Force
                $tileSize = $imageList.ImageSize.Width
                $tile = New-Object System.Drawing.Bitmap $tileSize, $tileSize
                $tg = [System.Drawing.Graphics]::FromImage($tile)
                $tg.Clear([System.Drawing.Color]::FromArgb(240, 240, 240))
                $tg.InterpolationMode = 'HighQualityBicubic'
                $ratio = [Math]::Min(($tileSize - 4) / $thumb.Width, ($tileSize - 4) / $thumb.Height)
                $tw = [int][Math]::Max(1, $thumb.Width * $ratio); $th = [int][Math]::Max(1, $thumb.Height * $ratio)
                $tg.DrawImage($thumb, [int](($tileSize - $tw) / 2), [int](($tileSize - $th) / 2), $tw, $th)
                $tg.Dispose(); $thumb.Dispose()
                $imageList.Images.Add($tile) | Out-Null
                $tile.Dispose()
                $label = if ($found[$i].Width) { "{0}x{1}" -f $found[$i].Width, $found[$i].Height } else { '' }
                $item = New-Object System.Windows.Forms.ListViewItem ($label, ($imageList.Images.Count - 1))
                $item.Tag = $found[$i]
                $item.ToolTipText = "$($found[$i].Title)`n$($found[$i].FullUrl)"
                [void]$lv.Items.Add($item)
            }
        } finally {
            $lv.EndUpdate()
        }
        if ($lv.Items.Count -eq 0) { Set-Status 'Results were found but no thumbnails could be downloaded.' }
        else { Set-Status ("{0} results. Select one to load it." -f $lv.Items.Count) }
        if (-not $script:State.NameTouched -and -not $script:State.Recipe) { $txtFile.Text = ConvertTo-IconFileName $text }
    } catch {
        Set-Status "Search error: $($_.Exception.Message)"
    } finally {
        # A result clicked during the search belonged to the old list, which is gone now.
        $script:State.Pending = $null
        if (-not $script:State.Closing) { $form.UseWaitCursor = $false }
        $script:State.Busy = $false
    }
}

function Get-CandidateBitmap {
    # Full image, then Bing's thumbnail at 512px, then the small list thumbnail.
    param($Candidate)
    if ($Candidate.FullUrl -notmatch '\.svg(\?|$)') {
        Set-Status 'Downloading full-size image...'
        $bmp = ConvertFrom-ImageBytes (Get-UrlBytes $Candidate.FullUrl 12000)
        if ($bmp) { return [pscustomobject]@{ Bitmap = $bmp; Source = 'full image' } }
    }
    if ($script:State.Closing) { return $null }
    if ($Candidate.ThumbUrl) {
        Set-Status 'Full image unavailable; downloading a 512px copy from Bing...'
        $bmp = ConvertFrom-ImageBytes (Get-UrlBytes (Get-BingThumbnailUrl -ThumbUrl $Candidate.ThumbUrl -Size 512) 10000)
        if ($bmp) { return [pscustomobject]@{ Bitmap = $bmp; Source = 'Bing 512px copy (full image unavailable)' } }
    }
    if ($Candidate.PSObject.Properties['ThumbBytes']) {
        $bmp = ConvertFrom-ImageBytes $Candidate.ThumbBytes
        if ($bmp) { return [pscustomobject]@{ Bitmap = $bmp; Source = 'small thumbnail (full image unavailable)' } }
    }
    $null
}

function Import-Candidate {
    param($Candidate)
    $script:State.LoadGen++
    if ($script:State.Busy) {
        # Loaded once the current download finishes; anything it returns is discarded.
        $script:State.Pending = $Candidate
        return
    }
    $script:State.Busy = $true
    $form.UseWaitCursor = $true
    try {
        $next = $Candidate
        while ($next -and -not $script:State.Closing) {
            $script:State.Pending = $null
            $gen = $script:State.LoadGen
            try {
                $loaded = Get-CandidateBitmap $next
                if ($script:State.Closing) { if ($loaded) { $loaded.Bitmap.Dispose() }; break }
                if ($gen -ne $script:State.LoadGen) {
                    if ($loaded) { $loaded.Bitmap.Dispose() }
                } elseif ($loaded) {
                    Initialize-Image -Bitmap $loaded.Bitmap -Source $loaded.Source
                } else {
                    Set-Status 'Could not load that image. Try another result.'
                }
            } catch {
                Set-Status "Load error: $($_.Exception.Message)"
            }
            $next = $script:State.Pending
        }
    } finally {
        if (-not $script:State.Closing) { $form.UseWaitCursor = $false }
        $script:State.Busy = $false
    }
}

function Import-File {
    param([string]$Path)
    try {
        $bmp = ConvertFrom-ImageBytes ([System.IO.File]::ReadAllBytes($Path))
        if (-not $bmp) { Set-Status "Could not decode '$Path' as an image."; return }
        Initialize-Image -Bitmap $bmp -Source ([System.IO.Path]::GetFileName($Path))
    } catch {
        Set-Status "Open error: $($_.Exception.Message)"
    }
}

function Import-Exe {
    param([string]$Path)
    try {
        $bmp = Get-ExeIconBitmap -Path $Path
        if (-not $bmp) { Set-Status "No icon found in '$([System.IO.Path]::GetFileName($Path))'."; return }
        Initialize-Image -Bitmap $bmp -Source ("icon from {0}" -f [System.IO.Path]::GetFileName($Path))
    } catch {
        Set-Status "Icon extraction error: $($_.Exception.Message)"
    }
}

function Invoke-Edit {
    param([scriptblock]$Action, [string]$Message)
    if (-not $script:State.Current) { return }
    try {
        $result = & $Action
        Set-CurrentImage -Bitmap $result
        Set-Status ("{0} ({1}x{2})" -f $Message, $result.Width, $result.Height)
    } catch {
        Set-Status "Edit failed: $($_.Exception.Message)"
    }
}

function Save-Icon {
    if (-not $script:State.Current) { return }
    $fileName = ($txtFile.Text.Trim() -replace '[\\/:*?"<>|]', '')
    if (-not $fileName) { Set-Status 'Enter a file name first.'; return }
    if ($fileName -notmatch '\.png$') { $fileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + '.png' }
    $path = Join-Path $OutputFolder $fileName

    $bmp = $script:State.Current
    $toSave = $bmp
    if ($bmp.Width -gt $Size -or $bmp.Height -gt $Size) {
        # Intune rejects larger icons, so there is no "save as is" option.
        $answer = [System.Windows.Forms.MessageBox]::Show(
            ("The image is {0}x{1}, larger than the {2}px Intune limit. It will be shrunk to fit when saved." -f $bmp.Width, $bmp.Height, $Size),
            'Image too large', 'OKCancel', 'Information')
        if ($answer -ne 'OK') { return }
        $toSave = [YardstickIconOps]::Resize($bmp, $Size, $false)
    }
    try {
        if (Test-Path -LiteralPath $path) {
            $overwrite = [System.Windows.Forms.MessageBox]::Show("'$fileName' already exists in the icons folder. Overwrite it?", 'Overwrite icon', 'YesNo', 'Question')
            if ($overwrite -ne 'Yes') { return }
        }
        if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
        $toSave.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)

        $recipe = $script:State.Recipe
        $hint = if ($recipe -and $recipe.IconFile -and $recipe.FileName -eq $fileName) { "Recipe '$($recipe.Id)' already points at it." }
                else { "Recipe field:  iconFile: $fileName" }
        Update-MissingList
        Set-Status ("Saved {0} ({1}x{2}). {3}" -f $path, $toSave.Width, $toSave.Height, $hint)
    } catch {
        Set-Status "Save failed: $($_.Exception.Message)"
    } finally {
        if ($toSave -ne $bmp) { $toSave.Dispose() }
    }
}
#endregion

#region Wire up events
$btnSearch.Add_Click({ Invoke-Search })
$txtSearch.Add_KeyDown({ if ($_.KeyCode -eq 'Return') { $_.SuppressKeyPress = $true; Invoke-Search } })
$txtFile.Add_TextChanged({
    if ($txtFile.Focused) {
        $script:State.NameTouched = $true
        $script:State.Recipe = $null
    }
})
$lv.Add_SelectedIndexChanged({
    if ($lv.SelectedItems.Count -gt 0) { Import-Candidate $lv.SelectedItems[0].Tag }
})
$cmbMissing.Add_SelectedIndexChanged({
    if ($script:State.Rebuilding -or $cmbMissing.SelectedIndex -lt 1) { return }
    $m = $script:State.Missing[$cmbMissing.SelectedIndex - 1]
    $recipe = Find-Recipe -Recipe $script:Recipes -Key $m.Id
    if ($recipe) { Select-Recipe $recipe }
})

$btnOpen.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Images|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.ico;*.webp;*.tif;*.tiff|All files|*.*'
    if ($dlg.ShowDialog($form) -eq 'OK') { Import-File $dlg.FileName }
    $dlg.Dispose()
})
$btnPaste.Add_Click({
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $img = [System.Windows.Forms.Clipboard]::GetImage()
            try { Initialize-Image -Bitmap ([YardstickIconOps]::ToArgb($img)) -Source 'clipboard image' } finally { $img.Dispose() }
        } elseif ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
            Import-File ([System.Windows.Forms.Clipboard]::GetFileDropList()[0])
        } else {
            Set-Status 'The clipboard does not contain an image.'
        }
    } catch { Set-Status "Paste error: $($_.Exception.Message)" }
})
$btnExe.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Programs and libraries|*.exe;*.dll|All files|*.*'
    if ($dlg.ShowDialog($form) -eq 'OK') { Import-Exe $dlg.FileName }
    $dlg.Dispose()
})

$btnCrop.Add_Click({
    $sel = $script:State.Selection
    if ($sel.IsEmpty) { return }
    Invoke-Edit { [YardstickIconOps]::Crop($script:State.Current, $sel) } 'Cropped'
})
$btnTrim.Add_Click({
    Invoke-Edit {
        $r = [YardstickIconOps]::TrimBounds($script:State.Current, [Math]::Max(8, $trkTol.Value))
        [YardstickIconOps]::Crop($script:State.Current, $r)
    } 'Trimmed borders'
})
$btnResize.Add_Click({
    $px = 0
    if (-not [int]::TryParse($cmbSize.Text, [ref]$px) -or $px -lt 8 -or $px -gt 4096) { Set-Status 'Enter a size between 8 and 4096.'; return }
    Invoke-Edit { [YardstickIconOps]::Resize($script:State.Current, $px, $chkSquare.Checked) } "Resized to $px"
})

$btnPick.Add_CheckedChanged({
    $script:State.Picking = $btnPick.Checked
    $canvas.Cursor = if ($btnPick.Checked) { [System.Windows.Forms.Cursors]::Cross } else { [System.Windows.Forms.Cursors]::Default }
    if ($btnPick.Checked) { Set-Status 'Click a background colour on the image to make it transparent.' }
})
$trkTol.Add_ValueChanged({ $lblTol.Text = "Tolerance: $($trkTol.Value)" })
$btnWhite.Add_Click({
    $tol = [Math]::Max(10, $trkTol.Value)
    Invoke-Edit { [YardstickIconOps]::RemoveWhite($script:State.Current, $tol) } 'Removed white'
})
$btnColour.Add_Click({
    $dlg = New-Object System.Windows.Forms.ColorDialog
    $dlg.Color = $script:State.BgColour
    $dlg.FullOpen = $true
    if ($dlg.ShowDialog($form) -eq 'OK') {
        $script:State.BgColour = $dlg.Color
        $btnColour.BackColor = $dlg.Color
    }
    $dlg.Dispose()
})
$btnApplyBg.Add_Click({
    Invoke-Edit { [YardstickIconOps]::AddBackground($script:State.Current, $script:State.BgColour, $chkRounded.Checked) } 'Added background'
})

$btnUndo.Add_Click({
    if ($script:State.Undo.Count -eq 0) { return }
    $last = $script:State.Undo[$script:State.Undo.Count - 1]
    $script:State.Undo.RemoveAt($script:State.Undo.Count - 1)
    if ($script:State.Current) { $script:State.Current.Dispose() }
    $script:State.Current = $last
    $script:State.Selection = [System.Drawing.Rectangle]::Empty
    $canvas.Invalidate()
    Update-State
    Set-Status ("Undone ({0}x{1})" -f $last.Width, $last.Height)
})
$btnReset.Add_Click({
    if (-not $script:State.Original) { return }
    $copy = $script:State.Original.Clone()
    Set-CurrentImage -Bitmap $copy
    Set-Status 'Reset to the original image.'
})
$btnSave.Add_Click({ Save-Icon })

$toolTip.SetToolTip($cmbMissing, 'Recipes whose iconFile is not in the icons folder. Pick one to search for it and pre-fill the file name.')
$toolTip.SetToolTip($btnExe, 'Extracts the largest icon embedded in an installed program (.exe or .dll).')
$toolTip.SetToolTip($btnTrim, 'Removes transparent (or uniform-colour) borders. Uses the tolerance slider, minimum 8.')
$toolTip.SetToolTip($btnWhite, 'Makes white pixels (tolerance from the slider, minimum 10) transparent, softening anti-aliased edges.')
$toolTip.SetToolTip($trkTol, 'Colours within this distance are removed; up to twice this distance they fade out (soft edges).')
$toolTip.SetToolTip($btnResize, 'Scales so the longest side matches; optionally pads to a square canvas.')

# Closing while a download is running: stop the requests and let the waiting handler
# unwind. Bitmaps and the HTTP client are disposed only after ShowDialog returns,
# when no handler can still be using them.
$form.Add_FormClosing({
    $script:State.Closing = $true
    try { $script:Http.CancelPendingRequests() } catch { Write-Verbose "Cancel failed: $($_.Exception.Message)" }
})
$form.Add_Shown({
    Update-MissingList
    Update-State
    if ($txtSearch.Text.Trim()) { Invoke-Search }
})
#endregion

try {
    [void]$form.ShowDialog()
} finally {
    $script:State.Closing = $true
    foreach ($b in $script:State.Undo) { $b.Dispose() }
    if ($script:State.Current) { $script:State.Current.Dispose() }
    if ($script:State.Original) { $script:State.Original.Dispose() }
    $checkerBrush.Dispose()
    $checker.Dispose()
    $form.Dispose()
    $script:Http.Dispose()
}
