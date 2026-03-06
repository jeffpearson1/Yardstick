$ErrorActionPreference = 'Stop'

# Load support module
Import-Module "$PSScriptRoot\..\Modules\YardstickSupport.psm1" -Force

$recipesPath = "$PSScriptRoot\..\Recipes"
$failed = 0
$passed = 0

Get-ChildItem $recipesPath -Recurse -Include *.yaml, *.yml | ForEach-Object {
    $content = Get-Content $_.FullName | ConvertFrom-Yaml
    $id = $_.BaseName

    # Resolve base recipe inheritance before validation
    if ($content.ContainsKey('base')) {
        try {
            $content = Merge-RecipeWithBase -Recipe $content -RecipesPath $recipesPath
        } catch {
            Write-Host "ERROR: $($_.Name): Failed to resolve base recipe: $_" -ForegroundColor Red
            $script:failed++
            return
        }
    }

    $result = Test-RecipeSchema -Recipe $content -RecipeId $id
    if (-not $result.IsValid) {
        foreach ($e in $result.Errors) {
            Write-Host "ERROR: $($_.Name): $e" -ForegroundColor Red
        }
        $script:failed++
    } else {
        $script:passed++
    }
    foreach ($w in $result.Warnings) {
        Write-Host "WARNING: $($_.Name): $w" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "$passed recipe(s) passed, $failed recipe(s) failed schema validation."

if ($failed -gt 0) {
    exit 1
}
