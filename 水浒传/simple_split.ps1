# Simple Water Margin Chapter Splitter
# Author: AI Assistant

# Set file paths
$sourceFile = "d:\AI\books\水浒传\水浒全传.txt"
$outputDir = "d:\AI\books\水浒传\章回"

# Create output directory if it doesn't exist
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force
}

# Read source file content
$content = Get-Content $sourceFile -Raw -Encoding UTF8

# Define regex pattern for chapters
$pattern = "第.{0,5}?回"

# Find all chapter titles
$matches = [regex]::Matches($content, $pattern)

Write-Host "Found $($matches.Count) chapters"

# Split and save chapters
for ($i = 0; $i -lt $matches.Count; $i++) {
    $start = $matches[$i].Index
    $title = $matches[$i].Value
    
    # Determine end position
    if ($i -lt $matches.Count - 1) {
        $end = $matches[$i + 1].Index
        $chapterText = $content.Substring($start, $end - $start).Trim()
    } else {
        $chapterText = $content.Substring($start).Trim()
    }
    
    # Generate filename
    $fileName = "{0:D3}.txt" -f ($i + 1)
    $filePath = Join-Path $outputDir $fileName
    
    # Save chapter
    $chapterText | Out-File -FilePath $filePath -Encoding UTF8 -Force
    
    Write-Host "Created: $fileName - $title"
}

Write-Host "Split complete!"