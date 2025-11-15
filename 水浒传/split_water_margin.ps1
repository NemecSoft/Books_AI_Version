# Water Margin Chapter Splitter Script
# Author: AI Assistant
# Function: Split Water Margin text file by chapters using regex

# Set file paths
$sourceFilePath = "d:\AI\books\水浒传\水浒全传.txt"
$regexFilePath = "d:\AI\books\水浒传\正则.txt"
$outputDirectory = "d:\AI\books\水浒传\章回"

# Read regex pattern
$regexPattern = Get-Content $regexFilePath -Raw

# Create output directory if it doesn't exist
if (-not (Test-Path $outputDirectory)) {
    New-Item -ItemType Directory -Path $outputDirectory -Force
    Write-Host "Created output directory: $outputDirectory"
}

# Read source file content
$fileContent = Get-Content $sourceFilePath -Raw -Encoding UTF8

Write-Host "Starting to split Water Margin..."
Write-Host "Using regex pattern: $regexPattern"

# Find all chapter titles using regex
$chapterMatches = [regex]::Matches($fileContent, $regexPattern)

Write-Host "Found $($chapterMatches.Count) chapter titles"

# If no chapters found, try other possible formats
if ($chapterMatches.Count -eq 0) {
    Write-Host "No chapter titles found, trying other formats..."
    # Try other possible chapter formats
    $backupRegex = "第[一二三四五六七八九十百零\d]+回"
    $chapterMatches = [regex]::Matches($fileContent, $backupRegex)
    Write-Host "Found $($chapterMatches.Count) chapter titles using backup regex"
}

# Split file
for ($i = 0; $i -lt $chapterMatches.Count; $i++) {
    $currentChapter = $chapterMatches[$i]
    $currentTitle = $currentChapter.Value
    $currentPosition = $currentChapter.Index
    
    # Determine chapter end position (next chapter start or end of file)
    if ($i -lt $chapterMatches.Count - 1) {
        $nextPosition = $chapterMatches[$i + 1].Index
        $chapterContent = $fileContent.Substring($currentPosition, $nextPosition - $currentPosition).Trim()
    } else {
        $chapterContent = $fileContent.Substring($currentPosition).Trim()
    }
    
    # Generate filename (extract chapter number)
    $chapterNumber = [regex]::Match($currentTitle, "第(.{0,5}?)回").Groups[1].Value
    if (-not $chapterNumber) {
        $chapterNumber = "{0:D3}" -f ($i + 1)
    }
    
    # Ensure filename is 3-digit format
    if ($chapterNumber -match "^\d+$") {
        $fileName = "{0:D3}.txt" -f [int]$chapterNumber
    } else {
        $fileName = "{0:D3}.txt" -f ($i + 1)
    }
    
    $outputFilePath = Join-Path $outputDirectory $fileName
    
    # Write chapter content
    $chapterContent | Out-File -FilePath $outputFilePath -Encoding UTF8 -Force
    
    Write-Host "Generated: $fileName - $currentTitle"
}

Write-Host "Split completed! Generated $($chapterMatches.Count) chapter files."
Write-Host "Files saved in: $outputDirectory"