# GitHub Repository Auto-Commit Script
# This script finds git repositories, analyzes changes, and commits with realistic messages

# Function to generate a commit message based on the changes
function Get-CommitMessage {
    param (
        [string]$repoPath
    )
    
    # Get status information
    $status = git -C $repoPath status --porcelain
    
    # If no changes, return empty
    if ([string]::IsNullOrEmpty($status)) {
        return $null
    }
    
    # Count types of changes
    $addedFiles = ($status | Where-Object { $_ -match '^\?\?' }).Count
    $modifiedFiles = ($status | Where-Object { $_ -match '^.M' }).Count
    $deletedFiles = ($status | Where-Object { $_ -match '^.D' }).Count
    
    # Get file extensions to identify primary language/type of changes
    $changedFiles = $status | ForEach-Object {
        if ($_.Length -gt 3) {
            $_.Substring(3)
        } else {
            ""
        }
    }
    
    $fileExtensions = @{}
    foreach ($file in $changedFiles) {
        try {
            $ext = [System.IO.Path]::GetExtension($file)
            if ($ext) {
                if ($fileExtensions.ContainsKey($ext)) {
                    $fileExtensions[$ext]++
                } else {
                    $fileExtensions[$ext] = 1
                }
            }
        } catch {
            # Skip files with problematic paths
            Write-Host "Skipping extension extraction for file: $file" -ForegroundColor Yellow
        }
    }
    
    # Determine primary file type based on extensions or common patterns
    $primaryType = "files"
    
    # Check for Python files (common in your repository)
    $pythonCount = 0
    $pythonCount += ($changedFiles | Where-Object { $_ -match '\.py$|\.pyw$' }).Count
    
    if ($pythonCount -gt 0) {
        $primaryType = "Python"
    }
    
    # Get the most common extension if we have valid extensions
    if ($fileExtensions.Count -gt 0) {
        $primaryExtension = $fileExtensions.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First 1
        
        # Create a realistic message based on the changes
        $extensionMap = @{
            ".js" = "JavaScript"
            ".ts" = "TypeScript"
            ".jsx" = "React"
            ".tsx" = "React TypeScript"
            ".py" = "Python"
            ".html" = "HTML"
            ".css" = "CSS"
            ".scss" = "SCSS"
            ".c" = "C"
            ".cpp" = "C++"
            ".cs" = "C#"
            ".php" = "PHP"
            ".go" = "Go"
            ".java" = "Java"
            ".rb" = "Ruby"
            ".md" = "documentation"
            ".json" = "JSON configuration"
            ".yml" = "YAML configuration"
            ".yaml" = "YAML configuration"
            ".xml" = "XML"
            ".sh" = "shell script"
            ".ps1" = "PowerShell script"
            ".cmd" = "batch script"
            ".bat" = "batch script"
            ".txt" = "text files"
        }
        
        if ($primaryExtension.Name -and $extensionMap.ContainsKey($primaryExtension.Name)) {
            $primaryType = $extensionMap[$primaryExtension.Name]
        }
    }
    
    # Determine the primary type of change
    $changeType = if ($addedFiles -gt $modifiedFiles -and $addedFiles -gt $deletedFiles) {
        "Add"
    } elseif ($deletedFiles -gt $modifiedFiles -and $deletedFiles -gt $addedFiles) {
        "Remove"
    } else {
        "Update"
    }
    
    # Generate appropriate verb forms
    $verb = switch ($changeType) {
        "Add" { "Add"; break }
        "Remove" { "Remove"; break }
        "Update" { "Update"; break }
    }
    
    # Generate action-specific descriptions
    if ($changeType -eq "Add" -and $addedFiles -eq 1) {
        # For single file additions
        try {
            $addedFile = ($status | Where-Object { $_ -match '^\?\?' } | Select-Object -First 1)
            if ($addedFile.Length -gt 3) {
                $addedFile = $addedFile.Substring(3)
                return "Add $addedFile"
            }
        } catch {
            # Fall back to generic message
        }
    } elseif ($changeType -eq "Update" -and $modifiedFiles -eq 1) {
        # For single file modifications
        try {
            $modifiedFile = ($status | Where-Object { $_ -match '^.M' } | Select-Object -First 1)
            if ($modifiedFile.Length -gt 3) {
                $modifiedFile = $modifiedFile.Substring(3)
                return "Update $modifiedFile"
            }
        } catch {
            # Fall back to generic message
        }
    }
    
    # For multiple changes
    $totalChanges = $addedFiles + $modifiedFiles + $deletedFiles
    if ($totalChanges -eq 1) {
        return "$verb $primaryType"
    } else {
        return "$verb multiple $primaryType ($addedFiles added, $modifiedFiles modified, $deletedFiles deleted)"
    }
}

# Function to process a git repository
function Process-GitRepo {
    param (
        [string]$repoPath
    )
    
    Write-Host "`n========================================="
    Write-Host "Processing repository at: $repoPath" -ForegroundColor Cyan
    
    # Check if repository is initialized
    if (-not (Test-Path (Join-Path $repoPath ".git"))) {
        Write-Host "Not a git repository. Skipping." -ForegroundColor Yellow
        return
    }
    
    # Get repository status
    Set-Location $repoPath
    $status = git status --porcelain
    
    if ([string]::IsNullOrEmpty($status)) {
        Write-Host "No changes to commit." -ForegroundColor Green
        return
    }
    
    # Show status
    Write-Host "Changes detected:" -ForegroundColor Yellow
    git status --short
    
    # Generate commit message
    $commitMessage = Get-CommitMessage -repoPath $repoPath
    
    if ($null -eq $commitMessage) {
        Write-Host "No changes to commit." -ForegroundColor Green
        return
    }
    
    # Confirm with user
    Write-Host "`nProposed commit message: '$commitMessage'" -ForegroundColor Magenta
    $confirm = Read-Host "Proceed with commit and push? (y/n)"
    
    if ($confirm -eq 'y') {
        # Stage changes
        git add .
        
        # Commit changes
        git commit -m $commitMessage
        
        # Check if remote exists and push
        $remoteExists = git remote -v
        if ($remoteExists) {
            $currentBranch = git rev-parse --abbrev-ref HEAD
            git push origin $currentBranch
            Write-Host "Changes pushed successfully to branch '$currentBranch'." -ForegroundColor Green
        } else {
            Write-Host "No remote repository configured. Commit created locally only." -ForegroundColor Yellow
        }
    } else {
        Write-Host "Skipped committing changes." -ForegroundColor Yellow
    }
}

# Main script execution
Write-Host "GitHub Repository Auto-Commit Tool" -ForegroundColor Cyan
Write-Host "===================================`n" -ForegroundColor Cyan

# Ask for the directory to scan for git repositories
$scanPath = Read-Host "Enter path to scan for git repositories (leave empty for current directory)"

if ([string]::IsNullOrEmpty($scanPath)) {
    $scanPath = Get-Location
}

# Find all git repositories in the specified path
Write-Host "Scanning for git repositories in $scanPath ..." -ForegroundColor Cyan
$gitRepos = Get-ChildItem -Path $scanPath -Recurse -Force -Directory -ErrorAction SilentlyContinue | 
            Where-Object { $_.Name -eq ".git" } | 
            ForEach-Object { $_.Parent.FullName }

if ($gitRepos.Count -eq 0) {
    Write-Host "No git repositories found in the specified path." -ForegroundColor Yellow
    exit
}

Write-Host "Found $($gitRepos.Count) git repositories." -ForegroundColor Green

# Process each repository
foreach ($repo in $gitRepos) {
    Process-GitRepo -repoPath $repo
}

Write-Host "`nAll repositories processed." -ForegroundColor Cyan
